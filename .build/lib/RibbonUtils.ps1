# Script:   RibbonUtils.ps1
# Purpose:  Ribbon XML schema validation and zip injection helpers for the build pipeline.
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-RibbonValidity {
    param (
        [string]$XmlPath,
        [string]$ModulesDir
    )

    if (-not (Test-Path $XmlPath)) { return $true }

    Write-Host "Validating Ribbon XML..." -ForegroundColor Cyan
    $isValid = $true
    $absoluteXmlPath = Resolve-Path $XmlPath

    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.XmlResolver = $null
        $settings.ValidationType = [System.Xml.ValidationType]::Schema
        $settings.ValidationFlags = $settings.ValidationFlags -bor [System.Xml.Schema.XmlSchemaValidationFlags]::ProcessIdentityConstraints
        $settings.ValidationFlags = $settings.ValidationFlags -bor [System.Xml.Schema.XmlSchemaValidationFlags]::ReportValidationWarnings

        $onValidationError = [System.Xml.Schema.ValidationEventHandler] {
            param($evtSource, $e)
            if ($e.Message -match "Could not find schema information") { return }
            
            $script:isValid = $false
            $line = $e.Exception.LineNumber
            $col = $e.Exception.LinePosition
            Write-Host "  Ribbon XML Error [Line $line, Col $col]: $($e.Message)" -ForegroundColor Red
        }
        $settings.add_ValidationEventHandler($onValidationError)

        $reader = [System.Xml.XmlReader]::Create($absoluteXmlPath, $settings)
        while ($reader.Read()) { }
        $reader.Close()
    } catch {
        Write-Error "Ribbon XML failed to load or is malformed: $($_.Exception.Message)"
        $isValid = $false
    }

    if (-not $isValid) { return $false }

    $xml = [xml](Get-Content $XmlPath -Raw)
    
    $ids = $xml.SelectNodes("//@id") | ForEach-Object { $_.Value }
    $duplicates = $ids | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($duplicates) {
        Write-Error "Duplicate IDs found in ribbon.xml: $($duplicates.Name -join ', ')"
        $isValid = $false
    }

    $callbacks = $xml.SelectNodes("//@onAction") | ForEach-Object { $_.Value } | Select-Object -Unique
    if ($callbacks) {
        Write-Host "  Checking $($callbacks.Count) callbacks across all modules..."
        $vbaFiles = Get-ChildItem -Path $ModulesDir -Include *.bas, *.cls -Recurse
        $sb = New-Object System.Text.StringBuilder
        foreach ($f in $vbaFiles) {
            [void]$sb.AppendLine([System.IO.File]::ReadAllText($f.FullName))
        }
        $vbaCode = $sb.ToString()
        
        foreach ($cb in $callbacks) {
            if ($vbaCode -notmatch "Sub\s+$cb\s*\(") {
                Write-Error "Ribbon callback '$cb' not found in any module in $ModulesDir"
                $isValid = $false
            }
        }
    }

    return $isValid
}

function Update-RibbonInWorkbook {
    param ([string]$WorkbookPath, [string]$RibbonXmlPath)
    if (-not (Test-Path $RibbonXmlPath)) { return }
    Write-Host "Injecting Ribbon XML..."
    $zip = $null
    try {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
        
        $ribbonEntryPath = "customUI/customUI14.xml"
        $ribbonEntry = $zip.GetEntry($ribbonEntryPath)
        if ($null -ne $ribbonEntry) { $ribbonEntry.Delete() }
        $ribbonEntry = $zip.CreateEntry($ribbonEntryPath)
        $writer = New-Object System.IO.StreamWriter($ribbonEntry.Open())
        $writer.Write((Get-Content $RibbonXmlPath -Raw))
        $writer.Close()

        $relsEntryPath = "_rels/.rels"
        $relsEntry = $zip.GetEntry($relsEntryPath)
        if ($null -eq $relsEntry) {
            throw "_rels/.rels not found in workbook. Invalid Excel file structure."
        }

        $relsXml = [xml]""
        $stream = $relsEntry.Open()
        try {
            $reader = New-Object System.IO.StreamReader($stream)
            $relsXml = [xml]$reader.ReadToEnd()
        } finally {
            $stream.Close()
        }

        $nsMgr = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
        
        $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        $existingRel = $relsXml.SelectSingleNode("//r:Relationship[@Target='$ribbonEntryPath']", $nsMgr)
        
        if ($null -eq $existingRel) {
            Write-Host "  Adding Custom UI relationship to _rels/.rels..."
            $ids = $relsXml.SelectNodes("//r:Relationship/@Id", $nsMgr) | ForEach-Object { 
                if ($_.Value -match "rId(\d+)") { [int]$matches[1] } else { 0 }
            }
            $nextId = ($ids | Measure-Object -Maximum).Maximum + 1
            $newId = "rId$nextId"
            
            $root = $relsXml.DocumentElement
            $newRel = $relsXml.CreateElement("Relationship", $relsXml.DocumentElement.NamespaceURI)
            $newRel.SetAttribute("Id", $newId)
            $newRel.SetAttribute("Type", $relType)
            $newRel.SetAttribute("Target", $ribbonEntryPath)
            $root.AppendChild($newRel) | Out-Null
            
            $relsEntry.Delete()
            $relsEntry = $zip.CreateEntry($relsEntryPath)
            $writer = New-Object System.IO.StreamWriter($relsEntry.Open())
            $relsXml.Save($writer)
            $writer.Close()
        }

        $zip.Dispose()
        Write-Host "  Ribbon XML injected and registered successfully."
    } catch {
        throw "Failed to update Ribbon XML: $($_.Exception.Message)"
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}
