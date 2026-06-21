$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $true
    $wb = $excel.Workbooks.Add()
    $vbe = $excel.VBE
    $menuBar = $vbe.CommandBars.Item("Menu Bar")
    
    $missing = [System.Reflection.Missing]::Value
    
    Write-Host "Trying 2 arguments (Type, Id)..."
    try {
        $btn = $menuBar.FindControl(1, 578)
        Write-Host "  Success: $($btn.Caption)"
    } catch {
        Write-Host "  Failed: $($_.Exception.Message)"
    }
    
    Write-Host "Trying 5 arguments with 1 as Type..."
    try {
        $btn = $menuBar.FindControl(1, 578, $missing, $missing, $true)
        Write-Host "  Success: $($btn.Caption)"
    } catch {
        Write-Host "  Failed: $($_.Exception.Message)"
    }

    Write-Host "Trying 5 arguments with [System.Reflection.Missing]::Value as Type..."
    try {
        $btn = $menuBar.FindControl($missing, 578, $missing, $missing, $true)
        Write-Host "  Success: $($btn.Caption)"
    } catch {
        Write-Host "  Failed: $($_.Exception.Message)"
    }
    
    $wb.Close($false)
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
