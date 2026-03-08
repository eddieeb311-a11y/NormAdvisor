try {
    $dllPath = "C:\Users\Byamba Erdene\Desktop\Norm\NormAdvisorPro\NormAdvisor.AutoCAD1\bin\Debug\NormAdvisor.AutoCAD1.dll"
    $asm = [System.Reflection.Assembly]::LoadFrom($dllPath)
    Write-Host "Assembly loaded OK: $($asm.FullName)"
    Write-Host ""
    Write-Host "Referenced assemblies:"
    foreach ($ref in $asm.GetReferencedAssemblies()) {
        Write-Host "  - $($ref.Name) v$($ref.Version)"
    }
    Write-Host ""
    Write-Host "Types:"
    foreach ($t in $asm.GetExportedTypes()) {
        Write-Host "  - $($t.FullName)"
    }
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    if ($_.Exception.InnerException) {
        Write-Host "INNER: $($_.Exception.InnerException.Message)"
    }
}
