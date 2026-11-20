# Removes the Mermaid Diagrams add-in installed by install-windows.ps1.
$addinId = "__ADDIN_ID__"
$registryKey = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
Remove-ItemProperty -Path $registryKey -Name $addinId -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force (Join-Path $env:LOCALAPPDATA "Mermaid2PowerPoint") -ErrorAction SilentlyContinue
Write-Host "Mermaid Diagrams add-in removed. Restart PowerPoint to finish."
