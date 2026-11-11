$WshShell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath('Desktop')
$ShortcutPath = Join-Path $Desktop "Compliance Automation Dashboard.lnk"

$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "d:\Complience\start_dashboard.bat"
$Shortcut.WorkingDirectory = "d:\Complience"
$Shortcut.WindowStyle = 1 # Normal window
$Shortcut.Description = "Compliance Automation Suite (ESIC & EPFO)"
# Icon from shell32.dll (index 13 is network/globe, 220 is shield, 221 is key/lock, or imageres.dll)
$Shortcut.IconLocation = "shell32.dll,13"
$Shortcut.Save()

Write-Output "Shortcut created at: $ShortcutPath"
