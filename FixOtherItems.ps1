#Run def updates
."C:\Program Files\Windows Defender\MpCmdRun.exe" -removedefinitions -dynamicsignatures
."C:\Program Files\Windows Defender\MpCmdRun.exe" -SignatureUpdate

#Scomis
if ((Test-Path -Path "C:\Program Files (x86)\Scomis\Hosted Applications\HostedAppsConnector.exe")) {

    #start Menu

    $target = "C:\Program Files (x86)\Scomis\Hosted Applications\BootStrap.exe"
    $shortcut_path = "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Scomis\Scomis Hosted Applications.lnk"
    $description = "Scomis Hosted Applications"

    $workingdirectory = (Get-ChildItem $target).DirectoryName
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcut_path)
    $Shortcut.TargetPath = $target
    $Shortcut.Description = $description
    $shortcut.WorkingDirectory = $workingdirectory
    $Shortcut.Save()
    Start-Sleep -Seconds 1

    #Desktop

    $target = "C:\Program Files (x86)\Scomis\Hosted Applications\BootStrap.exe"
    $shortcut_path = "C:\Users\Public\Desktop\Scomis Hosted Applications.lnk"
    $description = "Scomis Hosted Applications"

    $workingdirectory = (Get-ChildItem $target).DirectoryName
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcut_path)
    $Shortcut.TargetPath = $target
    $Shortcut.Description = $description
    $shortcut.WorkingDirectory = $workingdirectory
    $Shortcut.Save()
    Start-Sleep -Seconds 1

}
else {
    Write-Host "No Scomis"
}