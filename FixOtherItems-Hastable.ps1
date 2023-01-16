#Run def updates
."C:\Program Files\Windows Defender\MpCmdRun.exe" -removedefinitions -dynamicsignatures
."C:\Program Files\Windows Defender\MpCmdRun.exe" -SignatureUpdate

#Create a Hashtable

$listofprograms = @{

    #value order: TestPath, TargetApp, Program Location, $description

    "Key1" = @("C:\Program Files (x86)\Scomis\Hosted Applications\HostedAppsConnector.exe", "C:\Program Files (x86)\Scomis\Hosted Applications\BootStrap.exe", "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Scomis\Scomis Hosted Applications.lnk", "Scomis Hosted Applications")
    "Key2" = @("Value2a", "Value2b", "Value2c", "Value2d")
    "Key3" = @("Value3a", "Value3b", "Value3c", "Value3d")
    "Key4" = @("Value4a", "Value4b", "Value4c", "Value4d")
}

#Loop though hash table

$listofprograms.GetEnumerator() | ForEach-Object {
    
    $key = $_.Key
    $TestPathVar = $_.Value[0]
    $target = $_.Value[1]
    $shortcut_path = $_.Value[2]
    $description = $_.Value[3]

    #Write-Host "$TestPathVar,$target,$shortcut_path,$description"

    #Scomis
    if ((Test-Path -Path "$TestPathVar")) {

        #Start Menu

        $workingdirectory = (Get-ChildItem $target).DirectoryName
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcut_path)
        $Shortcut.TargetPath = $target
        $Shortcut.Description = $description
        $shortcut.WorkingDirectory = $workingdirectory
        $Shortcut.Save()
        Start-Sleep -Seconds 1

        #Desktop

        $workingdirectory = (Get-ChildItem $target).DirectoryName
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcut_path)
        $Shortcut.TargetPath = $target
        $Shortcut.Description = $description
        $shortcut.WorkingDirectory = $workingdirectory
        $Shortcut.Save()
        Start-Sleep -Seconds 1

        Write-Host "Repaired $description"
    }
    else {
        Write-Host "No $description"
    }
}
