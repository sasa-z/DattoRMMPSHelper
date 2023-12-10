function Install-AdobeReader{



#endregion app variables to change ###################################################################################################################
$SoftwareName = "Adobe Acrobat (64-bit)" #name visible via get-package PS command
$ChocolateyName = "adobereader" #name used in chocolatey to install software

$processNames = @('acrobat')
#endregion variables to change ###################################################################################################################


$preSoftwareCheck = check-softwarePresence -softwareName $SoftwareName -includeChoco -chocolateyName $ChocolateyName

foreach ($item in $processNames){

    if( get-process $item -ErrorAction SilentlyContinue){
        $AppRunning = $true
    }

}

#region stop app process
if ($AppRunning){

    send-Log -logText "$($Softwarename) is running, Closing now." 

    foreach ($process in $processNames){
        get-process $process -ErrorAction silentlycontinue | stop-process -force -ErrorAction silentlycontinue
        
    }

}
#endregion stop app process

    if ($preSoftwareCheck.InstalledviaChoco -and -not $preSoftwareCheck.UpdatedNeededViaChoco){ 

        send-log -logText "Latest version of $($Softwarename) is already installed ($($PreVersionCheck)). `nExiting script." -addDashes Below -addTeamsMessage
        send-CustomToastNofication -text "Latest version already installed." 

    }else{
        
        send-Log -logText "Starting $($Softwarename) installation."
        send-CustomToastNofication -text "Starting installation."
    
        try{
            start-process choco -ArgumentList "install $($ChocolateyName) -y --force" -ErrorAction Stop -Wait  | Out-Null

            $postCheckInstallation = check-softwarePresence -softwareName $SoftwareName -includeChoco -chocolateyName $ChocolateyName
    
            if ($postCheckInstallation.InstalledviaChoco -and -not $postCheckInstallation.UpdatedNeededViaChoco){
    
                send-log -logText "Successfully installed $($Softwarename) version: $($postcheckInstallation.versionviaChoco)" -addDashes Below -addTeamsMessage
                send-CustomToastNofication -text "Successfully installed" 
    
            }else{

                send-log -logText "Failed to install $($Softwarename)" -addDashes Below -type warning  -addTeamsMessage
                send-CustomToastNofication  -text "Failed to install $($Softwarename)" -type warning
        
            }
            
        }catch{

            send-log -logText "Failed to install $($Softwarename)" -addDashes Below -type warning -addTeamsMessage -catch
            send-CustomToastNofication  -text "Failed to install" -type warning
        }

    }

}
