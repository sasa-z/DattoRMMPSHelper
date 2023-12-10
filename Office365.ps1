     ##################################################################################################################
     function install-Office365{
        
        
    $SoftwareName = "Microsoft 365 Apps for business - en-us" #name visible via get-package PS command
     $ChocolateyName = "office365business" #name used in chocolatey to install software

     $processNames = @('outlook','excel','powerpnt','winword')
     ###################################################################################################################




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

         if ($preSoftwareCheck.installed){

             send-log -logText "$($Softwarename) is already installed: version $($preSoftwareCheck.version). `nExiting script." -addDashes Below -addTeamsMessage
             send-CustomToastNofication -text "Already installed." 

         }else{

             If ($versionToInstall -eq 'x64'){$Version = '64'}else{ $Version = '32'; $force = "--forcex86"} 

$config = @"
                     <Configuration ID="218c69aa-e349-4443-a4e3-218c577beb80">
                     <Add OfficeClientEdition="$($version)" Channel="Monthly" ForceUpgrade="TRUE">
                         <Product ID="O365BusinessRetail">
                         <Language ID="en-us" />
                         <ExcludeApp ID="Groove" />
                         <ExcludeApp ID="Lync" />
                         <ExcludeApp ID="OneNote" />
                         <ExcludeApp ID="Bing"/>
                         </Product>
                     </Add>
                     <Property Name="SharedComputerLicensing" Value="0" />
                     <Property Name="PinIconsToTaskbar" Value="TRUE" />
                     <Property Name="SCLCacheOverride" Value="0" />
                     <Property Name="AUTOACTIVATE" Value="0" />
                     <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
                     <Property Name="DeviceBasedLicensing" Value="0" />
                     <Updates Enabled="TRUE" />
                     <RemoveMSI>
                         <IgnoreProduct ID="InfoPath" />
                         <IgnoreProduct ID="InfoPathR" />
                         <IgnoreProduct ID="PrjPro" />
                         <IgnoreProduct ID="PrjStd" />
                         <IgnoreProduct ID="SharePointDesigner" />
                         <IgnoreProduct ID="VisPro" />
                         <IgnoreProduct ID="VisStd" />
                     </RemoveMSI>
                     <AppSettings>
                         <User Key="software\microsoft\office\16.0\common\graphics" Name="disablehardwareacceleration" Value="1" Type="REG_DWORD" App="office16" Id="L_DoNotUseHardwareAcceleration" />
                     </AppSettings>
                     <Display Level="Full" AcceptEULA="TRUE" />
                     <Logging Level="Off" />
                     </Configuration>
                 
"@
                     $config | Out-File "$scriptFolderLocation\OfficeConfig.xml"
                 
                 
                     send-CustomToastNofication -text "Starting installation"

                     send-log -logText "Installation Argumentlist is "
                     try{
                             #start-process choco -ArgumentList "install office365business --version=16731.20354 --forcex86 /configpath:$ScriptFolderLocation\config.xml -y --force" -ErrorAction Stop -Wait
                             start-process choco -ArgumentList "install office365business -y --force --params `"'/configpath:$scriptfolderlocation\OfficeConfig.xml /eula:FALSE'`"" -ErrorAction Stop -Wait
                 
                             $postSoftwareCheck = check-softwarePresence -softwareName $SoftwareName -includeChoco -chocolateyName $ChocolateyName
             
                 
                         if ($postSoftwareCheck.installed){
                 
                             
                             send-Log -logText "Successfully installed Microsoft 365 Apps for business" -addDashes Below -addTeamsMessage
                             send-CustomToastNofication -text "Successfully installed O365" 
                 
                 
                         }else{ #try to install again with ignore checksums
                 
                 
                             try{
                             #  start-process choco -ArgumentList "install googlechrome -y --force --ignore-checksums" -ErrorAction Stop -Wait  | Out-Null
                                 start-process choco -ArgumentList "install office365business -y --force --ignore-checksums --params `"'/configpath:$scriptfolderlocation\OfficeConfig.xml /eula:FALSE'`"" -ErrorAction Stop -Wait
                 
                                 $postSoftwareCheck = check-softwarePresence -softwareName $SoftwareName -includeChoco -chocolateyName $ChocolateyName
                 
                                 if ($postSoftwareCheck.installed){
                 
                                     send-Log -logText "Successfully installed Microsoft 365 Apps for business" -addDashes Below -addTeamsMessage
                                     send-CustomToastNofication -text "Successfully installed O365" 
                         
                                 }else{
                                     send-Log -logText "Failed to install Microsoft 365 Apps for business" -addDashes Below -type warning -addTeamsMessage
                                     send-CustomToastNofication -text "Failed to install O365" -type warning
                                 }
                                 
                             }catch{
                                 send-Log -logText "Failed to install Microsoft 365 Apps for business" -addDashes Below -type warning -catch
                                 send-CustomToastNofication -text "Failed to install O365" -type warning
                             }
                 
                 
                         
                         }
                 
                     }catch{
                         send-Log "Failed to install Microsoft 365 Apps for business*. Trying again with --ignore-checksums option in Chocolatey" -addDashes Below
                         
                 
                                 try{
                                     start-process choco -ArgumentList "install office365business -y --force --ignore-checksums --params `"'/configpath:$scriptfolderlocation\OfficeConfig.xml /eula:FALSE'`"" -ErrorAction Stop -Wait
             
                                     $postSoftwareCheck = check-softwarePresence -softwareName $SoftwareName -includeChoco -chocolateyName $ChocolateyName
             
                                     if ($postSoftwareCheck.installed){
                                     
                                         send-Log -logText "Successfully installed Microsoft 365 Apps for business" -addDashes Below  -addTeamsMessage
                                         send-CustomToastNofication -text "Successfully installed O365" 
                 
                                     }else{
             
                                         send-Log -logText "Failed to install Microsoft 365 Apps for business" -addDashes Below -type warning -addTeamsMessage
                                         send-CustomToastNofication -text "Failed to install O365" -type warning
                 
                                     }
                             
                                 }catch{
                                     send-Log -logText "Failed to install Microsoft 365 Apps for business" -addDashes Below -type warning -catch -addTeamsMessage
                                     send-CustomToastNofication -text "Failed to install O365" -type warning
                                 }
                     }
                     
                 
         }

        }