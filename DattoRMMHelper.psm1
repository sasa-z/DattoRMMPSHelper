

$rootScriptFolder = "c:\yw-data\automate"
$FolderForToastNotifications = "c:\yw-data\Toast_Notification_Files"

# $EnvDattoVariablesValuesHashTable = @{}
# $EnvDattoVariablesValuesHashTable.Add("$($env:Action)", "What action you want to do?") #change this variable value according to Datto variables in this case, replace Action, and description etc..
# $EnvDattoVariablesValuesHashTable.Add("$($env:ToastNotifications)", "Do you want to show toast notifications?") #change this according to Datto variables
# $EnvDattoVariablesValuesHashTable.Add("$($env:AppRunning)", "What if Chrome  is running?") #change this according to Datto variables
# $EnvDattoVariablesValuesHashTable.Add("$($env:NumberOfPCsToPushTo)", "How many PCs to run against?") #change this according to Datto variables
# $EnvDattoVariablesValuesHashTable.Add("$($env:SendFinalResultToTeams)", "Send final script result to Teams channel?") #change this according to Datto variables

$EnvDattoVariablesValuesHashTable = @{}
$EnvDattoVariablesValuesHashTable.Add("Action", "What action you want to do?") #change this variable value according to Datto variables in this case, replace Action, and description etc..
$EnvDattoVariablesValuesHashTable.Add("ToastNotifications", "Do you want to show toast notifications?") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("AppRunning", "What if Chrome  is running?") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("NumberOfPCsToPushTo", "How many PCs to run against?") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("SendFinalResultToTeams", "Send final script result to Teams channel?") #change this according to Datto variables


# $dattoEnvironmentVaribleValue = $env:ToastNotifications #pull from Datto
$dattoEnvironmentVaribleValue = "all" 



function send-Log {

     <#
        .SYNOPSIS
            Writes log to a text file and console
        .DESCRIPTION
            Writes log to a text file stored in script folder and some of the output will be written to console. Some of this information 
            is used by other functions in this module such as sending notification to Microsoft Teams
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder and it is required parameter
            send-log -scriptname "some script name" -logText "some text" -type "Info" -addDashes "Below" -addTeamsMessage 
        .PARAMETER rootScriptFolder
            This paramter is required. It is a root folder for all scripts. E.g. c:\automations. It shold be full path.
        .PARAMETER logText
            It is a text that we want to write to log file or/and console
        .PARAMETER type
            Depends on the type, we write log file to specific file. E.g. Errors are stored in Errors.txt file in script folder
        .PARAMETER addDashes
            It is used to make script output look nice and to make it easier to read. You can add dashes above and below the text
        .PARAMETER addTeamsMessage
            It stores log file we want into a text file that will be used by send-ToastNotification function within this module
        .EXAMPLE
            send-log -scriptname "Foxit_PDF_Reader" rootScriptFolder "c:\automations" -logText "some text" -type "error" 
            send-log -scriptname "some script name" rootScriptFolder "c:\automations" -logText "some text" -type "Info" -addDashes "Below" -addTeamsMessage 
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$logText,
        [ValidateSet('Info', 'Error', 'Warning')]
        [string]$type = "Info",
        [ValidateSet('Below', 'Above')]
        [string]$addDashes,
        [switch]$addTeamsMessage,
        [Parameter(Mandatory=$true)]
        [string]$scriptname = $scriptname,
        [Parameter(Mandatory=$true)]
        [string]$rootScriptFolder = $rootFolderForAllScriptFullPath,
        [switch]$skipWriteHost

    )

    $scriptFolderLocation = "$rootScriptFolder\$scriptName"
    if(-not (test-path "$rootScriptFolder\$scriptName")){New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null}

    if(-not (test-path "$scriptFolderLocation\Logs.txt")){New-Item -Path "$scriptFolderLocation" -Name 'Logs.txt' -ItemType File -Force | out-null } 
    if(-not (test-path "$scriptFolderLocation\Hidden_Files")){New-Item -Path $ScriptFolderLocation -Name "Hidden_Files" -ItemType Directory -Force -ErrorAction Stop | out-null } 
    $Folder = get-item "$($ScriptFolderLocation)\Hidden_Files" -Force
    $Folder.Attributes = "Hidden"


    $ErrorLogFile = "$rootScriptFolder\$scriptName\errors.txt"
    $WarningLogFile = "$rootScriptFolder\$scriptName\warnings.txt"
    $logFile = "$rootScriptFolder\$scriptName\logs.txt"

   
 
    $DateStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($addDashes -eq "Above"){
        if($type -eq "Error"){
            $logString = "---------------------------------------------------------`n$DateStamp : _Error | $logText"
        }elseif($type -eq "Warning"){
            $logString = "---------------------------------------------------------`n$DateStamp : _Warning | $logText"
        }elseif($type -eq "Info"){
            $logString = "---------------------------------------------------------`n$DateStamp :  $logText"
        }

    }elseif($addDashes -eq "Below"){
        if($type -eq "Error"){
            $logString = "$DateStamp : _Error | $logText`n---------------------------------------------------------"
        }elseif($type -eq "Warning"){
            $logString = "$DateStamp : _Warning | $logText`n---------------------------------------------------------"
        }elseif($type -eq "Info"){
            $logString = "$DateStamp :  $logText`n---------------------------------------------------------"
        }        
    }else{
        $logString = "$DateStamp :  $logText"
    }
   
    if ($type -eq "Error"){
            Add-Content -Value $logString -Path $ErrorLogFile 
            if(-not $skipWriteHost.IsPresent){write-host $logString}     
    }elseif($type -eq "Warning"){
            Add-Content -Value $logString  -Path $WarningLogFile
            if(-not $skipWriteHost.IsPresent){write-host $logString}    
    }else{
            Add-Content -Value $logString -Path $LogFile
            if(-not $skipWriteHost.IsPresent){write-host $logString}   
    }

   

    if($addTeamsMessage.IsPresent){
         $logText | out-file "$rootScriptFolder\$scriptName\Hidden_Files\TeamsMessage.txt" -Force -ErrorAction SilentlyContinue
    }
}

Function add-ScriptWorkingFoldersAndFiles{

     <#
        .SYNOPSIS
            Create all files into script folder
        .DESCRIPTION
            It creates all files into script and root scripts working folder that will be used by other functions such as logs.txt, errors.txt, warnings.txt
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder and it is required parameter. 
            It is required parameter
        .PARAMETER rootScriptFolder
            This paramter is required. It is a root folder for all scripts. E.g. c:\automations. It shold be full path. 
            It is required parameter
        .PARAMETER ToastNotificationAppLogo
            This parameter is used to copy app logo into script folder from Datto RMM. It is used in send-ToastNotification function within this module
        .PARAMETER EnvDattoVariablesValuesHashTable
            It is a hashtable that contains all Datto RMM environment variables so we see all variables used and their values in log file and Datto component output
        .PARAMETER toastNotificationTableCSV
            This parameter defines where CSV file is located. It is used in send-ToastNotification function within this module
            as run-ascurrentUser module can't read variables outside of its scope so we have to store it in CSV file and read it from there
        .EXAMPLE
            add-ScriptWorkingFoldersAndFiles -scriptname "Foxit_PDF_Reader" -rootScriptFolder  "c:\automations" -ToastNotificationAppLogo "c:\automations\FoxitPDFReader.png"            
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>
    
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$rootScriptFolder,
    [Parameter(Mandatory=$true)]
    [string]$scriptname,
    [string]$FolderForToastNotifications = $FolderForToastNotifications,
    [string]$ToastNotificationAppLogo,
    [hashtable]$EnvDattoVariablesValuesHashTable = $EnvDattoVariablesValuesHashTable
)

$scriptFolderLocation = "$rootScriptFolder\$scriptName"
$root = $rootScriptFolder
$toastNotificationsTableCSV = "$rootScriptFolder\$scriptName\Hidden_Files\ToastNotificationValuesTable.csv" #this is used in send-ToastNotification function

#region create Automate folder and log file in c:\yw-data\ (no values/variables to change)
try{
    New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null
    
    #remove previous folder is exists
    if (test-path "$rootScriptFolder\$scriptName"){
        Remove-Item -Path "$rootScriptFolder\$scriptName\" -Recurse -Force
    }
    New-Item -Path "$scriptFolderLocation" -Name 'Logs.txt' -ItemType File -Force | out-null

    #create hidden folder to copy files for toast notifications
    try{
        
        New-Item -Path (Split-Path $FolderForToastNotifications -Parent) -Name (Split-Path $FolderForToastNotifications -Leaf) -ItemType Directory -Force -ErrorAction Stop | out-null
        $Folder = get-item "$FolderForToastNotifications" -Force
        $Folder.Attributes = "Hidden"

        <#
        As runas-currentuser module can't use variables from this script, we have to store them into files into this hidden folder. 
        We will use this hidden folder to copy files into it. We also, may use it for some other things
        #>
        New-Item -Path $ScriptFolderLocation -Name "Hidden_Files" -ItemType Directory -Force -ErrorAction Stop | out-null
        $Folder = get-item "$($ScriptFolderLocation)\Hidden_Files" -Force
        $Folder.Attributes = "Hidden"

        #region ToastNotification files

        #create CSV file for Toast Notification information
        "" | select-object ToastHeader, ToastText, ToastAppLogo, ToastIdentifierName, type, DattoRMMValue, UniqueIdentifier, ifUserLoggedIn | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
        $WorkingCSVFile = Import-Csv "$toastNotificationsTableCSV"
        $WorkingCSVFile.ToastHeader = $ToastNotificationHeader
        $WorkingCSVFile.ToastIdentifierName = $toastIdentifierName
        $WorkingCSVFile.ToastAppLogo = $ToastNotificationAppLogo 
        $WorkingCSVFile.DattoRMMValue = $dattoEnvironmentVaribleValue 
        $WorkingCSVFile.UniqueIdentifier = "$($toastIdentifierName)-0"
        $WorkingCSVFile | export-csv -path "$toastNotificationsTableCSV" -NoTypeInformation -ErrorAction Stop

        #endregion ToastNotification files

    }catch{
        send-Log -scriptname $scriptName -rootScriptFolder $root -logText "Failed to create "$FolderForToastNotifications" folder. Error:  $($_.exception.message)" -type Error
    
    }
   
    #file for last message that will be sent to Teams
    New-Item -Path $ScriptFolderLocation\Hidden_Files -Name "TeamsMessage.txt" -ItemType File -Force -ErrorAction Stop | out-null
    
    
    #copy Toast Notification logos uploaded to Datto RMM
    
    try{
        copy-item Warning.png -Destination "$($FolderForToastNotifications)\Warning.png" -Force -ErrorAction SilentlyContinue
        copy-item Error.png -Destination "$($FolderForToastNotifications)\Error.png" -Force -ErrorAction SilentlyContinue
        copy-item Success.png -Destination "$($FolderForToastNotifications)\Success.png" -Force -ErrorAction SilentlyContinue
        
    if ($ToastNotificationAppLogo){
        copy-item $ToastNotificationAppLogo -Destination "$($FolderForToastNotifications)\$($ToastNotificationAppLogo)" -Force -ErrorAction SilentlyContinue
    }
        copy-item Chocolatey.png -Destination "$($FolderForToastNotifications)\Chocolatey.png" -Force -ErrorAction SilentlyContinue
    }catch{
        send-Log -scriptname $scriptName -rootScriptFolder $root  -logText "Failed to copy Toast Notification logos :  $($_.exception.message)" -type Warning -scriptname $scriptname
    }
    
   
    #add variable values into log file
    send-log -scriptname $scriptName -rootScriptFolder $root -logText "Script name: $($ScriptName) " -scriptname $scriptname 
    send-log -scriptname $scriptName -rootScriptFolder $root -logText "----------------------------------------------------------------------" -scriptname $scriptname 
    send-log -scriptname $scriptName -rootScriptFolder $root -logText "Script variables values: " -scriptname $scriptname 
    send-log -scriptname $scriptName -rootScriptFolder $root "----------------------------------------------------------------------" -scriptname $scriptname

    if ($EnvDattoVariablesValuesHashTable){
        foreach ($dattoVar in $EnvDattoVariablesValuesHashTable.GetEnumerator()){
            send-log -scriptname $scriptName -rootScriptFolder $root -logText "$($dattoVar.Value) : $($dattoVar.name)" -scriptname $scriptname -skipWriteHost
        }
    }
    send-log -scriptname $scriptName -rootScriptFolder $root -logText "----------------------------------------------------------------------" -scriptname $scriptname

    send-log -scriptname $scriptName -rootScriptFolder $root -logText "Successfully created script working folder $ScriptFolderLocation" -scriptname $scriptname
   


}catch{
    send-Log -scriptname $scriptName -rootScriptFolder $root -logText "Failed to create script working folder $ScriptFolderLocation :  $($_.exception.message)" -type Error -addDashes Below -scriptname $scriptname

    exit 1
}
}

function send-CustomToastNofication {

    <#
        .SYNOPSIS
            Send Windows Toast Notifications
        .DESCRIPTION
            It sends Toast Notifications so we can track script execution or see script afinal result
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder and it is required parameter. 
            It is required parameter
        .PARAMETER rootScriptFolder
            This paramter is required. It is a root folder for all scripts. E.g. c:\automations. It shold be full path. 
            It is required parameter
        .PARAMETER text
            This parameter is used to set text in toast notification
            It is required parameter
        .PARAMETER type
            It is notification type. Depending on the type, different toast notification logo will be used. We upload logo to Datto component and copy to local workstation
            If not provided, default toast notification logo will be used 'success.png'
            Values are Success, Error, Warning
        .PARAMETER header
            This parameter is used to set header in toast notification
            it is required parameter
        .PARAMETER DattoRMMToastValue
            This parameter is used to determine when and what Toast nofications will be sent. We pull this from Datto RMM variable
            Datto variable Values are All, WarningsErrors, Errors, None 
            e.g. If none is set in Datto, no toast notifications will be sent etc.
        .PARAMETER header
            This parameter is used to set header in toast notification
            it is required parameter
        .EXAMPLE
            add-ScriptWorkingFoldersAndFiles -scriptname "Foxit_PDF_Reader" -rootScriptFolder  "c:\automations" -ToastNotificationAppLogo "c:\automations\FoxitPDFReader.png"            
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$rootScriptFolder,    
        [ValidateSet('Success', 'Error', 'Warning')]
        [string]$type = "Success",
        [Parameter(Mandatory=$true)]
        [string]$text,
        [Parameter(Mandatory=$true)]
        [string]$scriptname,
        [Parameter(Mandatory=$true)]
        [string]$Header,
        [Parameter(Mandatory=$true)]
        [string]$DattoRMMToastValue,
        [string]$FolderForToastNotifications = $FolderForToastNotifications,
        [string]$ToastNotificationAppLogo = $ToastNotificationAppLogo
            
    )

    $scriptFolderLocation = "$rootScriptFolder\$scriptName"

    #region toast notification items

    if (-not (Test-Path -Path $FolderForToastNotifications)){
        New-Item -Path (Split-Path $FolderForToastNotifications -Parent) -Name (Split-Path $FolderForToastNotifications -Leaf) -ItemType Directory -Force -ErrorAction Stop | out-null
        $Folder = get-item "$FolderForToastNotifications" -Force
        $Folder.Attributes = "Hidden"
    }
a
    #create CSV file if it doesn't exist
    if (-not (Test-Path -Path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv")){   
        "" | select-object ToastHeader, ToastText, ToastAppLogo, ToastIdentifierName, type, DattoRMMValue, UniqueIdentifier, ifUserLoggedIn | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
    }

    $WorkingCSVFile = Import-Csv "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" 

    $ifUserLoggedInCheck  = (Get-WmiObject -ClassName Win32_ComputerSystem).Username

    #add info if user is logged in
    if ($ifUserLoggedInCheck ){

        $WorkingCSVFile | ForEach-Object {$_.ifUserLoggedIn = 'Yes'}
        $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
    }else{
        $WorkingCSVFile | ForEach-Object {$_.ifUserLoggedIn = 'No'}
        $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
    }
   

    #add Datto RMM variable value for ToastNotifications in CSV
    $WorkingCSVFile | ForEach-Object {$_.DattoRMMValue = $DattoRMMToastValue}
    $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
    
    #add initial value for ToastNotification UniqueIdentifier identifier
    $WorkingCSVFile | ForEach-Object {$_.UniqueIdentifier = "$($scriptname)---0"}
    $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop


    #add text for ToastNotifications in CSV
    $WorkingCSVFile | ForEach-Object {$_.ToastText = $text}
    $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop

    #add ToastHeader for ToastNotifications in CSV
    $WorkingCSVFile | ForEach-Object {$_.ToastHeader = $Header}
    $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
    
     #add ToastHeaderidentifier
     $WorkingCSVFile | ForEach-Object {$_.ToastIdentifierName = $scriptname}
     $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
     

    #Add what type of toast notification and we are sending in CSV and logo to be used
    if ($type -eq 'success'){

        $WorkingCSVFile | ForEach-Object {$_.type = 'Success'}
        $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
       
    }elseif($type -eq 'error'){
        $WorkingCSVFile | ForEach-Object {$_.type = 'Error'}
        $WorkingCSVFile | ForEach-Object {$_.ToastAppLogo = 'Error.png'}
        $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop

    }elseif ($type -eq 'warning'){
        $WorkingCSVFile | ForEach-Object {$_.type = 'Warning'}
        $WorkingCSVFile | ForEach-Object {$_.ToastAppLogo = 'Warning.png'}
        $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
    }
   


    #endregion

    #export root script info to csv as invoke-ascurrentuser can't read variables outside of its scope
    remove-item "$($rootScriptFolder)\tempInfo.csv" -ErrorAction SilentlyContinue
    "" | Select-Object "ScriptName", "ScriptFolderLocation", "rootScriptFolder","FolderForToastNotifications" | Export-Csv -Path "$($rootScriptFolder)\tempInfo.csv" -NoTypeInformation
    $ImportTempCSVInfo = Import-Csv "$($rootScriptFolder)\tempInfo.csv"
    $ImportTempCSVInfo.ScriptName = $scriptName
    $ImportTempCSVInfo.ScriptFolderLocation = $scriptFolderLocation
    $ImportTempCSVInfo.rootScriptFolder = $rootScriptFolder
    $ImportTempCSVInfo.FolderForToastNotifications = $FolderForToastNotifications
    $ImportTempCSVInfo | Export-Csv -Path "$($rootScriptFolder)\tempInfo.csv" -NoTypeInformation
    
    
    Invoke-AsCurrentUser {

        $ScriptName = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty ScriptName
        $ScriptFolderLocation = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty ScriptFolderLocation
        $rootScriptFolder = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty rootScriptFolder
        $FolderForToastNotifications = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty FolderForToastNotifications
        remove-item "$($rootScriptFolder)\tempInfo.csv" -ErrorAction SilentlyContinue

        $WorkingCSVFile = Import-Csv "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv"
       
        #get identifier number for toast notification
        [int]$UniqueIdentifierNumber = ($WorkingCSVFile.UniqueIdentifier -split '---' | select-object -last 1)
        $ToastHeader = $WorkingCSVFile.ToastHeader 
        $ToastText = $WorkingCSVFile.ToastText 
        $toastAppLogo = $WorkingCSVFile.ToastAppLogo 
        $toastIdentifierName = $WorkingCSVFile.ToastIdentifierName
        $toastType = $WorkingCSVFile.type
        $toastIdentifier = $toastIdentifierName + '---' + $UniqueIdentifierNumber
        $DattoToastNotificationVar = $WorkingCSVFile.DattoRMMValue
        $userIsLoggedIn = $WorkingCSVFile.ifUserLoggedIn
        
        if ($userIsLoggedIn -eq "Yes"){ #skip notifications if user not logged in
            #handle when and if Toast notification will be pushed depending on what was selected in Datto and what type of notification in script (success, error or warning)
            if ($DattoToastNotificationVar -eq 'All'){ #alway push toast notifications
            
                $toastIdentifier = $toastIdentifierName + '---' + $UniqueIdentifierNumber++
                New-BurntToastNotification -Text "$($ToastHeader)","$($ToastText)" -AppLogo "$FolderForToastNotifications\$ToastAppLogo" -UniqueIdentifier "$toastIdentifier"
                $toastIdentifier = $toastIdentifierName + '---' + $UniqueIdentifierNumber++
                
                #increase unique identifier number so the new notification doesn't overwrite previous one
                $WorkingCSVFile | foreach-object {
                    $_.UniqueIdentifier = $toastIdentifier
                }
                $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
                
            }elseif ($DattoToastNotificationVar -eq 'Errors' -and ($toastType -eq 'Error')){ #only push toast notifications with errors 
            
                $toastIdentifier = $toastIdentifierName + '-' + $UniqueIdentifierNumber++
                New-BurntToastNotification -Text "$($ToastHeader)","$($ToastText)" -AppLogo "c:\yw-data\Toast_Notification_Files\$($ToastAppLogo)" -UniqueIdentifier "$toastIdentifier"
                
                $toastIdentifier = $toastIdentifierName + '-' + $UniqueIdentifierNumber++
                
                #increase unique identifier number so the new notification doesn't overwrite previous one
                $WorkingCSVFile | foreach-object {
                    $_.UniqueIdentifier = $toastIdentifier
                }
                $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
                
            }elseif($DattoToastNotificationVar -eq 'WarningsErrors' -and ($toastType -eq 'Error' -or $toastType -eq 'Warning')){ #only push toast notifications with errors or warnings     

                $toastIdentifier = $toastIdentifierName + '-' + $UniqueIdentifierNumber++
                New-BurntToastNotification -Text "$($ToastHeader)","$($ToastText)" -AppLogo "c:\yw-data\Toast_Notification_Files\$($ToastAppLogo)" -UniqueIdentifier "$toastIdentifier"
    
                $toastIdentifier = $toastIdentifierName + '-' + $UniqueIdentifierNumber++
                
                #increase unique identifier number so the new notification doesn't overwrite previous one
                $WorkingCSVFile | foreach-object {
                    $_.UniqueIdentifier = $toastIdentifier
                }
                $WorkingCSVFile | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
                
            }
        }
        

    }


}


