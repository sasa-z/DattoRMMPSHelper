

function send-Log {

     <#
        .SYNOPSIS
            Writes log to a text file and console
        .DESCRIPTION
            Writes log to a text file stored in script folder and some of the output will be written to console. Some of this information 
            is used by other functions in this module such as sending notification to Microsoft Teams
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
                $scriptName
                $rootScriptFolder (it can be specified as global variable in Datto as well)
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder
        .PARAMETER rootScriptFolder
            It is a root folder for all scripts. E.g. c:\automations. It shold be full path.
        .PARAMETER logText
            It is a text that we want to write to log file or/and console
        .PARAMETER type
            Depends on the type, we write log file to specific file. E.g. Errors are stored in Errors.txt file in script folder
        .PARAMETER addDashes
            It is used to make script output look nice and to make it easier to read. You can add dashes above and below the text
        .PARAMETER addTeamsMessage
            It stores log file we want into a text file that will be used by send-ToastNotification function within this module
        .EXAMPLE
            send-log -logText "some text" -type error
            send-log -logText "some text" -type Info -addDashes "Below" -addTeamsMessage 
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
        [switch]$skipWriteHost,
        [switch]$catch,
        [string]$scriptname = $scriptname,
        [string]$rootScriptFolder = $rootScriptFolder

    )
    Write-Verbose " --------------------------"
    Write-Verbose "Executing send-Log function"
    Write-Verbose " --------------------------"

    if (-not $rootScriptFolder){
        $rootScriptFolder = $env:rootScriptFolder

        Write-Verbose "rootScriptFolder value pulled from Datto global variable"
    }
    Write-Verbose "Value of rootScriptFolder is $rootScriptFolder"

    if ($rootScriptFolder[-1] -like '\'){
        $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
    }else{
        $rootScriptFolder = $rootScriptFolder
    }

    $scriptFolderLocation = "$rootScriptFolder\$scriptName"
    if(-not (test-path "$rootScriptFolder\$scriptName")){New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null}
    Write-Verbose "Created script folder $scriptFolderLocation"

    if(-not (test-path "$scriptFolderLocation\Logs.txt")){New-Item -Path "$scriptFolderLocation" -Name 'Logs.txt' -ItemType File -Force | out-null } 
    if(-not (test-path "$scriptFolderLocation\Hidden_Files")){New-Item -Path $ScriptFolderLocation -Name "Hidden_Files" -ItemType Directory -Force -ErrorAction Stop | out-null } 
    $Folder = get-item "$($ScriptFolderLocation)\Hidden_Files" -Force
    $Folder.Attributes = "Hidden"
    Write-Verbose "Created Hidden_Files folder"
    

    Write-Verbose "Parameters are:"
    Write-Verbose "logText is : $logText " 
    Write-Verbose "type is : $type " 
     Write-Verbose "addDashes is :  $addDashes "
    Write-Verbose "addTeamsMessage is:  $addTeamsMessage "
    Write-Verbose "scriptname is : $scriptname " 
    Write-Verbose "rootScriptFolder is $rootScriptFolder"

    $ErrorLogFile = "$rootScriptFolder\$scriptName\errors.txt"
    Write-Verbose "Location for ErrorLogFile is $ErrorLogFile"
    $WarningLogFile = "$rootScriptFolder\$scriptName\warnings.txt"
    Write-Verbose "Location for WarningLogFile is $WarningLogFile"
    $logFile = "$rootScriptFolder\$scriptName\logs.txt"
    Write-Verbose "Location for logFile is $logFile"

   
 
    $DateStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
        if ($catch.IsPresent){
            $logString = "$logtext. Error: $($_.exception.message), at line $($_.InvocationInfo.ScriptLineNumber)"
        }

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
            Write-Verbose "Added $logString to $ErrorLogFile"
            if(-not $skipWriteHost.IsPresent){write-host $logString}     
    }elseif($type -eq "Warning"){
            Add-Content -Value $logString  -Path $WarningLogFile
            Write-Verbose "Added $logString to $WarningLogFile"
            if(-not $skipWriteHost.IsPresent){write-host $logString}    
    }else{
            Add-Content -Value $logString -Path $LogFile
            Write-Verbose "Added $logString to $LogFile"
            if(-not $skipWriteHost.IsPresent){write-host $logString}   
    }

   

    if($addTeamsMessage.IsPresent){
         $logText | out-file "$rootScriptFolder\$scriptName\Hidden_Files\TeamsMessage.txt" -Force -ErrorAction SilentlyContinue
         Write-Verbose "Added $logText to $rootScriptFolder\$scriptName\Hidden_Files\TeamsMessage.txt"
    }
}

function get-runAsUserModule {

     <#
        .SYNOPSIS
            Check if RunAsUserModule is installed
        .DESCRIPTION
            Check if RunAsUserModule is installed and add information to log file
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
                $scriptName
                $rootScriptFolder
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder
        .PARAMETER rootScriptFolder
            It is a root folder for all scripts. E.g. c:\automations. It shold be full path.
        .EXAMPLE
            get-runAsUserModule 
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>

    [CmdletBinding()]
        param(
            [string]$rootScriptFolder = $rootScriptFolder,
            [string]$scriptname = $scriptname
        )


        if (-not $rootScriptFolder){
            $rootScriptFolder = $env:rootScriptFolder
        }

        if ($rootScriptFolder[-1] -like '\'){
            $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
        }else{
            $rootScriptFolder = $rootScriptFolder
        }

        Write-Verbose "Parameters are: "
        Write-Verbose "rootScriptFolder = $rootScriptFolder"
        Write-Verbose "scriptName = $scriptName"

#create script folder if it doesn't exist
if (-not (test-path "$rootScriptFolder\$scriptName")){New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null}


    $RunAsUser = get-module RunAsUser -ListAvailable -ErrorAction stop

    if ($RunAsUser){
        import-module RunAsUser  
    }else{
        try{
            Install-Module RunAsUser -Force -confirm:$false
    
            send-log  -logText "Successfully installed RunAsUser module" -type Info -addDashes Below
    
        }catch{
            send-log  -logText "Failed to install RunAsUser module with error. Exiting script" -type Error -addDashes Below -catch
            exit 1            
        }
    }
    
}

function get-Chocolatey {

    <#
       .SYNOPSIS
           Check if Chocolatey is installed
       .DESCRIPTION
           Check if Chcoolatey is installed and add information to log file
           For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
               $scriptName
               $rootScriptFolder
       .PARAMETER scriptName
           It is a script name that we use to create a folder for this script in root scripts working folder
       .PARAMETER rootScriptFolder
           It is a root folder for all scripts. E.g. c:\automations. It shold be full path.
       .EXAMPLE
           get-chocolatey 
       .OUTPUTS
       .NOTES
           FunctionName : 
           Created by   : Sasa Zelic
           Date Coded   : 12/2019
    #>

   [CmdletBinding()]
       param(
           [string]$rootScriptFolder = $rootScriptFolder,
           [string]$scriptname = $scriptname
       )


       if (-not $rootScriptFolder){
           $rootScriptFolder = $env:rootScriptFolder
       }

       if ($rootScriptFolder[-1] -like '\'){
           $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
       }else{
           $rootScriptFolder = $rootScriptFolder
       }

#create script folder if it doesn't exist
if (-not (test-path "$rootScriptFolder\$scriptName")){New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null}

try{

    $ChocoCheck = get-command choco.exe -ErrorAction SilentlyContinue

    if ($ChocoCheck){   

        send-log -logText "Chocolatey is already installed" -addDashes Below
        
        $ChocoUpdateNeeded = choco outdated -r | select-string 'chocolatey'

        if ($ChocoUpdateNeeded){

            try{
                start-process -FilePath choco -ArgumentList "upgrade chocolatey -y" -ErrorAction stop -Wait  | Out-Null
                send-log -logText "Successfully updated Chocolatey" -addDashes Below
    
            }catch{
                send-log -logText "Failed to update Chocolatey" -type Warning -catch
            }

        }

       

    }else{ #install chocolatey
        try{
            Set-ExecutionPolicy Bypass -Scope Process -Force; 
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; 
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) -ErrorAction Stop

            send-log -logText "Successfully installed Chocolatey" -addDashes Below
        }catch{
            send-log -logText "Failed to install chocolatey" -type Error -addTeamsMessage -catch
            exit 1

        }

    }
    
    }catch{
        send-log -logText "Failed to install chocolatey" -type Error -addTeamsMessage -catch
        exit 1
}


   
}

function get-BurntToastModule {

     <#
        .SYNOPSIS
            Check if RunAsUserModule is installed
        .DESCRIPTION
            Check if RunAsUserModule is installed and add information to log file
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
                $scriptName
                $rootScriptFolder 
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder
        .PARAMETER rootScriptFolder
            It is a root folder for all scripts. E.g. c:\automations. It shold be full path.
        .EXAMPLE
            get-burntToastModule 
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>


[CmdletBinding()]
     param(
         [string]$rootScriptFolder = $rootScriptFolder,
         [string]$scriptname = $scriptname
     )

     if (-not $rootScriptFolder){
        $rootScriptFolder = $env:rootScriptFolder
    }

    if ($rootScriptFolder[-1] -like '\'){
        $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
    }else{
        $rootScriptFolder = $rootScriptFolder
    }

#create script folder if it doesn't exist
if (-not (test-path "$rootScriptFolder\$scriptName")){New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null}


    $BruntToast = get-module BurntToast -ListAvailable -ErrorAction stop
if ($BruntToast){
    try{
        Import-Module BurntToast -ErrorAction stop

        send-log -logText "Successfully imported BurntToast module"
        
    }catch{
        send-log -logText "Failed to import BurntToast module with error" -type Warning -catch
        
    }
    
}else{
    try{
        Install-Module -Name BurntToast -RequiredVersion 0.8.5  -Confirm:$false -Force

        send-log -logText "Successfully installed BurntToast module"
        

    }catch{
        send-log -logText "Failed to install BurntToast module with error" -type Warning -catch

    }
}
}

Function add-ScriptWorkingFoldersAndFiles{

     <#
        .SYNOPSIS
            Create all files into script folder
        .DESCRIPTION
            It creates all files into script and root scripts working folder that will be used by other functions such as logs.txt, errors.txt, warnings.txt
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
                $scriptName
                $rootScriptFolder
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder and it is required parameter. 
            It is required parameter
        .PARAMETER rootScriptFolder
            It is a root folder for all scripts. E.g. c:\automations. It shold be full path. 
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

    [string]$ToastNotificationAppLogo = $ToastNotificationAppLogo,
    [string]$FolderForToastNotifications,
    [hashtable]$EnvDattoVariablesValuesHashTable = $EnvDattoVariablesValuesHashTable,
    [string]$rootScriptFolder = $rootScriptFolder,
    [string]$scriptname = $scriptname
)

Write-Verbose " --------------------------"
Write-Verbose "Executing add-ScriptWorkingFoldersAndFiles function"
Write-Verbose " --------------------------"

if (-not $rootScriptFolder){
    $rootScriptFolder = $env:rootScriptFolder
}


if ($rootScriptFolder[-1] -like '\'){
    $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
}else{
    $rootScriptFolder = $rootScriptFolder
    
}

$scriptFolderLocation = "$rootScriptFolder\$scriptName"


if(-not $FolderForToastNotifications){
    $partForToastNOtifications = (Split-Path $rootScriptFolder -Parent)
    $FolderForToastNotifications = "$partForToastNOtifications\Toast_Notification_Files"
}

Write-Verbose "Parameters are: "
Write-Verbose "ToastNotificationAppLogo: $ToastNotificationAppLogo"
Write-Verbose "rootScriptFolder: $rootScriptFolder"
Write-Verbose "scriptFolderLocation: $scriptFolderLocation"
Write-Verbose "FolderForToastNotifications: $FolderForToastNotifications"
Write-Verbose "EnvDattoVariablesValuesHashTable: $EnvDattoVariablesValuesHashTable"
Write-Verbose "scriptName: $scriptName"


#region create Automate folder and log file in c:\yw-data\ (no values/variables to change)
try{
    
    #remove previous folder is exists
    if (test-path "$rootScriptFolder\$scriptName"){
        Write-Verbose "Removing previous folder:  $rootScriptFolder\$scriptName "
        Remove-Item -Path "$rootScriptFolder\$scriptName\" -Recurse -Force
    }
    
    New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null
    Write-Verbose "Created scriptname folder:  $rootScriptFolder\$scriptName "

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
      
        #endregion ToastNotification files

    }catch{
        send-Log -logText "Failed to create "$FolderForToastNotifications" folder. " -type Error -catch
    
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
        send-Log -logText "Failed to copy Toast Notification logos" -type Warning  -catch
    }
    
   
    #add variable values into log file
    send-log  -logText "Script name: $($ScriptName) " 
    send-log  -logText "----------------------------------------------------------------------" 
    send-log  -logText "Script variables values: "
    send-log  -logText  "----------------------------------------------------------------------"

    if ($EnvDattoVariablesValuesHashTable){
        foreach ($dattoVar in $EnvDattoVariablesValuesHashTable.GetEnumerator()){
            send-log  -logText "$($dattoVar.Value) : $($dattoVar.name)" -skipWriteHost
        }
    }
    send-log  -logText "----------------------------------------------------------------------" 

    send-log  -logText "Successfully created script working folder $ScriptFolderLocation" 
   


}catch{
    send-Log -logText "Failed to create script working folder $ScriptFolderLocation" -type Error -addDashes Below  -catch

    exit 1
}
}

function send-CustomToastNofication {

    <#
        .SYNOPSIS
            Send Windows Toast Notifications
        .DESCRIPTION
            It sends Toast Notifications so we can track script execution or see script final result
            
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
                scriptName
                rootScriptFolder
                ToastNotifications (with 'All', 'WarningsErrors', 'Errors', 'None' values in Datto RMM)
            
            These variables needs to be defined in script or global scrope
                toastNotificationAppLogo e.g. 'Chocolatey.png'
                ToastNotificationHeader
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder. 
            It is required parameter
        .PARAMETER rootScriptFolder
            This paramter is required. It is a root folder for all scripts. E.g. c:\automations. It shold be full path. 
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
        .PARAMETER ToastNotifications
            This parameter is used to determine when and what Toast nofications will be sent. We pull this from Datto RMM variable
            Datto variable Values are All, WarningsErrors, Errors, None 
            e.g. If none is set in Datto, no toast notifications will be sent etc.
        .PARAMETER header
            This parameter is used to set header in toast notification
            it is required parameter
        .PARAMETER FolderForToastNotifications
            This parameter determine where Toast notification files will be stored.
            If not specified, script will use default folder
        .EXAMPLE
            send-CustomToastNofication -header "Foxit PDF Reader" -text "Installation completed successfully" -type warning 
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>
    
    [CmdletBinding()]
    param(

        [string]$Header = $ToastNotificationHeader,
        [Parameter(Mandatory=$true)]
        [string]$text,
        [ValidateSet('Success', 'Error', 'Warning')]
        [string]$type = "Success",
        [ValidateSet('All', 'WarningsErrors', 'Errors', 'None')]
        [string]$ToastNotifications = $ToastNotifications,
        [string]$ToastNotificationAppLogo = $ToastNotificationAppLogo,
        [string]$FolderForToastNotifications,
        [string]$scriptname = $scriptname,
        [string]$rootScriptFolder = $rootScriptFolder
            
    )

    Write-Verbose " --------------------------"
    Write-Verbose "Executing send-CustomToastNofication function"
    Write-Verbose " --------------------------"

    Write-Verbose "Function send-CustomToastNofication started executing"

    if (-not $rootScriptFolder){
        $rootScriptFolder = $env:rootScriptFolder
        Write-Verbose "Root Script Folder not provided. Pulling from Datto RMM variable: $($rootScriptFolder)"

    }

    if (-not $ToastNotifications){
        $ToastNotifications = $env:ToastNotifications
        Write-Verbose "ToastNotifications not provided. Pulling from Datto RMM variable: $($ToastNotifications)"
    }

    if ($rootScriptFolder[-1] -like '\'){
        $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
    }else{
        $rootScriptFolder = $rootScriptFolder
        
    }

    if (-not $header){
        $header = $ToastNotificationHeader}
        Write-Verbose "Header not provided via parameter. Pulling from script or global scope: $($ToastNotificationHeader)"

    
    
    $scriptFolderLocation = "$rootScriptFolder\$scriptName"
    $CSVTAblePath = "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv"
    #region toast notification items
    
    
    #if folder for toast notifications not provided via parameter, create it below
    if(-not $FolderForToastNotifications){
        $partForToastNOtifications = (Split-Path $rootScriptFolder -Parent)
        $FolderForToastNotifications = "$partForToastNOtifications\Toast_Notification_Files"
        Write-Verbose "Folder for toast notifications: $FolderForToastNotifications"
    }

    #create hidden folder if it doesn't exist in script folder location
    if(-not (Test-Path -Path "$($ScriptFolderLocation)\Hidden_Files")){
        New-Item -Path $ScriptFolderLocation -Name 'Hidden_Files' -ItemType Directory -Force -ErrorAction Stop | out-null
        $Folder = get-item "$($ScriptFolderLocation)\Hidden_Files" -Force
        $Folder.Attributes = "Hidden"
        Write-Verbose "Folder for toast notifications created: HIdden_Files in $($scriptFolderLocation)"
    }


    if (-not (Test-Path -Path $FolderForToastNotifications)){
        New-Item -Path (Split-Path $FolderForToastNotifications -Parent) -Name (Split-Path $FolderForToastNotifications -Leaf) -ItemType Directory -Force -ErrorAction Stop | out-null
        $Folder = get-item "$FolderForToastNotifications" -Force
        $Folder.Attributes = "Hidden"
        Write-Verbose "Folder for toast notifications created: $($FolderForToastNotifications)"
    }

    Write-Verbose " "
    write-verbose "Parameter values:"
    Write-Verbose "Root script folder: $rootScriptFolder"
    Write-Verbose "Script name: $scriptName"
    Write-Verbose "Toast notifications: $ToastNotifications"
    Write-Verbose "Toast notification app logo: $ToastNotificationAppLogo"
    Write-Verbose "Toast notification header: $header"
    Write-Verbose "Toast notification text: $text"
    Write-Verbose "Toast notification type: $type"
    Write-Verbose "Folder for toast notifications: $FolderForToastNotifications"

    #create CSV file if it doesn't exist
    if (-not (test-path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv")){

    Write-Verbose "Creating CSV file: $($CSVTAblePath)"

        "" | select-object ToastHeader, ToastText, ToastAppLogo, ToastIdentifierName, type, DattoRMMValue, UniqueIdentifier, ifUserLoggedIn | export-csv -path "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv" -NoTypeInformation -ErrorAction Stop
       
        $WorkingCSVFile = Import-Csv $CSVTAblePath
        $WorkingCSVFile.UniqueIdentifier = "$($scriptname)---0"
        $WorkingCSVFile | export-csv -path $CSVTAblePath -NoTypeInformation -ErrorAction Stop
        Write-Verbose "CSV file created: $($CSVTAblePath)"


    }
        $WorkingCSVFile = Import-Csv $CSVTAblePath
        $WorkingCSVFile.ToastHeader = $Header
        $WorkingCSVFile.ToastIdentifierName = ($scriptname -replace " ", '')
        $WorkingCSVFile.ToastAppLogo = $ToastNotificationAppLogo 
        $WorkingCSVFile.DattoRMMValue = $dattoEnvironmentVaribleValue 
   
        $WorkingCSVFile.DattoRMMValue = $ToastNotifications
        $WorkingCSVFile.ToastText = $text
        $WorkingCSVFile | export-csv -path $CSVTAblePath -NoTypeInformation -ErrorAction Stop

     
   


    $WorkingCSVFile = Import-Csv $CSVTAblePath 

    $ifUserLoggedInCheck  = (Get-WmiObject -ClassName Win32_ComputerSystem).Username

    #add info if user is logged in
    if ($ifUserLoggedInCheck ){

        $WorkingCSVFile | ForEach-Object {$_.ifUserLoggedIn = 'Yes'}
        Write-Verbose "User is logged in"
        $WorkingCSVFile | export-csv -path  $CSVTAblePath -NoTypeInformation -ErrorAction Stop
    }else{
        $WorkingCSVFile | ForEach-Object {$_.ifUserLoggedIn = 'No'}
        Write-Verbose "User is not logged in"
        $WorkingCSVFile | export-csv -path  $CSVTAblePath -NoTypeInformation -ErrorAction Stop
    }
   

    #Add what type of toast notification and we are sending in CSV and logo to be used
    if ($type -eq 'success'){
        Write-Verbose "Type is success. Logo is $toastnoficationapplogo"
        $WorkingCSVFile | ForEach-Object {$_.type = 'Success'}
        $WorkingCSVFile | ForEach-Object {$_.ToastAppLogo = $ToastNotificationAppLogo}
        $WorkingCSVFile | export-csv -path  $CSVTAblePath -NoTypeInformation -ErrorAction Stop
       
    }elseif($type -eq 'error'){
        Write-Verbose "Type is error. Logo is Error.png"
        $WorkingCSVFile | ForEach-Object {$_.type = 'Error'}
        $WorkingCSVFile | ForEach-Object {$_.ToastAppLogo = 'Error.png'}
        $WorkingCSVFile | export-csv -path  $CSVTAblePath -NoTypeInformation -ErrorAction Stop

    }elseif ($type -eq 'warning'){
        Write-Verbose "Type is warning. Logo is Warning.png"
        $WorkingCSVFile | ForEach-Object {$_.type = 'Warning'}
        $WorkingCSVFile | ForEach-Object {$_.ToastAppLogo = 'Warning.png'}
        $WorkingCSVFile | export-csv -path  $CSVTAblePath -NoTypeInformation -ErrorAction Stop
    }
   

    #endregion

    #export root script info to csv as invoke-ascurrentuser can't read variables outside of its scope
    try{remove-item "$($rootScriptFolder)\tempInfo.csv" -ErrorAction SilentlyContinue}catch{}
    "" | Select-Object "ScriptName", "ScriptFolderLocation", "rootScriptFolder","FolderForToastNotifications" | Export-Csv -Path "$($rootScriptFolder)\tempInfo.csv" -NoTypeInformation
    $ImportTempCSVInfo = Import-Csv "$($rootScriptFolder)\tempInfo.csv"
    $ImportTempCSVInfo.ScriptName = $scriptName
    $ImportTempCSVInfo.ScriptFolderLocation = $scriptFolderLocation
    $ImportTempCSVInfo.rootScriptFolder = $rootScriptFolder
    $ImportTempCSVInfo.FolderForToastNotifications = $FolderForToastNotifications
    $ImportTempCSVInfo | Export-Csv -Path "$($rootScriptFolder)\tempInfo.csv" -NoTypeInformation
    
   
    
    Invoke-AsCurrentUser {

        $ScriptName = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty ScriptName
        $rootScriptFolder = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty rootScriptFolder
        $ScriptFolderLocation = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty ScriptFolderLocation
        $FolderForToastNotifications = import-csv c:\yw-data\automate\tempInfo.csv | select-object -expandproperty FolderForToastNotifications
        try{remove-item "$($rootScriptFolder)\tempInfo.csv" -ErrorAction SilentlyContinue}catch{}

        
        $CSVTAblePath = "$($ScriptFolderLocation)\Hidden_Files\ToastNotificationValuesTable.csv"

        $WorkingCSVFile = Import-Csv  $CSVTAblePath
       
        #get values from CSV
        [int]$UniqueIdentifierNumber = ($WorkingCSVFile.UniqueIdentifier -split '---' | select-object -last 1)
        $ToastHeader = $WorkingCSVFile.ToastHeader 
        $ToastText = $WorkingCSVFile.ToastText 
        $toastAppLogo = $WorkingCSVFile.ToastAppLogo 
        $toastIdentifierName = $WorkingCSVFile.ToastIdentifierName
        $toastType = $WorkingCSVFile.type
        $toastIdentifier = $toastIdentifierName + '---' + $UniqueIdentifierNumber
        $DattoToastNotificationVar = $WorkingCSVFile.DattoRMMValue
        $userIsLoggedIn = $WorkingCSVFile.ifUserLoggedIn

       


        #increase unique identifier number so the new notification doesn't overwrite previous one
        $Increase = $UniqueIdentifierNumber + 1
        $newToastIdentifier = $toastIdentifierName + '---' + $Increase
        $WorkingCSVFile | foreach-object {
            $_.UniqueIdentifier = $newToastIdentifier
        }
        $WorkingCSVFile | export-csv -path $CSVTAblePath -NoTypeInformation -ErrorAction Stop

        Add-Content -Value $toastAppLogo -Path "$FolderForToastNotifications\1.txt"
        
        if ($userIsLoggedIn -eq "Yes"){ #skip notifications if user not logged in
            #handle when and if Toast notification will be pushed depending on what was selected in Datto and what type of notification in script (success, error or warning)
            if ($DattoToastNotificationVar -eq 'All'){ #alway push toast notifications
                New-BurntToastNotification -Text "$($ToastHeader)","$($ToastText)" -AppLogo "$FolderForToastNotifications\$ToastAppLogo" -UniqueIdentifier "$toastIdentifier"
            }elseif ($DattoToastNotificationVar -eq 'Errors' -and ($toastType -eq 'Error')){ #only push toast notifications with errors        
                New-BurntToastNotification -Text "$($ToastHeader)","$($ToastText)" -AppLogo "c:\yw-data\Toast_Notification_Files\$($ToastAppLogo)" -UniqueIdentifier "$toastIdentifier"                           
            }elseif($DattoToastNotificationVar -eq 'WarningsErrors' -and ($toastType -eq 'Error' -or $toastType -eq 'Warning')){ #only push toast notifications with errors or warnings     
                New-BurntToastNotification -Text "$($ToastHeader)","$($ToastText)" -AppLogo "c:\yw-data\Toast_Notification_Files\$($ToastAppLogo)" -UniqueIdentifier "$toastIdentifier"                
            }
        }
        

    }


}


Function send-CustomFinalToastNotification {

     <#
        .SYNOPSIS
            Send Windows Toast Notifications
        .DESCRIPTION
            It sends Toast Notifications so we can track script execution or see script final result
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
                rootScriptFolder
                ToastNotifications (with 'All', 'WarningsErrors', 'Errors', 'None' values in Datto RMM or in script)
                SendFinalResultToTeams (with 'yes', 'no', 'ifsuccess', 'iferrors' values in Datto RMM )
            You need to provide these variable in script or global scope 
                scriptName e.g 'Foxit_PDF_Reader'
                ToastNotificationHeader e.g. 'Foxit PDF Reader'
        .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder and it is required parameter. 
            It is required parameter
        .PARAMETER rootScriptFolder
            This paramter is required. It is a root folder for all scripts. E.g. c:\automations. It shold be full path. 
            It is required parameter
        .PARAMETER header
            This parameter is used to set header in toast notification
        .PARAMETER ToastNotifications
            This parameter is used to determine when and what Toast nofications will be sent. We pull this from Datto RMM variable
            Datto variable Values are All, WarningsErrors, Errors, None 
            e.g. If none is set in Datto, no toast notifications will be sent etc.
        .PARAMETER Company
            This parameter is piece of information that are sent to teams
        .PARAMETER Action
            This parameter is piece of information that are sent to teams
        .PARAMETER SendToTeams
            This parameter determines if toast notification will be sent to teams
        .EXAMPLE
            send-CustomToastNofication -header "Foxit PDF Reader"   
        .OUTPUTS
        .NOTES
            FunctionName : 
            Created by   : Sasa Zelic
            Date Coded   : 12/2019
     #>
    
     [CmdletBinding()]
     param(
         [string]$Header = $ToastNotificationHeader,
         [ValidateSet('All', 'WarningsErrors', 'Errors', 'None')]
         [string]$ToastNotifications = $ToastNotifications,
         [string]$Company,
         [string]$Action,
         [switch]$SendToTeams,
         [string]$rootScriptFolder = $rootScriptFolder,    
         [string]$scriptname = $scriptname

             
     )

    Write-Verbose " --------------------------"
    Write-Verbose "Executing send-CustomFinalToastNotification function"
    Write-Verbose " --------------------------"
 
     if (-not $ToastNotifications){
        $ToastNotifications = $env:ToastNotifications
    }


    
    if (-not $Company){$Company = $ENV:CS_PROFILE_NAME}
    if (-not $Action){
        $Action = $ENV:Action
        Write-Verbose "Action not provided in parameter. Pulling value from Datto RMM : $($Action)"
    }

    if (-not $header){
        Write-Verbose "Header not provided in parameter. Pulling value from script scope : $($ToastNotificationHeader)"
        $header = $ToastNotificationHeader
    }
    $SendFinalResultToTeams = $ENV:SendFinalResultToTeams
    Write-Verbose "SendFinalResultToTeams value for finalToastlNotification is :  $SendFinalResultToTeams "

    if (-not $rootScriptFolder){
        $rootScriptFolder = $env:rootScriptFolder
        Write-Verbose "Root Script Folder not provided. Pulling from Datto RMM variable: $($rootScriptFolder)"

    }

    if ($rootScriptFolder[-1] -like '\'){
        $rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)
    }else{
        $rootScriptFolder = $rootScriptFolder 
    }

    $scriptFolderLocation = "$rootScriptFolder\$scriptName"

    Write-Verbose "scriptFolderLocation value for finalToastlNotification is :  $scriptFolderLocation "

        #folder for toast notifications not provided via parameter, create it below
        $partForToastNOtifications = (Split-Path $rootScriptFolder -Parent)
        $FolderForToastNotifications = "$partForToastNOtifications\Toast_Notification_Files"

        Write-Verbose "FolderForToastNotifications value for finalToastlNotification is :  $FolderForToastNotifications "
  
    Write-Verbose "Function parameters are: "
    Write-Verbose "Header :  $Header "
    Write-Verbose "ToastNotifications :  $ToastNotifications "
    Write-Verbose "Company :  $Company "
    Write-Verbose "Action :  $Action "
    Write-Verbose "SendToTeams :  $SendToTeams "
    Write-Verbose "rootScriptFolder :  $rootScriptFolder "
    Write-Verbose "scriptname :  $scriptname "


    $ifUserLoggedInCheck  = (Get-WmiObject -ClassName Win32_ComputerSystem).Username
    
    if($ifUserLoggedInCheck){
        $UserLoggedIn = 'Yes'
        Write-Verbose "UserLoggedIn value for finalToastlNotification is :  $UserLoggedIn "
    }else{
        $UserLoggedIn = 'No'
        Write-Verbose "UserLoggedIn value for finalToastlNotification is :  $UserLoggedIn "
    }

     #export root script info to csv as invoke-ascurrentuser can't read variables outside of its scope
     try{remove-item "$($rootScriptFolder)\tempFinalInfo.csv" -ErrorAction SilentlyContinue}catch{}
     "" | Select-Object "ScriptName", "ScriptFolderLocation", "rootScriptFolder","ToastNotifications","FolderForToastNotifications", "ToastHeader","UserLoggedIn" | Export-Csv -Path "$($rootScriptFolder)\tempFinalInfo.csv" -NoTypeInformation
     $ImportTempCSVInfo = Import-Csv "$($rootScriptFolder)\tempFinalInfo.csv"
     $ImportTempCSVInfo.ScriptName = $scriptName
     $ImportTempCSVInfo.ScriptFolderLocation = $scriptFolderLocation
     $ImportTempCSVInfo.rootScriptFolder = $rootScriptFolder
     $ImportTempCSVInfo.UserLoggedIn = $UserLoggedIn
     $ImportTempCSVInfo.FolderForToastNotifications = $FolderForToastNotifications
     $ImportTempCSVInfo.ToastHeader = $Header
     $ImportTempCSVInfo.ToastNotifications = $ToastNotifications
     $ImportTempCSVInfo | Export-Csv -Path "$($rootScriptFolder)\tempFinalInfo.csv" -NoTypeInformation
     
    
    if ((test-path $ScriptFolderLocation\logs.txt) -and -not (test-path $ScriptFolderLocation\errors.txt) -and -not (test-path $ScriptFolderLocation\warnings.txt))  {
    
        #script completed without errors or warning
        Write-Verbose "No Errors or Warnings. Script completed successfully"    
        send-log -scriptname $scriptname -rootScriptFolder $rootScriptFolder -logText "SCRIPT $($ScriptName) COMPLETED SUCCESSFULLY" -addDashes Below 
    
        if ($ToastNotifications -eq 'all'){  #send toast notification per Datto RMM variable

            Write-Verbose  "ToastNotifications Value is ALL"

            Invoke-AsCurrentUser {

                $ScriptName = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty ScriptName
                $rootScriptFolder = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty rootScriptFolder
                $ToastHeader = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty ToastHeader
                $userIsLoggedIn = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty UserLoggedIn
                $ToastNotifications = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty ToastNotifications
                $FolderForToastNotifications = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty FolderForToastNotifications

                #try{remove-item "$($rootScriptFolder)\tempFinalInfo.csv" -ErrorAction SilentlyContinue}catch{}    
       
               if ($userIsLoggedIn -eq "Yes"){ #skip notifications if user not logged in
                   if ($ToastNotifications -eq 'All'){ #alway push toast notifications
                       New-BurntToastNotification -Text "$($ToastHeader)","COMPLETED SUCCESSFULLY" -AppLogo "$($FolderForToastNotifications)\success.png" -UniqueIdentifier "$scriptname"
                   
                    }
               }
               
       
           }
        }
        
    
      #region Send teams notification
    
    $teamsMessageFile = get-content $ScriptFolderLocation\Hidden_Files\TeamsMessage.txt
     
    $JSONBody = [PSCustomObject][Ordered]@{
        "@type"      = "MessageCard"
        "@context"   = "http://schema.org/extensions"
        "summary"    = "$($ScriptName) - $env:computername"
        "themeColor" = '0078D7'
        "sections"   = @(
              @{
                  
             
                      
              "facts"            = @(
    
                        @{
                            "name"  = ""
                            "value" = "<strong style='color:#70B26C;'>$($env:COMPUTERNAME)</strong>"
                        },
                        @{
                            "name"  = ""
                            "value" = "<strong style='color:#2BB557;'>$($teamsMessageFile)</strong>"
                          },
                          @{
                            "name"  = ""
                            "value" = "<strong style='color:#3192E3;'>$($ScriptName)</strong>"
                          },
                        @{
                              "name"  = "Company:"
                              "value" = "$($Company)"
                            },
    
                        @{
                            "name"  = "Action:"
                            "value" = "$($action)"
                            }
                      )
    
                      "markdown" = $true
                  }
        )
      }
      
      $TeamMessageBody = ConvertTo-Json $JSONBody -Depth 100
      
      $parameters = @{
          "URI"         = "$env:DattoTeamsChannelWebhookURL"
          "Method"      = 'POST'
          "Body"        = $TeamMessageBody
          "ContentType" = 'application/json'
      }
      
    
    if ($SendToTeams.IsPresent){ #proceed if we specified we want to send to teams

        if ($SendFinalResultToTeams -eq 'ifsuccess' -or $SendFinalResultToTeams -eq 'yes'){
            
           Invoke-RestMethod @parameters | Out-Null
        }
    }
    #endregion Send teams notification
    
    
    }else{
    
        send-Log -logText "SCRIPT $($ScriptName) COMPLETED WITH ERRORS" -type Warning   -addDashes Above
    
        if ($ToastNotifications -eq 'Errors' -or $ToastNotifications -eq 'WarningsErrors' ){
    
            Invoke-AsCurrentUser {
               
                $ScriptName = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty ScriptName
                $rootScriptFolder = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty rootScriptFolder
                $ToastHeader = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty ToastHeader
                $userIsLoggedIn = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty UserLoggedIn
                $ToastNotifications = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty ToastNotifications
                $FolderForToastNotifications = import-csv c:\yw-data\automate\tempFinalInfo.csv | select-object -expandproperty FolderForToastNotifications

                try{remove-item "$($rootScriptFolder)\tempFinalInfo.csv" -ErrorAction SilentlyContinue}catch{}    
       
               if ($userIsLoggedIn -eq "Yes"){ #skip notifications if user not logged in
                   if ($ToastNotifications -eq 'All' -or $DattoToastNotificationVar -eq 'Errors' -or $DattoToastNotificationVar -eq 'WarningsErrors'){ 
                       New-BurntToastNotification -Text "$($ToastHeader)","COMPLETED WITH ERRORS/WARNINGS" -AppLogo "$($FolderForToastNotifications)\error.png" -UniqueIdentifier "$scriptname"
                   }
               }
       }
       
        }
      
     #region Send teams notification
     $teamsMessageFile = get-content $ScriptFolderLocation\Hidden_Files\TeamsMessage.txt
     
 
   
$JSONBody = [PSCustomObject][Ordered]@{
    "@type"      = "MessageCard"
    "@context"   = "http://schema.org/extensions"
    "summary"    = "$($ScriptName) - $env:computername"
    "themeColor" = '0078D7'
    "sections"   = @(
          @{
         
                  
            "facts"            = @(

            @{
                "name"  = ""
                "value" = "<strong style='color:#70B26C;'>$($env:COMPUTERNAME)</strong>"
            },
            @{
                "name"  = ""
                "value" = "<strong style='color:#D2395A;'>$($teamsMessageFile)</strong>"
              },
              @{
                "name"  = ""
                "value" = "<strong style='color:#3192E3;'>$($ScriptName)</strong>"
              },
            @{
                  "name"  = "Company:"
                  "value" = "$($Company)"
                },

            @{
                "name"  = "Action:"
                "value" = "$($action)"
                }
          )
          


                  "markdown" = $true
              }
    )
  }
      
      $TeamMessageBody = ConvertTo-Json $JSONBody -Depth 100
      
      $parameters = @{
          "URI"         = "$env:DattoTeamsChannelWebhookURL"
          "Method"      = 'POST'
          "Body"        = $TeamMessageBody
          "ContentType" = 'application/json'
      }
      
    
      if ($SendToTeams.IsPresent){ #proceed if we specified we want to send to teams

            if ($SendFinalResultToTeams -eq 'iferrors' -or $SendFinalResultToTeams -eq 'yes'){
                
                Invoke-RestMethod @parameters | Out-Null
            }
    }
    #endregion Send teams notification
    
    
    }
    
    }

function distribute-scriptExecution{
    write-host empty for now

    $NumberOfPCsToPush = $env:NumberOfPCsToPushTo

if($NumberOfPCsToPush -eq '10-50'){
    # $numberOfSec = Get-Random -Minimum 100 -Maximum 200 
    # send-Log -logText "Thorthling for $NumberOfPCsToPush PCs"
    # send-log -logText "Start sleep for $numberOfSec seconds"
    # prepare-YWToastNotification -ToastNotificationType Success -ToastNotificationText "Thorthling for $NumberOfPCsToPush PCs"
    # send-YWToastNotification
    # prepare-YWToastNotification Success -ToastNotificationText "Start sleep for $numberOfSec seconds"
    # send-YWToastNotification
    # Start-Sleep $numberOfSec

}elseif($NumberOfPCsToPush -eq '50-100'){

    # $numberOfSec = Get-Random -Minimum 200 -Maximum 400
    # send-Log -logText "Thorthling for $NumberOfPCsToPush PCs"
    # send-log -logText "Start sleep for $numberOfSec seconds"
    # prepare-YWToastNotification -ToastNotificationType Success -ToastNotificationText "Thorthling for $NumberOfPCsToPush PCs"
    # send-YWToastNotification
    # prepare-YWToastNotification Success -ToastNotificationText "Start sleep for $numberOfSec seconds"
    # send-YWToastNotification
    # Start-Sleep $numberOfSec

}elseif($NumberOfPCsToPush -eq '100-300'){
    # $numberOfSec = Get-Random -Minimum 400 -Maximum 800
    # send-Log -logText "Thorthling for $NumberOfPCsToPush PCs"
    # send-log -logText "Start sleep for $numberOfSec seconds"
    # prepare-YWToastNotification -ToastNotificationType Success -ToastNotificationText "Thorthling for $NumberOfPCsToPush PCs"
    # send-YWToastNotification
    # prepare-YWToastNotification Success -ToastNotificationText "Start sleep for $numberOfSec seconds"
    # send-YWToastNotification
    # Start-Sleep $numberOfSec
}elseif ($NumberOfPCsToPush -eq '300-500') {
    # $numberOfSec = Get-Random -Minimum 800 -Maximum 1600
    # send-Log -logText "Thorthling for $NumberOfPCsToPush PCs"
    # send-log -logText "Start sleep for $numberOfSec seconds"
    # prepare-YWToastNotification -ToastNotificationType Success -ToastNotificationText "Thorthling for $NumberOfPCsToPush PCs"
    # send-YWToastNotification
    # prepare-YWToastNotification Success -ToastNotificationText "Start sleep for $numberOfSec seconds"
    # send-YWToastNotification
    # Start-Sleep $numberOfSec
}elseif($NumberOfPCsToPush -eq '500+'){
    # $numberOfSec = Get-Random -Minimum 1600 -Maximum 3200
    # send-Log -logText "Thorthling for $NumberOfPCsToPush PCs"
    # send-log -logText "Start sleep for $numberOfSec seconds"
    # prepare-YWToastNotification -ToastNotificationType Success -ToastNotificationText "Thorthling for $NumberOfPCsToPush PCs"
    # send-YWToastNotification
    # prepare-YWToastNotification Success -ToastNotificationText "Start sleep for $numberOfSec seconds"
    # send-YWToastNotification
    # Start-Sleep $numberOfSec
}
}



function check-softwarePresence{

    <#
       .SYNOPSIS
           Check if software is installed
       .DESCRIPTION
            Check if software is installed using get-package, registry and chocolatey if specified and return PS object
            For function to work properly, you need to provide these variables either in script or global scrope or as Datto global variable 
               $scriptName
               $rootScriptFolder
       .PARAMETER scriptName
            It is a script name that we use to create a folder for this script in root scripts working folder
       .PARAMETER rootScriptFolder
            It is a root folder for all scripts. E.g. c:\automations. It shold be full path.
        .PARAMETER SoftwareName
            Name of the software. If exactNameMatch is specified, it should be the exact name of the software. Otherwise it should be part of the name.
       .PARAMETER ExactNametMatch
            If specified, it will check if exact name of the software is installed
       .PARAMETER includeChoco
            If specified, it will check if software is installed via chocolatey
       .PARAMETER chocolateyName
            Name of the software in chocolatey. This parameter is required if includeChoco is specified
       .EXAMPLE
           check-softwarePresence -softwareName "Google Chrome" -ExactNametMatch
           check-softwarePresence -softwareName "Google Chrome" -includeChoco -chocolateyName "googlechrome" 
       .OUTPUTS
       .NOTES
           FunctionName : 
           Created by   : Sasa Zelic
           Date Coded   : 12/2019
    #>

   [CmdletBinding()]
       param(
           [string]$SoftwareName,
           [switch]$ExactNametMatch,
           [switch]$includeChoco,
           [string]$ChocolateyName,
           [string]$rootScriptFolder = $rootScriptFolder,
           [string]$scriptname = $scriptname
           
       )

       if (-not $rootScriptFolder){
            $rootScriptFolder = $env:rootScriptFolder
        }

    if ($rootScriptFolder[-1] -like '\'){$rootScriptFolder = $rootScriptFolder.Substring(0, $rootScriptFolder.Length - 1)}else{$rootScriptFolder = $rootScriptFolder}

    #create script folder if it doesn't exist
    if (-not (test-path "$rootScriptFolder\$scriptName")){New-Item -Path "$rootScriptFolder" -Name "$scriptName" -ItemType Directory -Force -ErrorAction Stop | out-null}



       if ($includeChoco.IsPresent -and $ChocolateyName){

                try{

                
                    $chocoSoftwareCheck = (choco list) | select-string $ChocolateyName
                    
                    if($chocoSoftwareCheck){$VersionViaChoco = $chocoSoftwareCheck.Line.Split(" ")[1]}else{$VersionViaChoco = ""}

                    $UpdateNeededCheckViaChoco = choco outdated -r | select-string $ChocolateyName # switch -r means to limit the output to essential information only
                
 
                    if($UpdateNeededCheckViaChoco){$UpdatedNeededViaChoco = $true}else{$UpdatedNeededViaChoco = $false}
                    if($chocoSoftwareCheck){$InstalledViaChoco = $true}else{$InstalledViaChoco = $false}

                }catch{
                    send-Log -logText "Failed to check if $($Softwarename) is installed. Exiting script" -addDashes Below -type Error   -catch
                    exit 1
        
                }
           
        }else{

            $VersionViaChoco = ""
            $UpdatedNeededViaChoco = ""
            $InstalledViaChoco = ""
        }


        try{
            #via package
            if ($ExactNametMatch.IsPresent){$softwareName = $SoftwareName}else{$softwareName = "*$SoftwareName*"}  
               
                $softwareCheck = Get-Package
                $softwareCheck = $softwareCheck | Where-Object { $_.Name -like "$SoftwareName" }

                if($softwareCheck){

                    $softwareVersion = $softwareCheck.Version
                    
                }else{
                   #check via registry

                   # registry locations where installed software is logged
                   $pathAllUser = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
                   $pathCurrentUser = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
                   $pathAllUser32 = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
                   $pathCurrentUser32 = "Registry::HKEY_CURRENT_USER\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
      
                   # get all values
                   $softwareCheck = (Get-ItemProperty -Path $pathAllUser, $pathCurrentUser, $pathAllUser32, $pathCurrentUser32 |
                       # skip all values w/o displayname
                       Where-Object DisplayName -ne $null |
                       # apply user filters submitted via parameter:
                       Where-Object DisplayName -like $SoftwareName)

                    $softwareVersion = $softwareCheck.displayversion

               }

        }catch{
            send-Log -logText "Failed to check if $($Softwarename) is installed. Exiting script" -addDashes Below -type Error    -catch
            exit 1
        }



$finalResult =  [PSCustomObject]@{
        Installed = if ($softwareCheck){$true}else{$False}
        Version =  $softwareVersion
        InstalledviaChoco = $InstalledViaChoco
        VersionViaChoco =  $VersionViaChoco
        UpdatedNeededViaChoco = $UpdatedNeededViaChoco
        
      }

      return $finalResult
   
}


function remove-oldToastNotifications{

 <#
       .SYNOPSIS
           Remove old toast notifications related to this script
       .DESCRIPTION
           Remove old toast notifications related to this script
               $scriptName
       .PARAMETER scriptName
            It is a script name that is used to create uniqueToastNotification
       .EXAMPLE
            remove-oldToastNotifications -scriptname $scriptName
        .OUTPUTS
       .NOTES
           FunctionName : 
           Created by   : Sasa Zelic
           Date Coded   : 12/2019
    #>

    [CmdletBinding()]
    param(
        [string]$scriptname = $scriptname
        
    )
    
$ifUserLoggedInCheck  = (Get-WmiObject -ClassName Win32_ComputerSystem).Username

[int]$counterUniqueIdentifier = 0
if($ifUserLoggedInCheck){

    remove-BTNotification -group "$scriptName" -ErrorAction SilentlyContinue #main one

    1.15 | ForEach-Object {

        $counterUniqueIdentifier++

        $OldIdentifier = "$scriptName" + '---' + "$counterUniqueIdentifier"
        
        Invoke-AsCurrentUser {
            remove-BTNotification -group "$OldIdentifier" 
    
        }
    }
}
}