Function get-SoftwareVersionViaMSIFile {


     <#
       .SYNOPSIS
           Get software version via MSI file
       .DESCRIPTION
           There are some cases where we can't full software information using get-itemproperty of the MSI file.
           In that case, we have to pull the information from MSI database.
       .PARAMETER $msiPath
           Full path of the MSI file
       
       .EXAMPLE
           get-softwareVersionViaMSIFile -msiPath "C:\temp\test.msi" 
       .OUTPUTS
       .NOTES
           FunctionName : 
           Created by   : Sasa Zelic
           Date Coded   : 12/2019
    #>

    param (
    [parameter(Mandatory=$true)] 
    [ValidateNotNullOrEmpty()] 
    [System.IO.FileInfo] $MSIPATH
) 
if (!(Test-Path $MSIPATH.FullName)) { 
    throw "File '{0}' does not exist" -f $MSIPATH.FullName 
} 
try { 
    $WindowsInstaller = New-Object -com WindowsInstaller.Installer 
    $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($MSIPATH.FullName, 0)) 
    $Query = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
    $View = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $Database, ($Query)) 
    $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) | Out-Null
    $Record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null ) 
    $Version = $Record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $Record, 1 ) 
    return $Version
} catch { 
    throw "Failed to get MSI file version: {0}." -f $_

}

}