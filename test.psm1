

function store {

    [cmdletbinding()]
    param(
        
        [string]$variableName = $globalsometing
    )
    
    write-host $variableName
    write-host "testing"
    write-host $env:rootScriptFolder
    
    
}