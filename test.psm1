

function store {

    [cmdletbinding()]
    param(
        
        [string]$variableName = $globalsometing
    )
    
    write-host $variableName
    write-host $globalsometing1
    write-host $globalsometing2
    write-host $globalsometing3
    write-host $env:rootScriptFolder
    
    
    
    
}
