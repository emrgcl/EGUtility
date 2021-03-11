Function Write-Log {

    [CmdletBinding()]
    Param(
    
    
    [Parameter(Mandatory = $True)]
    [string]$Message,
    [string]$LogFilePath = "$($env:TEMP)\log_$((New-Guid).Guid).txt"
    
    )
    
    $LogFilePath = if ($Script:LogFilePath) {$Script:LogFilePath} else {$LogFilePath}
    
    $Log = "[$(Get-Date -Format G)][$((Get-PSCallStack)[1].Command)] $Message"
    
    Write-verbose $Log
    $Log | Out-File -FilePath $LogFilePath -Append -Force
    
}
Function New-SQLConnectionString {
    [CmdletBinding()]
    Param(
        # Parameter help description
        [Parameter(Mandatory = $True)]
        $QueryInfo
    )
    $ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=$($QueryInfo.DBName);Data Source=$($QueryInfo.SQLServer)\$($Queryinfo.SQLInstance),$($QueryInfo.Port)"

    Write-Log "Created Connection string '$ConnectionString'"
    $ConnectionString
}

Function ConvertTo-Int32 {

[CmdletBinding()]
Param(

[Parameter(ValueFromPipeLine=$true)]
[PSCustomObject]$InputObject

)

Process {

$ConvertedProperties = @{}
$Properties = ($_ | Get-Member -MemberType NoteProperty).Name
Foreach ($Property in $Properties) {

if ($_.$Property -as [int32]) {

$ConvertedProperties.Add($Property,($_.$Property -as [int32]))

} else {

$ConvertedProperties.Add($Property,($_.$Property))

}


}

[PsCustomObject]$ConvertedProperties

}


}
