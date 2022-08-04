Function Get-DurationString {
    Param(
        [Parameter(Mandatory = $true)]
        [DateTime]$Starttime,
        [Parameter(Mandatory = $true)]
        [string]$Section,
        [Parameter(Mandatory = $true)]
        [ValidateSet('TotalHours','TotalDays','TotalMinutes','TotalSeconds','TotalMilliSeconds')]
        [String]$TimeSelector,
        [Switch]$IncludeTime
    )
        switch($TimeSelector)
        {
            'TotalHours' {$TimeSelected = 'Hours'}
            'TotalDays' {$TimeSelected = 'Days'}
            'TotalMinutes' {$TimeSelected = 'Minutes'}
            'TotalSeconds' {$TimeSelected = 'Seconds'}
            'TotalMilliSeconds' {$TimeSelected = 'MilliSeconds'}
        }        
    $Duration = [Math]::Round(((Get-Date) - $Starttime).$timeSelector)
    if($IncludeTime.IsPresent) {
    "[$(Get-Date -Format G)][$Section] Completed in  $Duration $TimeSelected."
    } else {
        "[$Section] Completed in $Duration $TimeSelected."
    }
}
Function Write-Log {

    [CmdletBinding()]
    Param(
    
    
    [Parameter(Mandatory = $True)]
    [string]$Message,
    [string]$LogFilePath = "$($env:TEMP)\log_$((New-Guid).Guid).txt",
    [Switch]$DoNotRotateDaily
    )
    
    if ($DoNotRotateDaily) {

        
        $LogFilePath = if ($Script:LogFilePath) {$Script:LogFilePath} else {$LogFilePath}
            
    } else {
        if ($Script:LogFilePath) {

        $LogFilePath = $Script:LogFilePath
        $DayStamp = (Get-Date -Format 'yMMdd').Tostring()
        $Extension = ($LogFilePath -split '\.')[-1]
        $LogFilePath -match "(?<Main>.+)\.$extension`$" | Out-Null
        $LogFilePath = "$($Matches.Main)_$DayStamp.$Extension"
        
    } else {$LogFilePath}
    }
    $Log = "[$(Get-Date -Format G)][$((Get-PSCallStack)[1].Command)] $Message"
    
    Write-Verbose $Log
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

Function Get-StringHash {

    [CmdletBinding()]
    Param(
    
        [Parameter(Mandatory =$true,ValueFromPipeLine = $true)]
        [string]$String
    
    )

Process {

    $md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $utf8 = new-object -TypeName System.Text.UTF8Encoding
    $hash = [System.BitConverter]::ToString($md5.ComputeHash($utf8.GetBytes($String)))
    $hash -replace '-',''

}

}
<#
.SYNOPSIS
    Gets time difference of given datetime from now.
.DESCRIPTION
    Gets time difference of given datetime from now.
.EXAMPLE
        Get-TimeSpan -Time $StartTime -Span TotalHours  
        
        1.56523913183333
#>
Function Get-TimeSpan {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [DateTime]$Time,
        [ValidateSet('Days','Hours','MilliSeconds','Minutes','Seconds','TotalDays','TotalHours','TotalMilliSeconds','TotalMinutes','TotalSeconds')]
        [string]$Span='TotalHours'
    )
    
    ((Get-Date) - $Time).$Span
}
<#
.SYNOPSIS
    Converts the datetime to cimdatetime to be used in WMI and CIM.
.DESCRIPTION
    Converts the datetime to cimdatetime to be used in WMI and CIM.
    
    HighLEvel Steps

    1) Get String with UTC ofsset pattern. Ie: 20200408115224.000000+03:00 while doing so get the UTCSign and the UTCHour
    2) Split the the hour and make calculatations to convert to 3 digit minutes
    3) Replace the +3:00 with the calculated 180
.EXAMPLE
        Get-CIMDateTime
        
#>   
Function Get-CIMDateTime {


$CimDateString= get-date -Format "yyyyMMddHHmmss.000000K"
if ($CimDateString -match '(?<UTCSign>\+|-)(?<UTC>.+)')

{

$UTCArray = $Matches['UTC'] -split ':'

$UTCMinutes =  "{0:d3}" -f  ([int]$UTCArray[0] *60 + [int]$UTCArray[1])

$CimDateString -replace '\+(.+)' ,"$($Matches['UTCSign'])$UTCMinutes"
}

}

