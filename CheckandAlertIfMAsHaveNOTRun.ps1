# parsing XML syntax stolen from https://practical365.com/exchange-server/using-xml-settings-file-powershell-scripts/
[xml]$XMLConfig = Get-Content .\Config.xml ## this has settings in it, so we're not harcoding values in scripts

<#
.SYNOPSIS
    This test function will check MIM server run history to see if an MA has been run in a number of hours.  The idea is the his function can be used to ensure the expected MIM schedule is actually running.  
.DESCRIPTION
    This function requires the lithnet PS module for FIM/MIM named LithnetMiisAutomation
.EXAMPLE
    PS C:\> Test-HasMARunInLastXHours -MAName "HR SQL MA"
    Checking the HR SQL MA to see if it has been run in the last default 24 hours 
.EXAMPLE
    PS C:\> 'BNYM FIM MA' | Test-HasMARunInLastXHours
    Checking the HR SQL MA to see if it has been run in the last default 24 hours.  
    Using the option to pipe the name value to the cmdlet
.EXAMPLE
    PS C:\> Test-HasMARunInLastXHours -MAName AD -XHoursAgo 10 
    Checking the AD MA to see if it has been run in the last 10 hours 
.EXAMPLE
    PS C:\> Test-HasMARunInLastXHours -MAName "HR SQL MA" -XHoursAgo 36
    Checking the HR SQL MA to see if it has been run in the last 36 hours 
.EXAMPLE
    PS C:\> Test-HasMARunInLastXHours -MAName "AD" -XHoursAgo 10 -Verbose
    Checking the AD to MA to see if it has been run in the last 10 hours.  
    showing verbose information, used for t-shooting
#>
function Test-HasMARunInLastXHours {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true )]
        [string]$MAName,
        [int]$XHoursAgo = 24
    )
    $XHoursAgoTime = (Get-Date).AddHours(-$XHoursAgo)
    Write-Verbose "Checking $MAName for Xhours: $XHours or after $XHoursAgoTime"
    ## TODO:.  put a try/catch out this in case IT gets an error.  e.g. what if svc is stopped
    $LastRunDetails = Get-LastRunDetails -MA $MAName 
    ## notes about this query:  
    # it does NOT pull back stats of a currently running run, it pulls the last completed run.  
    # if there is NO run history for this MA, 
    if ($Null -ne $LastRunDetails) {
        Write-Verbose "$MAName HAS history of running on server."
        ## now checking if it started in last X hours
        $LastRunStartTime = $LastRunDetails.StartTime
        Write-Verbose "This is the start time of the last time this MA was run: $LastRunStartTime"
        if ($LastRunStartTime -gt $XHoursAgoTime ) {
            ## this means a run has been within the tested number of hours
            Write-Verbose "Checking if $LastRunStartTime is after $XHoursAgoTime"
            return $true 
        }
        else {
            ## this means a run has NOT been within the tested number of hours
            Write-Verbose "Checking if $LastRunStartTime is after $XHoursAgoTime"
            return $false 
        }
    }
    else {
        Write-Verbose "$MAName has NO history of running on server."
        return $false 
    }
    Write-Verbose "End of Check"
}


function Test-AllMAsandSendAlertIfNotCurrent {
    [CmdletBinding()]
    param (
        
    )
    foreach ($MA in Get-ManagementAgent) {
        $MAName = $MA.Name 
        Write-Verbose "Checking $MAName"
        if (Test-HasMARunInLastXHours -MAName $($MA.Name) ) { 
            Write-Verbose "$MAName passed the test, nothing to see here.  move along.  these are not droids you are looking for."
        }
        else {
            Write-Verbose "$MAName FAILED the test, sending alert"
            Send-MIMInfo -SubjectPrefix "MIM Alert: $MAName MA has NOT been run in the expected timeframe.  Please review" 
        }

    }
    
    
}

function Send-MIMInfo {
    [CmdletBinding()]
    param (
        $MailRecipients = $XMLConfig.Settings.Email.Alert1.to ,
        $smtpserver = $XMLConfig.Settings.Email.smtpserver,
        $mailFromAddress = $XMLConfig.Settings.Email.from, 
        $Message = '.', 
        $SubjectPrefix = "MIM Run Complete:"
    )
    Send-MailMessage -SmtpServer $smtpserver -Subject "$SubjectPrefix $($ENV:COMPUTERNAME)" `
        -To $MailRecipients -From $mailFromAddress -Body $Message -BodyAsHtml 


}

Test-AllMAsandSendAlertIfNotCurrent

