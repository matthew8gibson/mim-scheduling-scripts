#requires -module LithnetMiisAutomation
## this to go generate alerts when issues occur
$ErrorActionPreference = 'Stop'
<#
The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
#>

# parsing XML syntax stolen from https://practical365.com/exchange-server/using-xml-settings-file-powershell-scripts/
[xml]$XMLConfig = Get-Content .\Config.xml ## this has settings in it, so we're not harcoding values in scripts

<#
.SYNOPSIS
    This function checks the number of pending imports and/or exports of each type on an MA and compares it to the corresponding threshold. 
.DESCRIPTION
    
.EXAMPLE
    PS C:\> Test-MAisPendingThresholdExceeded -MA 'Active Directory' -ImportDeleteThreshold 3 
    This command is checking the Pending Import Delete value on the Management Agent named Active Directory.   
      If there are 2 Pending Import Deletes, the returne value will be $false because the threshold IS NOT exceeded. 
      If there are 3 Pending Import Deletes, the returne value will be $false because the threshold IS NOT exceeded. 
      If there are 4 Pending Import Deletes, the returne value will be $true because the threshold IS exceeded.  
    PS C:\> Test-MAisPendingThresholdExceeded -MA 'Active Directory' -ImportDeleteThreshold 75 -ImportAddThreshold 150 -ImportUpdateThreshold 500
    This command is checking 3 different parameters for the same Management Agent.  If ANY of the Thresholds are exceeded, then the function will return $true.  
.NOTES
    These are all the test use cases
#>


function Test-MAisPendingThresholdExceeded {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $MA, 
        $ImportDeleteThreshold, $ImportUpdateThreshold, $ImportAddThreshold,
        $ExportDeleteThreshold, $ExportUpdateThreshold, $ExportAddThreshold
    )
    # import-module LithnetMiisAutomation
    ## BEGIN Gathering Stats
    $MAStats = Get-MAStatistics -MA $MA 
    Write-Verbose "PendingImportDeletes: $(($MAStats).ImportDeletes)"
    Write-Verbose "PendingImportUpdates: $(($MAStats).ImportUpdates)"
    Write-Verbose "PendingImportAdds: $(($MAStats).ImportAdds)"

    Write-Verbose "PendingExportDeletes: $(($MAStats).ExportDeletes)"
    Write-Verbose "PendingExportUpdates: $(($MAStats).ExportUpdates)"
    Write-Verbose "PendingExportAdds: $(($MAStats).ExportAdds)"
    ## END Gathering Stats
    ## BEGIN Pending Import Checks
    if ($null -ne $ImportDeleteThreshold) { if ($ImportDeleteThreshold.ToInt32($null) -lt $MAStats.ImportDeletes) { Write-Verbose "importDeleteExceeded"; return $true } }
    if ($null -ne $ImportUpdateThreshold) { if ($ImportUpdateThreshold.ToInt32($null) -lt $MAStats.ImportUpdates) { Write-Verbose "importUpdateExceeded"; return $true } }
    if ($null -ne $ImportAddThreshold) { if ($ImportAddThreshold.ToInt32($null) -lt $MAStats.ImportAdds ) { Write-Verbose "importAddExceeded"; return $true } }
    ## END Pending Import Checks

    ## BEGIN Pending Export Checks
    if ($null -ne $ExportDeleteThreshold) { if ($ExportDeleteThreshold.ToInt32($null) -lt $MAStats.ExportDeletes ) { Write-Verbose "ExportDeleteExceeded"; return $true } }
    if ($null -ne $ExportUpdateThreshold) { if ($ExportUpdateThreshold.ToInt32($null) -lt $MAStats.ExportUpdates) { Write-Verbose "ExportUpdateExceeded"; return $true } }
    if ($null -ne $ExportAddThreshold) { if ($ExportAddThreshold.ToInt32($null) -lt $MAStats.ExportAdds ) { Write-Verbose "ExportDeleteExceeded"; return $true } }
    ## END Export Checks

    if ( ## all parameters are blank
        $null -eq $ImportDeleteThreshold -and 
        $null -eq $ImportUpdateThreshold -and 
        $null -eq $ImportAddThreshold -and 
        $null -eq $ExportDeleteThreshold -and 
        $null -eq $ExportUpdateThreshold -and 
        $null -eq $ExportAddThreshold
    ) {
        Write-Verbose 'All Threshold Parameters are blank'
        return $true
    }
    return $false # last line, if none others triggered then 
}


function Invoke-RunProfile {
    [CmdletBinding()]
    param (
        $MAName = 'AD', 
        $RunProfileName = 'DI'
    )
    $output = $null
    $output = New-Object -TypeName psobject
    Add-Member -InputObject $output -MemberType NoteProperty -Name MAName -Value $MAName
    Add-Member -InputObject $output -MemberType NoteProperty -Name RunProfile -Value $RunProfileName 
    try {
        Start-ManagementAgent -MA $MAName -RunProfileName $RunProfileName 
        $RS = Get-RunSummary -MA $MAName | Sort-Object STartTime -Descending | Select-Object -First 1
        $RD = Get-RunDetail -MA $MAName -RunNumber $($RS.RunNumber)
    }
    catch {
        Send-MIMErrorAlert -objError $Error 
        Throw 
    }
    Add-Member -InputObject $output -MemberType NoteProperty -Name LastStepStatus -Value $RD.LastStepStatus
    Add-Member -InputObject $output -MemberType NoteProperty -Name StartTime -Value $RD.StartTime
    Add-Member -InputObject $output -MemberType NoteProperty -Name EndTime -Value $RD.EndTime
    Add-Member -InputObject $output -MemberType NoteProperty -Name Duration -Value $("{0:hh}:{0:mm}:{0:ss}" -f ($RD.EndTime - $RD.StartTime))
    $output
}  
function Invoke-ScheduledRun {
    [CmdletBinding()]
    param (
        $Runtype = 'Delta' # Full or Delta
    )
    $Error.Clear() 
    $ErrorActionPreference = 'Stop' ## this to go generate alerts 
    ## Main
    $RunResultDetails = @() 
    switch ($Runtype) {
        Delta {
            $RunResultDetails += Invoke-DeltaRunProfiles 
            Clear-RunHistory -DaysToKeep 30 # this is a Lithnet PS module function to remove Sync Server history.  
        }
        Full { 
            $RunResultDetails += Invoke-FullRunProfiles
        }
        Default {
            'no valid case found.  Should have been Delta or Full'
            exit
        }
    }
    # this sends an email with summary of the run if the run completes 
    # comment it out, if no email is desired
    Send-MIMINfo -Message $($RunResultDetails | ConvertTo-Html -Property MAName, RunProfile, LastStepStatus, StartTime, EndTime, Duration | Out-String) -SubjectPrefix "MIM Run Complete:"
}

function Invoke-FullRunProfiles {
    # just getting a Full Import and Full Sync on all MA's.  
    # this will ensure any configuration changes have taken effect
    try {
        $Results = @()
        foreach ($MA in $(Get-ManagementAgent) ) {
            $Results += Invoke-RunProfile -MAName $MA -RunProfileName "Full Import"   
            $Results += Invoke-RunProfile -MAName $MA -RunProfileName "Full Sync"    
        }   ## end of foreach MA
        return $Results
    }
    catch {
        Send-MIMErrorAlert -objError $Error 
        Throw 
    }
}


function Invoke-DeltaRunProfiles {
    [CmdletBinding()]
    param (
        
    )
    ## Getting fresh data from each system with DI. 
    try {
        $ResultsDelta = @()

        # bringing in latest Azure Data 
        $ResultsDelta += Invoke-RunProfile -MAName "B2B Graph MA" -RunProfileName "Delta Import" 
        $ResultsDelta += Invoke-RunProfile -MAName "B2B Graph MA" -RunProfileName "Delta Sync" 

        ## getting latest updates from AD
        $ResultsDelta += Invoke-RunProfile -MAName "AD MA" -RunProfileName "Delta Import" 
        ## AD MA import and Checking before doing the  Syncing
        if (Test-MAisPendingThresholdExceeded -MA "AD MA" -ImportDeleteThreshold $XMLConfig.Settings.Thresholds.ADMA.Import.Delete -ImportUpdateThreshold $XMLConfig.Settings.Thresholds.ADMA.Import.Update -ImportAddThreshold $XMLConfig.Settings.Thresholds.ADMA.Import.Add) {
            Write-Verbose "Review! - AD MA Database Threshold Exceeded"
            Send-MIMInfo -SubjectPrefix 'Review! - AD MA Database Threshold Exceeded'
            return "Review! - AD MA Database Threshold Exceeded"
        }
        else {
            Write-Verbose "AD MA Database Threshold NOT Exceeded"
        }

        # sending data to the MIM MA and getting back updates
        $ResultsDelta += Invoke-RunProfile -MAName "MIM MA" -RunProfileName "Export" 
        Start-Sleep -Seconds 65 ## letting the MIM Service process updates.  
        $ResultsDelta += Invoke-RunProfile -MAName "MIM MA" -RunProfileName "Delta Import" 
        $ResultsDelta += Invoke-RunProfile -MAName "MIM MA" -RunProfileName "Delta Sync" 
        
       ## Azure Export and Checking Pending Export Thresholds before doing the export
        if (Test-MAisPendingThresholdExceeded -MA "B2B Graph MA" -ExportDeleteThreshold $XMLConfig.Settings.Thresholds.B2BGraphMA.Export.Delete -ExportUpdateThreshold $XMLConfig.Settings.Thresholds.B2BGraphMA.Export.Update -ExportAddThreshold $XMLConfig.Settings.Thresholds.B2BGraphMA.Export.Add) {
            Write-Verbose "Review! - B2BGraphMA Threshold Exceeded"
            Send-MIMInfo -SubjectPrefix 'Review! - B2BGraphMA Threshold Exceeded'
            return "Review! - B2BGraphMA Threshold Exceeded"
        }
        else {
            # doing the export b/c the threshold was NOT exceeded
            $ResultsDelta += Invoke-RunProfile -MAName "B2B Graph MA" -RunProfileName "Export" 
            $ResultsDelta += Invoke-RunProfile -MAName "B2B Graph MA" -RunProfileName "Delta Import" 
            $ResultsDelta += Invoke-RunProfile -MAName "B2B Graph MA" -RunProfileName "Delta Sync"  
        }
        
        return $ResultsDelta
    }
    catch {
        Send-MIMErrorAlert -objError $Error 
        Throw 
    }
}

function Send-MIMErrorAlert {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $objError, 
        $MailRecipients = $XMLConfig.Settings.Email.Alert1.to ,
        $smtpserver = $XMLConfig.Settings.Email.smtpserver,
        $mailFromAddress = $XMLConfig.Settings.Email.from
    )
    If ($null -eq $MailRecipients) { $MailRecipients = $global:AlertMailRecipients }
    $MessageBody = 'Error Details: ' + "`n"
    $MessageBody += $objError[0].ToString() + "`n"
    $MessageBody += "CategoryInfo: " + $objError[0].CategoryInfo.ToString() + "`n" + "`n"
    $MessageBody += "ScriptStackTrace: " + $objError[0].ScriptStackTrace.ToString() + "`n" + "`n"
    Send-MailMessage -SmtpServer $smtpserver -Subject "MIM Run Error:  $($ENV:COMPUTERNAME)" -To  $MailRecipients -From $mailFromAddress -Body $MessageBody 

}

function Send-MIMInfo {
    [CmdletBinding()]
    param (
        $MailRecipients = $XMLConfig.Settings.Email.Alert1.to ,
        $smtpserver = $XMLConfig.Settings.Email.smtpserver,
        $bcc = $XMLConfig.Settings.Email.Alert1.bcc,
        $mailFromAddress = $XMLConfig.Settings.Email.from, 
        $Message = '.', 
        $SubjectPrefix = "MIM Run Complete:"
    )
    Send-MailMessage -SmtpServer $smtpserver -Subject "$SubjectPrefix $($ENV:COMPUTERNAME)" `
        -To $MailRecipients -From $mailFromAddress -Body $Message -BodyAsHtml -Bcc $bcc

}

