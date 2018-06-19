$To = "gscales@datarumble.com"
Connect-EXRMailbox -MailboxName "gscales@datarumble.com"
Connect-EXRManagementAPI -UserName "gscales@datarumble.com"
$StatusArray = Get-EXRMCurrentStatus | Select-Object WorkLoad,Statusdisplayname,StatusTime
$ColorSwitchHash = @{}
$ColorSwitchHash.Add("Service degradation","attention")
$ColorSwitchHash.Add("Normal service","good")
$Card = New-EXRAdaptiveCard -Columns $StatusArray -ColorSwitchColumnNumber 1 -ColorSwitchHashTable $ColorSwitchHash -ColorSwitchDefault warning
Send-EXRAdaptiveCard -To $To -Subject "Office 365 Service Status" -AdaptiveCard $Card