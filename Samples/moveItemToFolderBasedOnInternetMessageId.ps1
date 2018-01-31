$MailboxName = "gscales@datarumble.com"
$MessageId = "<374020692.47135161.1517346591657.JavaMail.root@domain.com>"
$TargetFolder = "\Inbox\aa"
Connect-EXRMailbox -MailboxName $MailboxName
Find-EXRMessageFromMessageId -MailboxName $MailboxName -MessageId $MessageId | ForEach-Object{
    Move-EXRMessage -MailboxName $MailboxName -ItemURI $_.ItemRESTURI -TargetFolderPath $TargetFolder
    Write-Host ("Moved Message " + $_.Subject + " to " + $TargetFolder)
}
