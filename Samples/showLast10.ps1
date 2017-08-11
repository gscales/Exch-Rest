$MailboxName = "gscales@datarumble.com"
$AccessToken = Get-AccessToken -MailboxName $MailboxName -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL graph.microsoft.com
$FolderSelected = Invoke-MailFolderPicker -MailboxName $MailboxName -AccessToken $AccessToken -ShowFolderSize
Get-FolderItems -AccessToken $AccessToken -MailboxName $MailboxName -Folder $FolderSelected -Top 10 -TopOnly:$true -ReturnSize | Select Subject,
receivedDateTime,singleValueExtendedProperties