Import-Module Exch-Rest -Force
$MailboxName = "gscales@datarumble.com"
$OneDriveFile = "/test/test2/fileName.zip"
$DownloadDirectory = "c:\temp"
##Get the Access Token
$AccessToken =  Get-EXRAccessToken -MailboxName $MailboxName  -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl "urn:ietf:wg:oauth:2.0:oob" -ResourceURL graph.microsoft.com -Beta 
##Get OneDrive DownloadURI
$OneDriveAttachmentToSend = Get-EXROneDriveItemFromPath -MailboxName $MailboxName -AccessToken $AccessToken -OneDriveFilePath $OneDriveFile
$rtArray = @()
$rtArray += (New-EXRReferanceAttachment -Name $OneDriveAttachmentToSend.Name -SourceUrl $OneDriveAttachmentToSend.webUrl -Permission Edit)
Send-EXRMessageREST -MailboxName $MailboxName  -AccessToken $AccessToken -ToRecipients @(New-EXREmailAddress -Address user@domain.com) -Subject "Daily Send" -Body "See Attached" -ReferanceAttachments $rtArray