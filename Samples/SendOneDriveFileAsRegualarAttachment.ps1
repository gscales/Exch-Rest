$MailboxName = "gscales@datarumble.com"
$OneDriveFile = "/test/test2/FileName.zip"
$DownloadDirectory = "c:\temp"
##Get the Access Token
$AccessToken =  Get-EXRAccessToken -MailboxName $MailboxName  -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl "urn:ietf:wg:oauth:2.0:oob" -ResourceURL graph.microsoft.com 
##Get OneDrive DownloadURI
$OneDriveAttachmentToSend = Get-EXROneDriveItemFromPath -MailboxName $MailboxName -AccessToken $AccessToken -OneDriveFilePath $OneDriveFile
$DownloadFileName = $DownloadDirectory + "\" + $OneDriveAttachmentToSend.Name
Invoke-WebRequest -Uri $OneDriveAttachmentToSend.'@microsoft.graph.downloadUrl' -OutFile $DownloadFileName
Send-EXRMessageREST -MailboxName $MailboxName  -AccessToken $AccessToken -ToRecipients @(New-EXREmailAddress -Address user@domain.com) -Subject "Daily Send" -Body "See Attached" -Attachments @($DownloadFileName)