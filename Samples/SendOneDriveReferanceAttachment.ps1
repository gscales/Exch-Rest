Import-Module Exch-Rest -Force
$MailboxName = "gscales@datarumble.com"
$OneDriveFile = "/test/test2/fileName.zip"
$DownloadDirectory = "c:\temp"
##Get the Access Token
$AccessToken =  Get-AccessToken -MailboxName $MailboxName  -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl "urn:ietf:wg:oauth:2.0:oob" -ResourceURL graph.microsoft.com -Beta 
##Get OneDrive DownloadURI
$OneDriveAttachmentToSend = Get-OneDriveItemFromPath -MailboxName $MailboxName -AccessToken $AccessToken -OneDriveFilePath $OneDriveFile
$rtArray = @()
$rtArray += (New-referanceAttachment -Name $OneDriveAttachmentToSend.Name -SourceUrl $OneDriveAttachmentToSend.webUrl -Permission Edit)
Send-MessageREST -MailboxName $MailboxName  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address glenscales@yahoo.com) -Subject "Daily Send" -Body "See Attached" -ReferanceAttachments $rtArray