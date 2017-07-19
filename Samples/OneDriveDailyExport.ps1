#Import-Module Exch-Rest -Force
$MailboxName = "gscales@datarumble.com"
$Subject = "Daily Export"
$ProcessedFolderPath = "\Inbox\Processed"
$OneDriveUploadFilePath = "/test" 

##Get the Access Token
$AccessToken =  Get-AccessToken -MailboxName $MailboxName  -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl "urn:ietf:wg:oauth:2.0:oob" -ResourceURL graph.microsoft.com  
##Search the Inbox
$Filter = "IsRead eq false AND HasAttachments eq true AND Subject eq '" + $Subject + "'"
$Items = Get-FolderItems -MailboxName $MailboxName -AccessToken $AccessToken -FolderPath \Inbox -Filter $Filter
if($Items -ne $null){
   if($Items -is [system.array]){
         Write-Host ($Items.Count.ToString() + " Items Found ")
   }
   else{
        Write-Host "Found 1 item"
   }
   foreach ($item in $Items) {
        Write-Host ("Processing Item received " + $Item.receivedDateTime)
        $item
        Get-Attachments -MailboxName $MailboxName -ItemURI $item.ItemRESTURI -MetaData -AccessToken $AccessToken | ForEach-Object{
            $attach = Invoke-DownloadAttachment -MailboxName $MailboxName -AttachmentURI $_.AttachmentRESTURI -AccessToken $AccessToken
            $attachBytes = [System.Convert]::FromBase64String($attach.ContentBytes)   
            $OneDriveFilePath = $OneDriveUploadFilePath + "/" + $attach.Name.ToString()
		    Invoke-UploadOneDriveItemToPath -AccessToken $AccessToken -MailboxName $MailboxName -OneDriveUploadFilePath $OneDriveFilePath -FileBytes $attachBytes 
            write-host ("uploaded " + $OneDriveFilePath + " to one drive")
        }
        $UpdateProps = @()
        $UpdateProps += (Get-ItemProp -Name IsRead -Value true -NoQuotes)
        Update-Message -MailboxName $MailboxName -AccessToken $AccessToken -ItemURI $item.ItemRESTURI -StandardPropList $UpdateProps
        Move-Message -MailboxName $MailboxName -ItemURI $item.ItemRESTURI -TargetFolderPath $ProcessedFolderPath -AccessToken $AccessToken                
   }
  
}
else{
    Write-Host "No Item found"
}
