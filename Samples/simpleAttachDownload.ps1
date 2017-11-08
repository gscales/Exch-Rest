Import-Module Exch-Rest -Force
$MailboxName = "user@domain.com"
$Subject = "Daily Export"
$ProcessedFolderPath = "\Inbox\Processed"
$downloadDirectory = "c:\temp"
##Get the Access Token
$AccessToken =  Get-EXRAccessToken -MailboxName $MailboxName  -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl "urn:ietf:wg:oauth:2.0:oob" -ResourceURL graph.microsoft.com  
##Search the Inbox
$Filter = "IsRead eq false AND HasAttachments eq true AND Subject eq '" + $Subject + "'"
$Items = Get-EXRFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -FolderPath \Inbox -Filter $Filter
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
        Get-EXRAttachments -MailboxName $MailboxName -ItemURI $item.ItemRESTURI -MetaData -AccessToken $AccessToken | ForEach-Object{
            $attach = Invoke-EXRDownloadAttachment -MailboxName $MailboxName -AttachmentURI $_.AttachmentRESTURI -AccessToken $AccessToken
           	$fiFile = new-object System.IO.FileStream(($downloadDirectory  + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
            $attachBytes = [System.Convert]::FromBase64String($attach.ContentBytes)   
		    $fiFile.Write($attachBytes, 0, $attachBytes.Length)
		    $fiFile.Close()
		    write-host ("Downloaded Attachment : " + (($downloadDirectory + "\" + $attach.Name.ToString())))
        }
        $UpdateProps = @()
        $UpdateProps += (Get-EXRItemProp -Name IsRead -Value true -NoQuotes)
        Update-EXRMessage -MailboxName $MailboxName -AccessToken $AccessToken -ItemURI $item.ItemRESTURI -StandardPropList $UpdateProps
        Move-EXRMessage -MailboxName $MailboxName -ItemURI $item.ItemRESTURI -TargetFolderPath $ProcessedFolderPath -AccessToken $AccessToken                
   }
  
}
else{
    Write-Host "No Item found"
}