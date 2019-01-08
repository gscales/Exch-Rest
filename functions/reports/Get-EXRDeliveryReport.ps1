function Get-EXRDeliveryReport
{
	[CmdletBinding()]
	param (
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$MessageId,
        [Parameter(Position=3, Mandatory=$false)] [DateTime]$StartTime,
        [Parameter(Position=4, Mandatory=$false)] [DateTime]$EndTime,
        [Parameter(Position=5, Mandatory=$false)] [switch]$fullDetails
	)
	Process
	{
		if($AccessToken -eq $null)
		{
			$AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
			if($AccessToken -eq $null){
				$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
			}                 
		}
		if([String]::IsNullOrEmpty($MailboxName)){
			$MailboxName = $AccessToken.mailbox
		} 
		if([String]::IsNullOrEmpty($WellKnownFolder)){
			$WellKnownFolder = "AllItems"
			if($SearchDumpster.IsPresent){
				$WellKnownFolder = "RecoverableItemsDeletions"
			}
        }        
        $PropList = Get-EXRKnownProperty -PropertyName "PR_LAST_VERB_EXECUTED"
        $PropList = Get-EXRKnownProperty -PropertyName "PR_LAST_VERB_EXECUTION_TIME" -PropList $PropList
        if([String]::IsNullOrEmpty($MessageId)){
            $Filter = "receivedDateTime ge " + $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            $Filter += " And receivedDateTime le " + $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }else{
            $Filter = "internetMessageId eq '" + $MessageId + "'"
        }		
		if($InReplyTo.IsPresent){
			$Filter = "SingleValueExtendedProperties/Any(ep: ep/Id eq 'String 0x1042' and ep/Value eq '" + $MessageId + "')"
        }
        if([String]::IsNullOrEmpty($MessageId)){
            $MsgIndex = @{};
            Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder  -ReturnSize:$ReturnSize.IsPresent -ReturnAttachments -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList -ReturnFolderPath -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders | ForEach-Object{
                if(!$MsgIndex.ContainsKey($_.internetMessageId)){
                    $MsgIndex.add($_.internetMessageId,$_)
                }
            }            
            Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder RecoverableItemsDeletions  -ReturnSize:$ReturnSize.IsPresent -ReturnAttachments -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList -ReturnFolderPath -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders | ForEach-Object{
                if(!$MsgIndex.ContainsKey($_.internetMessageId)){
                    $MsgIndex.add($_.internetMessageId,$_)
                }
            }
             Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder RecoverableItemsPurges  -ReturnSize:$ReturnSize.IsPresent -ReturnAttachments -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList -ReturnFolderPath -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders | ForEach-Object{
                if(!$MsgIndex.ContainsKey($_.internetMessageId)){
                    $MsgIndex.add($_.internetMessageId,$_)
                }
            }
            foreach($item in $MessageTrace){
                if($MsgIndex.ContainsKey()){

                }
            }
            $rptCollection = @()
            Get-MessageTrace -StartDate $StartTime -EndDate $EndTime -RecipientAddress $MailboxName | ForEach-Object{
                if($MsgIndex.ContainsKey($_.MessageId)){
                      $rptObject = "" | Select Received,MessageId,SenderAddress,RecipientAddress,Subject,Size,Status,FromIP,ToIP,MailboxLocation,isRead,hasAttachments,inferenceClassification,LastAction,LastActionTime,SPF,DKIM,DMARC,CTRY,AttachmentNames
                      $rptObject.Received = $_.Received
                      $rptObject.MessageId = $_.MessageId
                      $rptObject.SenderAddress = $_.SenderAddress
                      $rptObject.RecipientAddress = $_.RecipientAddress
                      $rptObject.Subject = $_.Subject
                      $rptObject.Size = $_.Size
                      $rptObject.Status = $_.Status
                      $rptObject.FromIP = $_.FromIP
                      $rptObject.ToIP = $_.ToIP
                      $rptObject.MailboxLocation = $MsgIndex[$_.MessageId].FolderPath
                      $rptObject.isRead = $MsgIndex[$_.MessageId].isRead
                      $rptObject.hasAttachments = $MsgIndex[$_.MessageId].hasAttachments
                      $rptObject.inferenceClassification = $MsgIndex[$_.MessageId].inferenceClassification
                      $rptObject.LastAction = $MsgIndex[$_.MessageId].PR_LAST_VERB_EXECUTED_Displayname
                      $rptObject.LastActionTime = $MsgIndex[$_.MessageId].PR_LAST_VERB_EXECUTION_TIME
                      $rptObject.SPF = $MsgIndex[$_.MessageId].SPF
                      $rptObject.DKIM = $MsgIndex[$_.MessageId].DKIM
                      $rptObject.DMARC = $MsgIndex[$_.MessageId].DMARC
                      $rptObject.CTRY = $MsgIndex[$_.MessageId].CTRY
                      if($MsgIndex[$_.MessageId].AttachmentNames){
                        $rptObject.AttachmentNames = $MsgIndex[$_.MessageId].AttachmentNames
                      }
                      if($fullDetails.IsPresent){
                          Add-Member -InputObject $rptObject -NotePropertyName  internetMessageHeaders -NotePropertyValue $MsgIndex[$_.MessageId].internetMessageHeaders  -Force
                          Add-Member -InputObject $rptObject -NotePropertyName  SenderName -NotePropertyValue $MsgIndex[$_.MessageId].SenderName  -Force
                          Add-Member -InputObject $rptObject -NotePropertyName  SCL -NotePropertyValue $MsgIndex[$_.MessageId].SCL  -Force
                          Add-Member -InputObject $rptObject -NotePropertyName  PTR -NotePropertyValue $MsgIndex[$_.MessageId].PTR  -Force
                          Add-Member -InputObject $rptObject -NotePropertyName  id -NotePropertyValue $MsgIndex[$_.MessageId].id  -Force
                          if($MsgIndex[$_.MessageId].AttachmentDetails){
                            Add-Member -InputObject $rptObject -NotePropertyName  AttachmentDetails -NotePropertyValue $MsgIndex[$_.MessageId].AttachmentDetails  -Force
                          }
                          
                      }
                      $rptCollection += $rptObject

                }

            }
            return $rptCollection
        }
        else{
            $Item = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder  -ReturnSize:$ReturnSize.IsPresent -ReturnAttachments -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList -ReturnFolderPath -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders
            if(!$Item){
                $Item = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder RecoverableItemsDeletions  -ReturnSize:$ReturnSize.IsPresent -ReturnAttachments -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList -ReturnFolderPath -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders
                if(!$Item){
                    $Item = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder RecoverableItemsPurges  -ReturnSize:$ReturnSize.IsPresent -ReturnAttachments -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList -ReturnFolderPath -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders
    
                }
            }
            return $Item
        }

		
	}
}

