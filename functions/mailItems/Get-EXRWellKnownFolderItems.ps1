function Get-EXRWellKnownFolderItems{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$WellKnownFolder,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$Folder,
 	    [Parameter(Position=3, Mandatory=$false)] [String]$FolderPath,
        [Parameter(Position=4, Mandatory=$false)] [switch]$ReturnSize,
        [Parameter(Position=5, Mandatory=$false)] [string]$SelectProperties,
        [Parameter(Position=6, Mandatory=$false)] [string]$Filter,
        [Parameter(Position=7, Mandatory=$false)] [string]$Top,
        [Parameter(Position=8, Mandatory=$false)] [string]$OrderBy,
        [Parameter(Position=9, Mandatory=$false)] [switch]$TopOnly,
        [Parameter(Position=10, Mandatory=$false)] [PSCustomObject]$PropList,
        [Parameter(Position=11, Mandatory=$false)] [psobject]$ClientFilter,
        [Parameter(Position=12, Mandatory=$false)] [string]$ClientFilterTop,
        [Parameter(Position=13, Mandatory=$false)] [string]$Search,
        [Parameter(Position=14, Mandatory=$false)] [switch]$ReturnFolderPath,
        [Parameter(Position=14, Mandatory=$false)] [switch]$ReturnStats,
        [Parameter(Position=15, Mandatory=$false)] [switch]$ReturnAttachments,
        [Parameter(Position=16, Mandatory=$false)] [switch]$ReturnSentiment,
        [Parameter(Position=17, Mandatory=$false)] [switch]$ReturnEntryId,
        [Parameter(Position=18, Mandatory=$false)] [switch]$BatchReturnItems,
        [Parameter(Position=19, Mandatory=$false)] [switch]$ReturnInternetMessageHeaders,
        [Parameter(Position=20, Mandatory=$false)] [switch]$ProcessAntiSPAMHeaders,
        [Parameter(Position=22, Mandatory=$false)] [switch]$Todays,
        [Parameter(Position=21, Mandatory=$false)] [Int32]$MessageCount
        
    )
    Begin{
		if($AccessToken -eq $null)
        {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if($AccessToken -eq $null){
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if($Todays.isPresent){
            $Filter = "receivedDateTime ge " + (Get-Date).Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        if([String]::IsNullOrEmpty($MailboxName)){
            $MailboxName = $AccessToken.mailbox
        } 
        if(![String]::IsNullorEmpty($Filter)){
            $Filter = "`&`$filter=" + [System.Web.HttpUtility]::UrlEncode($Filter)
        }
        if(![String]::IsNullorEmpty($Orderby)){
            $OrderBy = "`&`$OrderBy=" + $OrderBy
        }        
        $TopValue = "1000"    
        if(![String]::IsNullorEmpty($Top)){
            if($Top -lt 1000){
                $TopValue =$top
            }           
        }      
        if(![String]::IsNullOrEmpty($ClientFilterTop)){
            $TopOnly = $false
        }
        $restrictProps = $false
        if([String]::IsNullorEmpty($SelectProperties)){            
            $SelectProperties = "`$select=ReceivedDateTime,Sender,Subject,IsRead,inferenceClassification,InternetMessageId,parentFolderId,hasAttachments,webLink"
        }
        else{
            $restrictProps = $true
            $SelectProperties = "`$select=" + $SelectProperties
        }
        if($ReturnInternetMessageHeaders.IsPresent){
            $SelectProperties += ",InternetMessageHeaders"
        }
        if(![String]::IsNullorEmpty($Search)){
            $Search = "`&`$Search=`"" + $Search + "`""
        }
        $BatchReturn = $false
        if($BatchReturnItems.IsPresent){
            $BatchReturn = $true
        }
        $ParentFolderCollection = New-Object 'system.collections.generic.dictionary[[string],[string]]'
        $stats = "" | Select TotalItems
        $stats.TotalItems = 0;
        $FolderOkay = $true
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        if(![String]::IsNullOrEmpty($WellKnownFolder))
        {
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders/" + $WellKnownFolder + "/messages/?" +  $SelectProperties + "`&`$Top=" + $TopValue 
            $folderURI =  $EndPoint + "('" + $MailboxName + "')/MailFolders/" + $WellKnownFolder
            $FolderOkay = $true
        }
        else{
           	if(![String]::IsNullorEmpty($FolderPath))
            {
		        $Folder = Get-ExrFolderFromPath -MailboxName $MailboxName -AccessToken $AccessToken -FolderPath $FolderPath	
            }
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders('" + $Folder.Id + "')/messages/?" +  $SelectProperties + "`&`$Top=" + $TopValue 
            $folderURI =  $EndPoint + "('" + $MailboxName + "')/MailFolders('" + $Folder.Id + "')"   
            $FolderOkay = $true
        }        
        if($FolderOkay){
            if($ReturnSize.IsPresent){
                $PropList = Get-EXRKnownProperty -PropertyName "MessageSize"
            }
            if($ReturnSentiment.IsPresent){
                $BatchReturn = $true
                $PropList = Get-EXRKnownProperty -PropList $PropList -PropertyName "Sentiment"
            }
            if($ReturnEntryId.IsPresent){
                $PropList = Get-EXRKnownProperty -PropList $PropList -PropertyName "PR_ENTRYID"
            }
            $RequestURL += $Filter + $Search + $OrderBy
            if($PropList -ne $null){
               $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            $returnCount = 0
            $clientReturnCount = 0;
            $BatchItems = @()
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
                    $returnCount++
                    if($BatchReturn){
                        if(![String]::IsNullOrEmpty($ClientFilter)){
                            switch($ClientFilter.Operator){
                                "eq" {
                                    if($Message.($ClientFilter.Property) -eq $ClientFilter.Value){
                                        $BatchItems += $Message
                                        $clientReturnCount++
                                    }   
                                }
                                "ne" {
                                    if($Message.($ClientFilter.Property) -ne $ClientFilter.Value){
                                        $BatchItems += $Message
                                        $clientReturnCount++
                                    }
                                }
                            }
                            if(![String]::IsNullOrEmpty($ClientFilterTop)){
                                if($clientReturnCount -ge [Int]::Parse($ClientFilterTop)){
                                    if($BatchItems.Count -gt 0){
                                       Write-Host("Getting Batch of " + $BatchItems.Count + " current Return Count " + $returnCount)
                                        Get-EXRBatchItems -Items $BatchItems -SelectProperties $SelectProperties -URLString ("/users" + "('" + $MailboxName + "')" + "/messages") -PropList $PropList -ReturnAttachments:$ReturnAttachments.isPresent -ProcessAntiSPAMHeaders:$ProcessAntiSPAMHeaders.IsPresent -RestrictProps:$restrictProps
                                        $BatchItems = @()
                                    }
                                    return 
                                }
                            }

                        }else{
                            $BatchItems += $Message
                        }                
                        if($BatchItems.Count -eq 20){
                           Write-Host("Getting Batch of " + $BatchItems.Count + " current Return Count " + $returnCount)
                            Get-EXRBatchItems -Items $BatchItems -SelectProperties $SelectProperties -URLString ("/users" + "('" + $MailboxName + "')" + "/messages") -PropList $PropList -ReturnAttachments:$ReturnAttachments.isPresent -ProcessAntiSPAMHeaders:$ProcessAntiSPAMHeaders.IsPresent -RestrictProps:$restrictProps
                            $BatchItems = @()
                        }
                    }else{
                        $stats.TotalItems++
                        Add-Member -InputObject $Message -NotePropertyName ItemRESTURI -NotePropertyValue ($EndPoint + "('" + $MailboxName + "')/messages('" + $Message.Id + "')")
                        Expand-MessageProperties -Item $Message
                        Expand-ExtendedProperties -Item $Message
                        if($ProcessAntiSPAMHeaders.IsPresent){
                            Invoke-EXRProcessAntiSpamHeaders -Item $Message
                        }
                        if($ReturnFolderPath.IsPresent){
                            if($ParentFolderCollection.ContainsKey($Message.parentFolderId)){
                                add-Member -InputObject $Message -NotePropertyName FolderPath -NotePropertyValue $ParentFolderCollection[$Message.parentFolderId]
                            }
                            else{
                                $Folder = Get-EXRFolderFromId -MailboxName $MailboxName -AccessToken $AccessToken -FolderId $Message.parentFolderId
                                if($Folder -ne $null){
                                    $ParentFolderCollection.Add($Message.parentFolderId,$Folder.PR_Folder_Path)
                                    
                                }else{
                                    $ParentFolderCollection.Add($Message.parentFolderId,"Unavailable")
                                }
                                add-Member -InputObject $Message -NotePropertyName FolderPath -NotePropertyValue $ParentFolderCollection[$Message.parentFolderId]                      

                            }
                        }
                        if($ReturnAttachments.IsPresent -band $Message.hasAttachments){
                            $AttachmentNames = @()
                            $AttachmentDetails = @()
                            Get-EXRAttachments -MailboxName $MailboxName -AccessToken $AccessToken -ItemURI $Message.ItemRESTURI | ForEach-Object{
                                $AttachmentNames += $_.name
                                $AttachmentDetails += $_    
                            }
                            add-Member -InputObject $Message -NotePropertyName AttachmentNames -NotePropertyValue $AttachmentNames
                            add-Member -InputObject $Message -NotePropertyName AttachmentDetails -NotePropertyValue $AttachmentDetails
                        }
                        if(![String]::IsNullOrEmpty($ClientFilter)){
                            switch($ClientFilter.Operator){
                                "eq" {
                                    if($Message.($ClientFilter.Property) -eq $ClientFilter.Value){
                                        Write-Output $Message
                                        $clientReturnCount++
                                    }   
                                }
                                "ne" {
                                    if($Message.($ClientFilter.Property) -ne $ClientFilter.Value){
                                        Write-Output $Message
                                        $clientReturnCount++
                                    }
                                }
                            }
                            if(![String]::IsNullOrEmpty($ClientFilterTop)){
                                if($clientReturnCount -ge [Int]::Parse($ClientFilterTop)){
                                    return 
                                }
                            }

                        }
                        else{
                            Write-Output $Message
                        }
                    }
                    if([String]::IsNullOrEmpty($ClientFilter)){
                        if($MessageCount -gt 0 ){
                            if($returnCount -ge $MessageCount){
                                if($BatchItems.Count -gt 0){
                                    Write-Host("Getting Batch of " + $BatchItems.Count + " current Return Count " + $returnCount)
                                    Get-EXRBatchItems -Items $BatchItems -SelectProperties $SelectProperties -URLString ("/users" + "('" + $MailboxName + "')" + "/messages") -PropList $PropList -ReturnAttachments:$ReturnAttachments.isPresent -ProcessAntiSPAMHeaders:$ProcessAntiSPAMHeaders.IsPresent -RestrictProps:$restrictProps
                                    $BatchItems = @()
                                }
                                return
                            }
                        }
                    }                                        
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL) -band (!$TopOnly))  
            if($BatchItems.Count -gt 0){
                 Write-Host("Getting Batch of " + $BatchItems.Count)
                 Get-EXRBatchItems -Items $BatchItems -SelectProperties $SelectProperties -URLString ("/users" + "('" + $MailboxName + "')" + "/messages") -PropList $PropList -ReturnAttachments:$ReturnAttachments.isPresent -ProcessAntiSPAMHeaders:$ProcessAntiSPAMHeaders.IsPresent -RestrictProps:$restrictProps
                 $BatchItems = @()
            }
            if($ReturnStats.IsPresent){
                Write-Host $stats -ForegroundColor Green
            }   
       } 
   

    }
}
