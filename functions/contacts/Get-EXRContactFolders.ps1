function Get-EXRContactFolders{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
        
    )
    Begin{
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
        $PropList = @()
        $ChildFolderCount = Get-EXRTaggedProperty -Id 0x6638 -DataType Integer
        $PropList += $ChildFolderCount
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        if($RootFolder.IsPresent){
            $RequestURL =   $EndPoint + "('$MailboxName')/contactfolders('contacts')?"
            if($PropList -ne $null){
           		 $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
           		 $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
       		}
            return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        }
        else{
            $RequestURL =   $EndPoint + "('$MailboxName')/contactfolders?"
            if($PropList -ne $null){
           		 $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
           		 $RequestURL += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
       		}
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
                    $Message | Add-Member -NotePropertyName FolderPath -NotePropertyValue ("\" + $Message.DisplayName)
                    Expand-ExtendedProperties -Item $Message
                    Write-Output $Message
                    if($Message.PR_FOLDER_CHILD_COUNT -gt 0){
                        Get-EXRChildContactFolders -MailboxName $MailboxName -AccessToken $AccessToken -Folder $Message -PropList $PropList
                    }
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
        }        

    }
}
