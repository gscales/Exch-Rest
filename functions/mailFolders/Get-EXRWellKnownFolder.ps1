function Get-EXRWellKnownFolder{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [String]$FolderName,
        [Parameter(Position=3, Mandatory=$false)] [PSCustomObject]$PropList
    )
    Process{

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
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders/$FolderName"
        $PropList = Get-EXRKnownProperty -PropertyName "FolderSize" -PropList $PropList
        if($PropList -ne $null){
            $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
            $RequestURL += "`?`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        }
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        $folderId = $JSONOutput.Id.ToString()
        Add-Member -InputObject $JSONOutput  -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
        Expand-MessageProperties -Item $JSONOutput
        Expand-ExtendedProperties -Item $JSONOutput
        return $JSONOutput    

    }
}
