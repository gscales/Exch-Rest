function Get-EXRWellKnownFolder{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [String]$FolderName
    )
    Begin{
        if($AccessToken -eq $null)
        {
            $AccessToken = Get-EXRProfiledToken -MailboxName $MailboxName  
            if($AccessToken -eq $null){
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }  
            $HttpClient =  Get-EXRHTTPClient -MailboxName $MailboxName
            $EndPoint =  Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders/$FolderName"
            if($PropList -ne $null){
               $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
               Write-Host $RequestURL
            }
            $JSONOutput = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            $folderId = $JSONOutput.Id.ToString()
            Add-Member -InputObject $JSONOutput  -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
            return $JSONOutput    

    }
}
