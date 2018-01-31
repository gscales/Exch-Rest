function Get-EXRSearchFolders{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
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
        $SearchFolder = Get-EXRWellKnownFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName SearchFolders
        Get-EXRChildFolders -Folder $SearchFolder -MailboxName $MailboxName -AccessToken  $AccessToken 

    }
}
