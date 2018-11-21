function Get-EXRGroupFiles {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,
				
        [Parameter(Position = 2, Mandatory = $false)]
        [psobject]
        $GroupId,

        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $FileName,

        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $filter

    )
    Process {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        }  
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName       
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups" 
        $RequestURL = $EndPoint + "/" + $GroupId + "/drive/root/children"
        if(![String]::IsNullOrEmpty($FileName)){
            $RequestURL += "?`$filter= Name eq '" + $FileName + "'"            
        }
        if(![String]::IsNullOrEmpty($filter)){
            $RequestURL += "?`$filter=" + $filter            
        }
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        foreach($Message in $JSONOutput.value){
            Write-Output $Message
        }
        
	
    }
}
