function Get-EXRContact {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 3, Mandatory = $true)] [string]$id,
        [Parameter(Position = 4, Mandatory = $false)][psobject]	$PropList
    )  
    Begin {
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
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users" 
        $RequestURL = $EndPoint + "('" + $MailboxName + "')/Contacts('" + $id + "')"
     	if($PropList -ne $null){
               $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
               $RequestURL += "?`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        }
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Expand-ExtendedProperties -Item $JSONOutput
		Expand-MessageProperties -Item $JSONOutput
	    return $JSONOutput

    } 
}

