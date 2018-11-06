function Get-EXRInsights{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [switch]$Used,
        [Parameter(Position=3, Mandatory=$false)] [switch]$Shared,
        [Parameter(Position=4, Mandatory=$false)] [switch]$Trending,
        [Parameter(Position=5, Mandatory=$false)] [Int32]$Top=100
      
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
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('" + $MailboxName + "')/insights/used?`$Top=$Top"
        if($Shared.IsPresent){
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/insights/shared?`$Top=$Top"
        }
        if($Trending.IsPresent){
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/insights/trending?`$Top=$Top"
        }
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		return $JSONOutput.value 
    }
}