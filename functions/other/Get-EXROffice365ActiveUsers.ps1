function  Get-EXROffice365ActiveUsers {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [String]$ViewType,
        [Parameter(Position=3, Mandatory=$true)] [String]$PeriodType   
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
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "reports"
        $RequestURL =  $EndPoint + "/Office365ActiveUsers(view='$ViewType',period='$PeriodType')/content"
        Write-Host $RequestURL
        $Output = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -NoJSON
        $OutPutStream = $Output.ReadAsStreamAsync().Result
        return ConvertFrom-Csv ([System.Text.Encoding]::UTF8.GetString($OutPutStream.ToArray()))
    }
}
