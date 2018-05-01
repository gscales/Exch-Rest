function  Get-EXRPeople {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [String]$filter,
        [Parameter(Position=3, Mandatory=$false)] [String]$Search,
        [Parameter(Position=4, Mandatory=$false)] [String]$EmailAddress     
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
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000"
        if(![String]::IsNullOrEmpty($emailAddress)){
             $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=2&`$filter=scoredemailAddresses/any(a:a/address eq '" + $EmailAddress + "')"
        }
        if(![String]::IsNullOrEmpty($filter)){
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000&`$filter=" + $filter
        }
        if(![String]::IsNullOrEmpty($Search)){
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000&`$search=" + $Search
        }
        Write-Host $RequestURL
        do{
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))     
    }
}
