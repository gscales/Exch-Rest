function Get-EXRUsers{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
	[Parameter(Position=2, Mandatory=$false)] [psobject]$filter
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-EXRHTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint
        if(![String]::IsNullOrEmpty($filter)){
                $RequestURL =  $EndPoint + "?`$Top=999&`$filter=" + $filter
        }        
        else{
             $RequestURL =  $EndPoint + "?`$Top=999"
        }
             write-host $RequestURL           
        do{
            $JSONOutput = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))       
        
        
    }
}
