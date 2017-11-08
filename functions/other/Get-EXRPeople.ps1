function  Get-EXRPeople {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [String]$filter,
        [Parameter(Position=3, Mandatory=$false)] [String]$Search     
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-EXRHTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000"
        if(![String]::IsNullOrEmpty($filter)){
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000&`$filter=" + $filter
        }
        if(![String]::IsNullOrEmpty($Search)){
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000&`$search=" + $Search
        }
        Write-Host $RequestURL
        do{
            $JSONOutput = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))     
    }
}
