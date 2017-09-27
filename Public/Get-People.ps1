function  Get-People {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [String]$filter      
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000"
        if(![String]::IsNullOrEmpty($filter)){
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/people?`$Top=1000&`$filter=" + $filter
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