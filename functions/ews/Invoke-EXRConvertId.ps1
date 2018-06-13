function Invoke-EXRConvertId {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,

        [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $ItemId

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
        $PostContent = "<?xml version=`"1.0`" encoding=`"utf-8`"?>"
        $PostContent += "<soap:Envelope xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:m=`"http://schemas.microsoft.com/exchange/services/2006/messages`" xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`">"
        $PostContent += "<soap:Header>"
        $PostContent += "<t:RequestServerVersion Version=`"Exchange2013`" />"
        $PostContent += "</soap:Header>"
        $PostContent += "<soap:Body>"
        $PostContent += "   <m:ConvertId DestinationFormat=`"EwsId`">"
        $PostContent += "     <m:SourceIds>"
        $PostContent += "<t:AlternateId Format=`"OwaId`" Id=`"" + $ItemId + "`" Mailbox=`"" +  $MailboxName + "`" />"
        $PostContent += "</m:SourceIds>"
        $PostContent += "</m:ConvertId>"
        $PostContent += "</soap:Body>"
        $PostContent += "</soap:Envelope>"        
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $RequestURL = "https://outlook.office365.com/EWS/Exchange.asmx" 
        $JSONOutput = Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $PostContent 
        return $JSONOutput 
		
    }
}
