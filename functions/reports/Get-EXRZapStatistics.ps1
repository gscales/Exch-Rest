function Get-EXRZapStatistics{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [switch]$RawData,
        [Parameter(Position = 3, Mandatory = $false)] [DateTime]$startdatetime = (Get-Date).AddDays(-365),
        [Parameter(Position = 4, Mandatory = $false)] [datetime]$enddatetime = (Get-Date)
    )
    Begin{
        $rptCollection = @()
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
        $Filter = "receivedDateTime ge " + $startdatetime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + " and receivedDateTime le " + $enddatetime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $Filter += " and singleValueExtendedProperties/any(ep: ep/id eq 'String {00062008-0000-0000-c000-000000000046} Name X-Microsoft-Antispam-ZAP-Message-Info' and ep/value ne null)"         
        $Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder junkemail -Filter $Filter -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders -AdditionalProperties "lastModifiedDateTime"
        if($RawData.IsPresent){
            return  $Items
        }else{
            foreach($Item in $Items){
                $rptObject = "" | Select DateTimeReceived,LastModified,MinutesInInbox,HoursInInbox,Subject,Sender,Read,InternetMessageId,SPF,DKIM,DMARC,CompAuth,PCL,BCP,CTRY,SFV,SRV,PTR,CIP,IPV,SCL     
                $rptObject.DateTimeReceived = [DateTime]$Item.receivedDateTime 
                $rptObject.LastModified = [DateTime]$Item.lastModifiedDateTime
                $TimeSpan = New-TimeSpan -Start $rptObject.DateTimeReceived -End $rptObject.LastModified
                $rptObject.MinutesInInbox = [Math]::Round($TimeSpan.TotalMinutes,0)
                $rptObject.HoursInInbox = [Math]::Round($TimeSpan.TotalHours,0)
                $rptObject.Subject = $Item.Subject
                $rptObject.Read = $Item.isRead
                $rptObject.Sender = $Item.SenderEmailAddress
                $rptObject.InternetMessageId = $Item.InternetMessageId
                $rptObject.SPF = $Item.SPF
                $rptObject.DKIM = $Item.DKIM
                $rptObject.DMARC = $Item.DMARC
                $rptObject.CompAuth = $Item.CompAuth
                $rptObject.PCL = $Item.PCL
                $rptObject.BCP = $Item.BCP
                $rptObject.CTRY = $Item.CTRY
                $rptObject.SFV = $Item.SFV
                $rptObject.SRV = $Item.SRV
                $rptObject.PTR = $Item.PTR
                $rptObject.CIP = $Item.CIP
                $rptObject.IPV = $Item.IPV
                $rptObject.SCL = $Item.SCL
                $rptCollection += $rptObject
            }
            return $rptCollection
        }
        
    }
}