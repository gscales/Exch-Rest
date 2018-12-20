function Get-EXRMessageTrace {
    [CmdletBinding()]
    param (
		
        [Parameter(Position = 1, Mandatory = $false)]
        [String]
        $MailboxName,		
		
        [Parameter(Position = 6, Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credentials,
        
        [Parameter(Position = 7, Mandatory = $false)]
        [datetime]$Start,

        [Parameter(Position = 8, Mandatory = $false)]
        [datetime]$End,

        [Parameter(Position = 9, Mandatory = $false)]
        [String]$ToAddress,

        [Parameter(Position = 10, Mandatory = $false)]
        [String]$SenderAddress,

        [Parameter(Position = 11, Mandatory = $false)]
        [String]$Status,

        [Parameter(Position = 12, Mandatory = $false)]
        [switch]$TraceDetail,

        [Parameter(Position = 13, Mandatory = $false)]
        [String]$extraQueryOptions,

        [Parameter(Position = 14, Mandatory = $false)]
        [String]$MessageId,

        [Parameter(Position = 15, Mandatory = $false)]
        [String]$ClientFilter
    )
    process {
        if ($Start -eq $null) {$Start = (Get-Date).AddDays(-7)}
        if ($End -eq $null) {$End = (Get-Date)}
        $HttpClient = Get-HTTPClient -MailboxName $Credentials.UserName
        $OdataOptions = "";
        $OdataOptions = "?`$filter=StartDate eq datetime'" + ($Start.ToUniversalTime().ToString("s") + "Z") + "' and EndDate eq datetime'" + ($End.ToUniversalTime().ToString("s") + "Z") + "'";
        if (![String]::IsNullOrEmpty($ToAddress)) {
            $OdataOptions += " and RecipientAddress eq '" + $ToAddress + "'"
        }
        if (![String]::IsNullOrEmpty($SenderAddress)) {
            $OdataOptions += " and SenderAddress eq '" + $SenderAddress + "'"
        }
        if (![String]::IsNullOrEmpty($Status)) {
            $OdataOptions += " and Status eq '" + $Status + "'"
        }
        if(![String]::IsNullOrEmpty($MessageId)){
            $OdataOptions += " and MessageId eq '" + $MessageId + "'"     
        }
        if (![String]::IsNullOrEmpty($extraQueryOptions)) {
            $OdataOptions += " and " + $extraQueryOptions
            Write-Host $OdataOptions
        }
        $ReportingURI = ("https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace" + $OdataOptions);
        do {
            $RequestURI = $ReportingURI.Replace("../../", "https://reports.office365.com/ecp/")
            $ReportingURI = ""
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURI -HttpClient $HttpClient -BasicAuthentication -Credentials $Credentials -MailboxName $Credentials.UserName
            $ReportingURI = $JSONOutput.'odata.nextLink'
            foreach ($Message in $JSONOutput.Value) {
                $ProcessTrace = $true
                if (![String]::IsNullOrEmpty($ClientFilter)) {
                        $ProcessTrace = $false
                        $Val =  $Message | Where-Object ([scriptblock]::Create($ClientFilter))
                        if($Val){
                                $ProcessTrace = $true
                        }                       
                        
                }
                if ($ProcessTrace) {
                    if ($TraceDetail.IsPresent) {

                        $Details = Get-EXRMessageTraceDetail -MailboxName $Credentials.UserName -MessageTraceId $Message.MessageTraceId -SenderAddress $Message.SenderAddress -ToAddress $Message.RecipientAddress -Start $Start -End $End -Credentials $Credentials
                        Add-Member -InputObject $Message -NotePropertyName TraceDetails -NotePropertyValue $Details
                    }
                    Write-Output $Message
                }               
            }
        }while (![String]::IsNullOrEmpty($ReportingURI))

    }
}

