function Get-EXRMessageTrace
{
	[CmdletBinding()]
	param (
		
		[Parameter(Position = 1, Mandatory = $true)]
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
        [switch]$TraceDetail
	)
	process
	{
        if($Start -eq $null){$Start = (Get-Date).AddDays(-7)}
        if($End -eq $null){$End = (Get-Date)}
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $OdataOptions = "";
        $OdataOptions = "?`$filter=StartDate eq datetime'" + ($Start.ToString("s") + "Z") + "' and EndDate eq datetime'" + ($End.ToString("s") + "Z") + "'";
        if(![String]::IsNullOrEmpty($ToAddress)){
                $OdataOptions += " and RecipientAddress eq '" + $ToAddress + "'"
        }
        if(![String]::IsNullOrEmpty($SenderAddress)){
                $OdataOptions += " and SenderAddress eq '" + $SenderAddress + "'"
        }
        if(![String]::IsNullOrEmpty($Status)){
                $OdataOptions += " and Status eq '" + $Status + "'"
        }
        $ReportingURI = ("https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace" + $OdataOptions);
        do{
            $RequestURI = $ReportingURI.Replace("../../","https://reports.office365.com/ecp/")
            $ReportingURI = ""
            $JSONOutput =  Invoke-RestGet -RequestURL $RequestURI -HttpClient $HttpClient -BasicAuthentication -Credentials $Credentials -MailboxName $MailboxName
            $ReportingURI = $JSONOutput.'odata.nextLink'
            foreach($Message in $JSONOutput.Value){
                if($TraceDetail.IsPresent){
                   $Details =  Get-EXRMessageTraceDetail -MailboxName $MailboxName -MessageTraceId $Message.MessageTraceId -SenderAddress $Message.SenderAddress -ToAddress $Message.RecipientAddress -Start $Start -End $End -Credentials $Credentials
                   Add-Member -InputObject $Message -NotePropertyName TraceDetails -NotePropertyValue $Details
                }
                Write-Output $Message
            }
        }while(![String]::IsNullOrEmpty($ReportingURI))

    }
}

