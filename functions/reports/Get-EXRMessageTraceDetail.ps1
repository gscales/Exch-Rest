function Get-EXRMessageTraceDetail
{
	[CmdletBinding()]
	param (
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$MailboxName,		
		
		[Parameter(Position = 2, Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credentials,
        
        [Parameter(Position = 3, Mandatory = $true)]
        [datetime]$Start,

        [Parameter(Position = 4, Mandatory = $true)]
        [datetime]$End,

        [Parameter(Position = 5, Mandatory = $true)]
        [String]$ToAddress,

        [Parameter(Position = 6, Mandatory = $true)]
        [String]$SenderAddress,

        [Parameter(Position = 7, Mandatory = $true)]
        [String]$MessageTraceId
	)
	process
	{
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $OdataOptions = "";
        $OdataOptions = "?`$filter=StartDate eq datetime'" + ($Start.ToString("s") + "Z") + "' and EndDate eq datetime'" + ($End.ToString("s") + "Z") + "'";
        if(![String]::IsNullOrEmpty($ToAddress)){
                $OdataOptions += " and RecipientAddress eq '" + $ToAddress + "'"
        }
        if(![String]::IsNullOrEmpty($SenderAddress)){
                $OdataOptions += " and SenderAddress eq '" + $SenderAddress + "'"
        }
        if(![String]::IsNullOrEmpty($MessageTraceId)){
                $OdataOptions += " and MessageTraceId eq guid'" + $MessageTraceId + "'"
        }    
        $ReportingURI = ("https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTraceDetail" + $OdataOptions);
        do{
            $RequestURI = $ReportingURI.Replace("../../","https://reports.office365.com/ecp/")
            $ReportingURI = ""
            $JSONOutput =  Invoke-RestGet -RequestURL $RequestURI -HttpClient $HttpClient -BasicAuthentication -Credentials $Credentials -MailboxName $MailboxName
            $ReportingURI = $JSONOutput.'odata.nextLink'
            foreach($Message in $JSONOutput.Value){
                Write-Output $Message
            }
        }while(![String]::IsNullOrEmpty($ReportingURI))

    }
}

