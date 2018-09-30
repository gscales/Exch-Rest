function Get-EXRSchedule {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,

        [Parameter(Position = 2, Mandatory = $false)]
        [psobject]
        $Mailboxes,

        [Parameter(Position = 3, Mandatory = $false)]
        [String]
		$TimeZone,

		[Parameter(Position = 4, Mandatory = $false)]
        [Datetime]
		$StartTime,

		[Parameter(Position = 5, Mandatory = $false)]
        [Datetime]
		$EndTime,

		[Parameter(Position = 6, Mandatory = $false)]
        [Int]		
		$availabilityViewInterval=15


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
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $SchPost = @{};
        $array = @()
        $StartTimeHash = @{}
        $EndTimeHash = @{}
        $SchPost.Add("schedules", $array)
        $SchPost.Add("startTime", $StartTimeHash)
        $SchPost.Add("endTime", $EndTimeHash)
        $SchPost.Add("availabilityViewInterval", $availabilityViewInterval)
        foreach ($Mailbox in $Mailboxes) {
            $SchPost.schedules += $Mailbox
        }
        if ([String]::IsNullOrEmpty($TimeZone)) {
            $TimeZone = [TimeZoneInfo]::Local.Id
		}		
		if($StartTime -eq $null){
			$StartTime = (Get-Date)
		}
		if($EndTime -eq $null){
			$EndTime = (Get-Date).AddHours(24)
		}
        $SchPost.startTime.Add("dateTime", $StartTime.ToString("yyyy-MM-ddTHH:mm:ss"))
        $SchPost.startTime.Add("timeZone", $TimeZone)
        $SchPost.endTime.Add("dateTime", $EndTime.ToString("yyyy-MM-ddTHH:mm:ss"))
        $SchPost.endTime.Add("timeZone", $TimeZone)

        $RequestURL = "https://graph.microsoft.com/beta/me/calendar/getschedule"
        $JSONOutput = Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content (ConvertTo-Json -InputObject $SchPost -Depth 10) -TimeZone $TimeZone
        foreach ($Value in $JSONOutput.value) {
            write-output $Value
        } 
		
    }
}
