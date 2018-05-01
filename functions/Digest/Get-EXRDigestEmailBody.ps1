
function Get-EXRDigestEmailBody
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$MessageList,
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
        $weblink
	)
	
 	process
	{
		$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:60%;`" ><b>Subject</b></td>" +"`r`n"
		$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>" +"`r`n"
		$rpReport = $rpReport + "</tr>" + "`r`n"
		foreach ($message in $MessageList){
			$fromstring = $message.SenderEmailAddress
			$Oulookid = $message.PR_EntryId
			if ($fromstring.length -gt 30){$fromstring = $fromstring.Substring(0,30)}
			$rpReport = $rpReport + "  <tr>"  + "`r`n"
			$rpReport = $rpReport + "<td>" + [DateTime]::Parse($message.receivedDateTime).ToString("G") + "</td>"  + "`r`n"
			$rpReport = $rpReport + "<td>" +  $fromstring + "</td>"  + "`r`n"
			if($weblink.IsPresent){
				$rpReport = $rpReport + "<td><a href=`"" + $message.weblink + "`">" + $message.Subject + "</td>"  + "`r`n"
			}
			else{
				$rpReport = $rpReport + "<td><a href=`"outlook:" + $Oulookid + "`">" + $message.Subject + "</td>"  + "`r`n"
			}			
			$rpReport = $rpReport + "<td>" +  ($message.Size/1024).ToString(0.00) + "</td>"  + "`r`n"
			$rpReport = $rpReport + "</tr>"  + "`r`n"
		}
		$rpReport = $rpReport + "</table>"  + "  " 
		return $rpReport
    }
}