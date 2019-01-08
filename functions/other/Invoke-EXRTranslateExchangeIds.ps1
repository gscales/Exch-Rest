function Invoke-EXRTranslateExchangeIds {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,
        
        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $SourceId, 
        
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $SourceHexId,
        
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $SourceEMSId,

        [Parameter(Position = 5, Mandatory = $false)]
        [String]
        $SourceFormat,  

        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $TargetFormat,  

        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $returnRawFormat  
    )
    Begin {
		
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
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users" -beta
        $RequestURL = $EndPoint + "/" + $MailboxName + "/translateExchangeIds"
        $ConvertRequest = @{}
        $ConvertRequest.Add("inputIds",@())
        if($SourceHexId){
            $byteArray = @($SourceHexId -split '([a-f0-9]{2})' | foreach-object { if ($_) {[System.Convert]::ToByte($_,16)}})
            $urlSafeString = [Convert]::ToBase64String($byteArray).replace("/","_").replace("+","-")
            if($urlSafeString.contains("==")){$urlSafeString = $urlSafeString.replace("==","2")}
            $ConvertRequest["inputIds"] += $urlSafeString

        }else{
            if($SourceEMSId){
                $HexEntryId = [System.BitConverter]::ToString([Convert]::FromBase64String($SourceEMSId)).Replace("-","").Substring(2)  
                $HexEntryId =  $HexEntryId.SubString(0,($HexEntryId.Length-2))
                $byteArray = @($HexEntryId -split '([a-f0-9]{2})' | foreach-object { if ($_) {[System.Convert]::ToByte($_,16)}})
                $urlSafeString = [Convert]::ToBase64String($byteArray).replace("/","_").replace("+","-")
                if($urlSafeString.contains("==")){$urlSafeString = $urlSafeString.replace("==","2")}
                $ConvertRequest["inputIds"] += $urlSafeString
            }else{
                $ConvertRequest["inputIds"] += $SourceId
            }

        }        
        $ConvertRequest.targetIdType = $TargetFormat
        $ConvertRequest.sourceIdType = $SourceFormat
        $JsonResult = Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content (ConvertTo-Json $ConvertRequest -Depth 9)
        if($TargetFormat.ToLower() -eq "entryid" -band (!$returnRawFormat.IsPresent)){
            $urldecodedstring = $JsonResult.value.targetId.replace("_", "/").replace("-", "+")
            $lastVal = $urldecodedstring.SubString($urldecodedstring.Length-1,1);
            if($lastVal -eq "2"){
                $urldecodedstring = $urldecodedstring.SubString(0,$urldecodedstring.Length-1) + "=="
            }
            return ([System.BitConverter]::ToString([Convert]::FromBase64String($urldecodedstring))).replace("-","")
        }else{
            return  $JsonResult.value.targetId
        }
       	
		
    }
}
