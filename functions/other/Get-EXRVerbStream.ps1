function Get-EXRVerbStream {

	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MessageClass,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[psobject]
		$VerbArray
		
	)
	Process
	{       
    
    $vCount = ($VerbArray.length + 4);
    $Header = "02010" + $vCount + "00000000000000";
    $ReplyToAllHeader = "055265706C790849504D2E4E6F7465074D657373616765025245050000000000000000";
    $ReplyToAllFooter = "0000000000000002000000660000000200000001000000";
    $ReplyToHeader = "0C5265706C7920746F20416C6C0849504D2E4E6F7465074D657373616765025245050000000000000000";
    $ReplyToFooter = "0000000000000002000000670000000300000002000000";
    $ForwardHeader = "07466F72776172640849504D2E4E6F7465074D657373616765024657050000000000000000";
    $ForwardFooter = "0000000000000002000000680000000400000003000000";
    $ReplyToFolderHeader = "0F5265706C7920746F20466F6C6465720849504D2E506F737404506F737400050000000000000000";
    $ReplyToFolderFooter = "00000000000000020000006C00000008000000";
    $VoteOptionExtras = "0401055200650070006C00790002520045000C5200650070006C007900200074006F00200041006C006C0002520045000746006F007200770061007200640002460057000F5200650070006C007900200074006F00200046006F006C0064006500720000";
    $DisableReplyAllVal = "00";
    $DisableReplyAllVal = "01";
    $DisableReplyVal = "00";
    $DisableReplyVal = "01";
    $DisableForwardVal = "00";
    $DisableForwardVal = "01";
    $DisableReplyToFolderVal = "00";
    $DisableReplyToFolderVal = "01";
    $OptionsVerbs = "";
    $VerbValue = $Header + $ReplyToAllHeader + $DisableReplyAllVal + $ReplyToAllFooter + $ReplyToHeader + $DisableReplyVal + $ReplyToFooter + $ForwardHeader + $DisableForwardVal + $ForwardFooter + $ReplyToFolderHeader + $DisableReplyToFolderVal + $ReplyToFolderFooter;
    for ($index = 0; $index -lt $VerbArray.length; $index++) {
        $VerbValue += Get-EXRWordVerb -Word $VerbArray[$index] -Postion ($index + 1) -MessageClass $MessageClass
        $VbValA = Invoke-convertToHexUnicode($VerbArray[$index])
        $VbhVal = Invoke-decimalToHexString($VerbArray[$index].length)
        $vbValB = Invoke-convertToHexUnicode($VerbArray[$index])
        $vbPos = Invoke-decimalToHexString($VerbArray[$index].length)
        $OptionsVerbs += $vbPos  + $VbValA  + $VbhVal + $vbValB
    }
    $VerbValue += $VoteOptionExtras + $OptionsVerbs;
    return $VerbValue;
}
}

function Get-EXRWordVerb {
   	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Word,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[decimal]
        $Postion,
        
        [Parameter(Position = 2, Mandatory = $true)]
		[psobject]
        $MessageClass        
		
	)
	Begin
	{
    $verbstart = "04000000";
    $length = Invoke-decimalToHexString($Word.length);
    $HexString =  [System.BitConverter]::ToString([System.Text.UnicodeEncoding]::ASCII.GetBytes($Word)).Replace("-","") 
    $mclength = Invoke-decimalToHexString($MessageClass.length);
    $mcHexString = [System.BitConverter]::ToString([System.Text.UnicodeEncoding]::ASCII.GetBytes($MessageClass)).Replace("-","") 
    $Option1 = "000000000000000000010000000200000002000000";
    $Option2 = "000000FFFFFFFF";
    $lo = Invoke-decimalToHexString -number $Postion
    return ($verbstart + $length + $HexString + $mclength + $mcHexString + "00" + $length + $HexString + $Option1 + $lo + $Option2) ;
    }
}

function Invoke-decimalToHexString {
       	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[Int]
        $number
		
    )
    Begin{
    if ($number -lt 0) {
        $number = 0xFFFFFFFF + $number + 1;
    }
    $numberret = "{0:x}" -f $number
    if ($numberret.length -eq 1) {
        $numberret = "0" + $numberret;
    }
    return $numberret;
    }
}


function Invoke-convertToHexUnicode {
           	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
        $str
		
    )
    Begin{
    $hex =  [System.BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes($str)).Replace("-","")
    return $hex;
    }
}

function Invoke-hex2binarray($hexString){
    $i = 0
    [byte[]]$binarray = @()
    while($i -le $hexString.length - 2){
        $strHexBit = ($hexString.substring($i,2))
        $binarray += [byte]([Convert]::ToInt32($strHexBit,16))
        $i = $i + 2
    }
    return ,$binarray
}