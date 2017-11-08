function Get-EXRPinnedEmailProperty(){
        $PR_RenewTime = Get-EXRTaggedProperty -DataType "SystemTime" -Id "0x0F02"   
 	$PR_RenewTime2 = Get-EXRTaggedProperty -DataType "SystemTime" -Id "0x0F01"   
        $Props = @()
        $Props += $PR_RenewTime
        $Props += $PR_RenewTime2         
	return $Props
}