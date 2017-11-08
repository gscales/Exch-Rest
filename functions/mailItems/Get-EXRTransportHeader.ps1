function Get-EXRTransportHeader(){
        $PR_TRANSPORT_MESSAGE_HEADERS = Get-EXRTaggedProperty -DataType "String" -Id "0x007D"  
        $Props = @()
        $Props +=$PR_TRANSPORT_MESSAGE_HEADERS
	return $Props
}