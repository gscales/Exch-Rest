function Get-EXRTransportHeader(){
        $PR_TRANSPORT_MESSAGE_HEADERS = Get-TaggedProperty -DataType "String" -Id "0x007D"  
        $Props = @()
        $Props +=$PR_TRANSPORT_MESSAGE_HEADERS
	return $Props
}