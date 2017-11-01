function Get-EXRItemSize(){
        $PR_MESSAGE_SIZE= Get-TaggedProperty -DataType "Integer" -Id "0x0E08"  
        $Props = @()
        $Props +=$PR_MESSAGE_SIZE
	return $Props
}