function Get-EXRInferenceClassificationResult(){
        $InferenceClassificationResult = Get-EXRNamedProperty -DataType "Integer" -Id "InferenceClassificationResult" -Type String -Guid '23239608-685D-4732-9C55-4C95CB4E8E33'
        $Props = @()
	$Props += $InferenceClassificationResult
	return $Props
}