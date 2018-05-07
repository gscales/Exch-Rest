function Invoke-EXRProcessSentiment {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $Item,
        [Parameter(Position = 2, Mandatory = $false)]
        [psobject]
        $JSONData
    )
	
    process {
        try {
            $emotiveProfile = ConvertFrom-Json -InputObject $JSONData
            Add-Member -InputObject $Item -NotePropertyName "Sentiment" -NotePropertyValue $emotiveProfile.sentiment.polarity
            Add-Member -InputObject $Item -NotePropertyName "EmotiveProfile" -NotePropertyValue $emotiveProfile
        }
        catch {}
		
    }
}