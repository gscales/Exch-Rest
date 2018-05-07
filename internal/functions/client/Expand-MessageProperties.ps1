function Expand-MessageProperties
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
        $Item
	)
	
 	process
	{
        try{
            if ([bool]($Item.PSobject.Properties.name -match "sender"))
            {
                $SenderProp = $Item.sender
                if ([bool]($SenderProp.PSobject.Properties.name -match "emailaddress"))
                {
                    Add-Member -InputObject $Item -NotePropertyName "SenderEmailAddress" -NotePropertyValue $SenderProp.emailaddress.address
                    Add-Member -InputObject $Item -NotePropertyName "SenderName" -NotePropertyValue $SenderProp.emailaddress.name
                }
                
            }
            if ([bool]($Item.PSobject.Properties.name -match "InternetMessageHeaders"))
            {
                $IndexedHeaders = New-Object 'system.collections.generic.dictionary[string,string]'
                foreach($header in $Item.InternetMessageHeaders){
                    if(!$IndexedHeaders.ContainsKey($header.name)){
                        $IndexedHeaders.Add($header.name,$header.value)
                    }
                }
                 Add-Member -InputObject $Item -NotePropertyName "IndexedInternetMessageHeaders" -NotePropertyValue $IndexedHeaders
            }
            
        }
        catch{

        }
    }
}