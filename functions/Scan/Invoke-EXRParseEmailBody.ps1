function Invoke-EXRParseEmailBody {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]$BodyText

    )
	
    process {
        $ParsedLinksObject = "" | Select HasBaseURL, ParsedBaseURL, Links, Images
        $ParsedLinksObject.HasBaseURL = $false
        $ParsedLinksObject.Links = @()      
        $ParsedLinksObject.Images = @()  
        
        $RegExHtmlLinks = "<`(.*?)>"  
        $matchedItems = [regex]::matches($BodyText, $RegExHtmlLinks,[system.Text.RegularExpressions.RegexOptions]::Singleline)  
        foreach($Match in $matchedItems){   
            if(!$Match.Value.StartsWith("</")){
                if($Match.Value.StartsWith("<a "),[System.StringComparison]::InvariantCultureIgnoreCase){
                    $Attributes = $Match.Value.Split(" ")
                    foreach($Attribute in $Attributes){
                        if($Attribute.StartsWith('href=',[System.StringComparison]::InvariantCultureIgnoreCase)){                                        
                            $ParsedLinksObject.Links += ([URI]($Attribute.Substring(6,$Attribute.Length-7)))   
                        }                                   
                    }                                
                }
                if($Match.Value.StartsWith("<base ",[System.StringComparison]::InvariantCultureIgnoreCase)){
                    $ParsedLinksObject.HasBaseURL = $true
                    $Attributes = $Match.Value.Split(" ")
                    foreach($Attribute in $Attributes){
                        if($Attribute.StartsWith('href=',[System.StringComparison]::InvariantCultureIgnoreCase)){                                        
                            $ParsedLinksObject.ParsedBaseURL = ([URI]($Attribute.Substring(6,$Attribute.Length-7)))   
                        }                                   
                    }                                  
                }
                if($Match.Value.StartsWith("<img ",[System.StringComparison]::InvariantCultureIgnoreCase)){
                    $Attributes = $Match.Value.Split(" ")
                    foreach($Attribute in $Attributes){
                        if($Attribute.StartsWith('src=',[System.StringComparison]::InvariantCultureIgnoreCase)){                                        
                            $ParsedLinksObject.Images += ([URI]($Attribute.Substring(5,$Attribute.Length-6)))   
                        }                                   
                    
                    }
                }                        
            }  
        }        



        
       return,$ParsedLinksObject
    }

    
}

function Get-EXRHTMLAttribute{
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [String]$Tag,
        [Parameter(Position = 2, Mandatory = $false)]
        [String]$Attribute
    )
	
    process {
        $Tag.Split(" ")
    }
}