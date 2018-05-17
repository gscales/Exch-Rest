
function Invoke-EXRParseEmailBodyLinks {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]$Item,
        [Parameter(Position = 2, Mandatory = $false)]
        [switch]$UseExtendedProperty

    )
	
    process {
        $ParsedLinksObject = "" | Select HasBaseURL, ParsedBaseURL, Links, Images
        $ParsedLinksObject.HasBaseURL = $false
        $ParsedLinksObject.Links = @()      
        $ParsedLinksObject.Images = @()  
        $RegExHtmlLinks = "<`(.*?)>"  
        if($UseExtendedProperty.IsPresent){
            $matchedItems = [regex]::matches($Item.PR_BODY_HTML, $RegExHtmlLinks,[system.Text.RegularExpressions.RegexOptions]::Singleline)
        }
        else{
            $matchedItems = [regex]::matches($Item.Body.Content, $RegExHtmlLinks,[system.Text.RegularExpressions.RegexOptions]::Singleline)
        }          
        foreach($Match in $matchedItems){   
            if(!$Match.Value.StartsWith("</")){
                try{
                    if($Match.Value.StartsWith("<base ",[System.StringComparison]::InvariantCultureIgnoreCase)){
                        $ParsedLinksObject.HasBaseURL = $true
                        $Attributes = $Match.Value.Split(" ")
                        foreach($Attribute in $Attributes){
                            if($Attribute.Length -gt 10){
                                if($Attribute.StartsWith('href=',[System.StringComparison]::InvariantCultureIgnoreCase)){                                        
                                    $ParsedLinksObject.ParsedBaseURL = ([URI]($Attribute.Substring(6,$Attribute.Length-7).Replace("`"","").Replace("'","").Replace("`r`n","")))   
                                }   
                            }                                
                        }                                  
                    }
                    if($Match.Value.StartsWith("<a "),[System.StringComparison]::InvariantCultureIgnoreCase){
                        $Attributes = $Match.Value.Split(" ")
                        foreach($Attribute in $Attributes){
                            if($Attribute.StartsWith('href=',[System.StringComparison]::InvariantCultureIgnoreCase)){     
                                if($Attribute.Length -gt 10){
                                    $hrefVal = ([URI]($Attribute.Substring(6,$Attribute.Length-7).Replace("`"","").Replace("'","").Replace("`r`n","")))
                                    if($ParsedLinksObject.HasBaseURL){
                                        if([String]::IsNullOrEmpty($hrefVal.DnsSafeHost)){
                                            $newHost = $ParsedLinksObject.ParsedBaseURL.OriginalString +  $hrefVal.OriginalString
                                            $hrefVal = ([URI]($newHost))
                                        }
                                    }
                                    $ParsedLinksObject.Links += $hrefVal   
                                }                                   
                                
                            }                                   
                        }                                
                    }
                    if($Match.Value.StartsWith("<img ",[System.StringComparison]::InvariantCultureIgnoreCase)){
                        $Attributes = $Match.Value.Split(" ")
                        foreach($Attribute in $Attributes){
                            if($Attribute.Length -gt 7){
                                if($Attribute.StartsWith('src=',[System.StringComparison]::InvariantCultureIgnoreCase)){                                        
                                    $ParsedLinksObject.Images += ([URI]($Attribute.Substring(5,$Attribute.Length-6).Replace("`"","").Replace("'","").Replace("`r`n","")))   
                                } 
                            }                                 
                        
                        }
                    } 
                }catch{
                    Write-host ("Parse exception " + $_.Exception.Message + " on Message " + $Item.Subject)
                    $Error.Clear()
                }                       
            }  
        }            
        $Item | Add-Member -Name "ParsedLinks" -Value $ParsedLinksObject -MemberType NoteProperty -Force          
    }

}