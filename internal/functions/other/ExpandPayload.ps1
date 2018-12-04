function ExpandPayload {
    [CmdletBinding()]
    Param (
        $response
    )
    if ($PSVersionTable.PSEdition -eq "Core") {
		ConvertFrom-JsonNewtonsoft $response
    }
    else {
        ## Start Code Attribution
        ## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
        ## https://www.powershellgallery.com/profiles/chriswahl/
        ## End Code Attribution
        [void][System.Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
        return ParseItem -jsonItem ((New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer -Property @{
                    MaxJsonLength = [Int32]::MaxValue
                }).DeserializeObject($response))
    }
}

function ConvertFrom-JsonNewtonsoft {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, ValueFromPipeline = $true)]$string) 
	    ## Start Code Attribution
        ## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
        ## https://www.powershellgallery.com/profiles/chriswahl/
        ## End Code Attribution
    $HandleDeserializationError = 
    {
        param ([object] $sender, [Newtonsoft.Json.Serialization.ErrorEventArgs] $errorArgs)
        $currentError = $errorArgs.ErrorContext.Error.Message
        write-warning $currentError
        $errorArgs.ErrorContext.Handled = $true
        
    }

    $settings = new-object "Newtonsoft.Json.JSonSerializerSettings"
    if ($ErrorActionPreference -eq "Ignore") {
        $settings.Error = $HandleDeserializationError
    }
    $obj = [Newtonsoft.Json.JsonConvert]::DeserializeObject($string, [Newtonsoft.Json.Linq.JObject], $settings)    

    return ConvertFrom-JObject $obj
}

function ConvertFrom-JObject($obj) {
	## Start Code Attribution
    ## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://www.powershellgallery.com/profiles/chriswahl/
    ## End Code Attribution
    if ($obj -is [Newtonsoft.Json.Linq.JArray]) {
        $a = foreach ($entry in $obj.GetEnumerator()) {
            @(convertfrom-jobject $entry)
        }
        return $a
    }
    elseif ($obj -is [Newtonsoft.Json.Linq.JObject]) {
        $h = [ordered]@{}
        foreach ($kvp in $obj.GetEnumerator()) {
            $val = convertfrom-jobject $kvp.value
            if ($kvp.value -is [Newtonsoft.Json.Linq.JArray]) { $val = @($val) }
            $h += @{ "$($kvp.key)" = $val }
        }
        return [pscustomobject]$h
    }
    elseif ($obj -is [Newtonsoft.Json.Linq.JValue]) {
        return $obj.Value
    }
    else {
        return $obj
    }
}
