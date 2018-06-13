function Get-EXREntryIdFromOWAId {
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]$DomainName
       
    )  
    Begin {
        try{
            $RequestURL = "https://login.windows.net/{0}/.well-known/openid-configuration" -f $DomainName
            $Response = Invoke-WebRequest -Uri  $RequestURL
            $JsonResponse = ConvertFrom-Json  $Response.Content
            $ValArray = $JsonResponse.authorization_endpoint.replace("https://login.windows.net/","").split("/")
            return $ValArray[0]
        }catch{

        }

    }
}