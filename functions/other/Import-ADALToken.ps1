function Import-AccessToken{ 
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [string]$ClientId,
        [Parameter(Position=2, Mandatory=$false)] [string]$AccessToken,
        [Parameter(Position=3, Mandatory=$false)] [string]$RefreshToken,
        [Parameter(Position=4, Mandatory=$false)] [string]$ResourceURL,
        [Parameter(Position=5, Mandatory=$false)] [string]$redirectUrl,
        [Parameter(Position=6, Mandatory=$false)] [switch]$Beta
    )  
 	Begin
		 {
            $JsonObject = "" | Select token_type,scope,expires_in,ext_expires_in,expires_on,not_before,resource,access_token,clientid,redirectUrl
	    $Decoded =  Invoke-DecodeToken -Token $AccessToken
            $JsonObject.access_token = $AccessToken
            
            if(![String]::IsNullOrEmpty($RefreshToken)){
                Add-Member -InputObject $JsonObject -NotePropertyName refresh_token -NotePropertyValue (Get-ProtectedToken -PlainToken $RefreshToken) 
            }    
            if([bool]($JsonObject.PSobject.Properties.name -match "access_token")){
                $JsonObject.access_token =  (Get-ProtectedToken -PlainToken $JsonObject.access_token) 
            }
            $JsonObject.token_type = "Bearer"
            $JsonObject.scope = $Decoded.claims.scp
            $JsonObject.expires_on = $Decoded.claims.exp
            $JsonObject.not_before = $Decoded.claims.nbf
            $JsonObject.resource = $Decoded.claims.aud
            $JsonObject.clientid = $Decoded.claims.appid
            if($Beta.IsPresent){
                 Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
            }
            return $JsonObject
         }
}