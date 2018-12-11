To use Certificate Authentication to create an App Only token you need to create a self signed certificate and add that to the application manifest in Azure. There is some Microsoft documentation on doing this [here](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread) ,

Or you can use the Exch-Rest module which  has a cmdlet that can do everything that is required if you pass in the objectId from the Azure Registration you created eg the following is need from the Azure app portal

![](https://cdn-images-1.medium.com/max/800/1*a0uUctPmKETsHhFj8TbppQ.jpeg)

With this object Id if you execute the following it will create a Self signed certificate and add that to the manifest of your application
Invoke-EXRCreateAppTokenCertificate -CertName PDLCert -CertFileName c:\temp\PDLCert.cer -ObjectId 9319a335-6f8a-4049-89af-b43bb625239e
(as long as the appId you are running the Exch-Rest under has the following delegated permissions)

![](https://cdn-images-1.medium.com/max/800/1*LfrUuftcntLt3za3dNSNfA.jpeg)

You can then assert the tenant admin rights of the user you have actually signed in as other wise you will just get a 403 error when trying to modify the application registration.

To Logon using the Certificate you can use the following

     Connect-EXRMailbox -MailboxName gscales@datarumble.com -certificateFileName C:\temp\PDLCert.cer -clientId "450ce1c4-5a75-447a-a67b-65031430cd7f"

If you want to also specify the Certificate password you can do that using a SecureString

    $password = Read-Host -AsSecureString
	Connect-EXRMailbox -MailboxName gscales@datarumble.com -certificateFileName C:\temp\PDLCert.cer -clientId "450ce1c4-5a75-447a-a67b-65031430cd7f -certificateFilePassword $password