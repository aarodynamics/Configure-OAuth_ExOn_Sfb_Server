<#
    .SYNOPSIS
    Configure-OAuth_ExOn_SfB_Server
   
    Aaron Marks
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF USE OR RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 2.0, 5/26/2016
    
    .DESCRIPTION
	Configure Oauth between Exchange Online and Skype for Business Server.
	
    .LINK
    http://www.turnpointtech.com
    
    .NOTES
	
    Requirements:
		- Must be ran from Skype4B Server
		- Run from Elevated AAD PowerShell
		- MS Online Service Sign-in Assistant RTW: 
			http://go.microsoft.com/fwlink/?LinkID=286152 
		- AAD PowerShell: 
			http://go.microsoft.com/fwlink/p/?linkid=236297 
	
	Special thanks to Christian Burke for providing the 
	commands used to assemble this script: http://goo.gl/25ZZfK
	
	And, thanks to @PatRichard for writing the part to export the OAuth certificate.
	
	And, thanks to Kory Olson & Tony Quintanilla for suggesting to automate grabbing
	the web external URL and modifying the script to support multiple external URLs.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release
	2.0     Auto-export of OAuth Cert
    2.1     Added the ability to utilize more than one Web External Url
   
	
    .PARAMETER WebExt
	Web External url from Skype for Business Front End Pool(s).  This will take 
    single or multiple entries
    
	.EXAMPLE
	.\Configure-OAuth_ExOn_SfB_Server.ps1
	
	.EXAMPLE
    .\Configure-OAuth_ExOn_SfB_Server.ps1 -WebExt "webext.contoso.com"

    .EXAMPLE
    .\Configure-OAuth_ExOn_SfB_Server.ps1 -WebExt "webext.contoso.com","webext2.contoso.com"

#>

[CmdletBinding(DefaultParameterSetName="Default")]

Param(
	[Parameter(Mandatory=$False, ParameterSetName="Default")]
		$WebExt
) #Param

Process {
	
	if ($WebExt) {
		Write-Verbose "Using custom WebExt parameter"
	} else {
		$WebExt = ((Get-CsService).ExternalFqdn)
	}
	
	$CertPath = "C:\oauth.cer"
	$Thumbprint = (Get-CsCertificate -Type OAuthTokenIssuer).Thumbprint
	$OAuthCert = Get-ChildItem -Path Cert:\LocalMachine\My\$Thumbprint
	Export-Certificate -Cert $OAuthCert -FilePath $CertPath -Type CERT

	
	# Script
	Import-Module LyncOnlineConnector

	# Logging into SfB Online
	Write-Output "Logging into SfB Online"
	$cred = Get-Credential

	$CSSession = New-CsOnlineSession -Credential $cred 

	Import-PSSession $CSSession -AllowClobber 

	# Clean up Old Entries if necessary
	Write-Output "Clean up Old Entries if necessary"

	If ( (Get-CsOauthServer).Identity -match "microsoft.sts") {
		Write-Output "Removing existing microsoft.sts OauthServer"
		Remove-CsOauthServer -Identity microsoft.sts
	}

	If ( (Get-CsPartnerApplication).Identity -match "microsoft.exchange") {
		Write-Output "Removing existing microsoft.exchange PartnerApplication"
		Remove-CsPartnerApplication -Identity microsoft.exchange
	}

	# Create New OAuth Server
	Write-Output "Create New OAuth Server"

	$TenantId = (Get-CsTenant).TenantId
	$metadataurl = "https://accounts.accesscontrol.windows.net/" `
	+ "$TenantId" + "/metadata/json/1"
	New-CsOAuthServer -Identity microsoft.sts -metadataurl $metadataurl

	# Create New Partner Application
	Write-Output "Create New Partner Application"

	New-CsPartnerApplication -Identity microsoft.exchange `
	-ApplicationIdentifier 00000002-0000-0ff1-ce00-000000000000 `
	-ApplicationTrustLevel Full -UseOAuthServer

	#Set-CsOAuthConfiguration -ServiceName 00000004-0000-0ff1-ce00-000000000000

	# Elevate the SfB PowerShell 
	Write-Output "Elevate the SfB PowerShell"

	Import-Module MSOnlineExtended

	Connect-MsolService -Credential $cred

	Get-MsolServicePrincipal

	# Upload your OAuth Certificate
	Write-Output "Upload your OAuth Certificate"

	$certificate = 
	New-Object System.Security.Cryptography.X509Certificates.X509Certificate

	$certificate.Import("$CertPath") 

	$binaryValue = $certificate.GetRawCertData() 

	$credentialsValue = [System.Convert]::ToBase64String($binaryValue) 

	New-MsolServicePrincipalCredential `
	-AppPrincipalId 00000004-0000-0ff1-ce00-000000000000 `
	-Type Asymmetric -Usage Verify -Value $credentialsValue 

	Set-MSOLServicePrincipal -AppPrincipalID `
	00000002-0000-0ff1-ce00-000000000000 -AccountEnabled $true 

	# Add Your Lync External Web Services Name
	Write-Output "Add Your Lync External Web Services Name"

	$lyncSP = Get-MSOLServicePrincipal -AppPrincipalID `
	00000004-0000-0ff1-ce00-000000000000 
    
    # Loop Through each WebExt discovered
    ForEach ($Fqdn in $WebExt){	
        $WebExtSpn = "00000004-0000-0ff1-ce00-000000000000/" + "$Fqdn"
        #Check to see if the SPN is already there.  If not, Add
        if ($lyncSP.ServicePrincipalNames -notcontains $WebExtSpn){
	        $lyncSP.ServicePrincipalNames.Add("$WebExtSpn")
        }
    }

	Set-MSOLServicePrincipal -AppPrincipalID `
	00000004-0000-0ff1-ce00-000000000000 `
	-ServicePrincipalNames $lyncSP.ServicePrincipalNames 
	
	Remove-PSSession $CSSession
	
} #Process