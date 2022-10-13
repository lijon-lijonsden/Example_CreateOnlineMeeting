# Example_CreateOnlineMeeting

Microsoft Graph /onlineMeeting requires special authority to create meetings on behalf of an identity

--- PowerShell instructions ---

Possibly required to execute scripts in your environment:
Set-ExecutionPolicy RemoteSigned



Install-Module -Name PowerShellGet -Force -AllowClobber


Install-Module -Name MicrosoftTeams -Force -AllowClobber



Import-Module MicrosoftTeams



Create new policy for the application:
New-CsApplicationAccessPolicy -Identity {policyName} -AppIds "{appId}" -Description "{description}"



Grant identity: 
Grant-CsApplicationAccessPolicy -PolicyName {policyName} -Identity "{identity}"



For tenant global grant:
Grant-CsApplicationAccessPolicy -PolicyName Test-policy -Global



Clean up, if execution policy was changed:
Set-ExecutionPolicy Restricted

