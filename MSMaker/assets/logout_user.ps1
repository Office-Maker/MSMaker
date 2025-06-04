Install-Module Microsoft.Graph -Force
Import-Module Microsoft.Graph.Users.Actions
Connect-MgGraph -Scopes 'User.Read.All', 'User.ReadWrite.All', 'User.RevokeSignInSession.All'
Revoke-MgUserSignInSession -UserId "user@example.com"
Disconnect-MgGraph
