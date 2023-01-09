$username = $env:username
$saccount = $username + "s"

$ACL = Get-ACL -Path "C:\Users\$username"
$permission = "humad\$saccount", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission.Trim()
$ACL.SetAccessRule($AccessRule)
$ACL | Set-Acl -Path "C:\Users\$username"
(Get-ACL -Path "C:\Users\$username").Access | Format-Table IdentityReference, FileSystemRights, AccessControlType, IsInherited, InheritanceFlags -AutoSize