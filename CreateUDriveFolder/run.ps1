Import-Module ActiveDirectory
 $path = "\\urbanengineers.local\files\urban\Users\"
 $newFolderName = Read-Host -Prompt "Enter Name of User"
 $newFolderFull = $path + $newFolderName
 
 Write-Output "New Folder will be: $newFolderFull"
 $confirm = Read-Host "Confirm? Y/N"
 If(($confirm) -ne "y")
 {

 }
 Else
 {
    #Create Directory Structure
    #Main Folder
    Write-Output "Add Folder.."
    New-Item $newFolderFull -ItemType Directory


    #Write-Output "Remove Inheritance.."
    #icacls $newFolderFull /inheritance:d
    
    $acl = Get-ACL -Path $newFolderFull
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($newFolderName,"Modify","ContainerInherit,ObjectInherit","None","Allow")
    
    $acl.SetAccessRuleProtection($True, $True)
    $acl.SetAccessRule($AccessRule)
    Set-Acl -Path $newFolderFull -AclObject $acl
 }