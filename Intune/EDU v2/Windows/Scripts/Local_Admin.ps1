$userName = "MCSAdmin"
$Password = ConvertTo-SecureString "BLF&sax9(%k4L8" -AsPlainText -Force
$Userexist = (Get-LocalUser).Name -Contains $userName
if($userexist -eq $false) {
  try{ 
     New-LocalUser -Name $username -password $Password -Description "Local Workstation Admin Account" -PasswordNeverExpires -ErrorAction SilentlyContinue
     Add-LocalGroupMember -Group "Administrators" -Member $userName -ErrorAction SilentlyContinue
     Exit 0
   }   
  Catch {
     Exit 1
   }
}