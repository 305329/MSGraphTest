#Input Parameters: 
$csv = Import-Csv 'Z:\test01\Users.csv'
$folderpath="Z:\test01\" 

#Download user profile pictures for specific users from Office 365 via CSV file:
New-Item -ItemType directory -Path $folderpath –force 
Foreach($user in $csv)
{
$userName=($user.UserPrincipalName -Split "@")[0] 
$path=$folderpath+$userName+".Jpg"
$photo=Get-Userphoto -identity $user.UserPrincipalName -ErrorAction SilentlyContinue
If($photo.PictureData -ne $null)
{
[io.file]::WriteAllBytes($path,$photo.PictureData)
Write-Host $userName “profile picture downloaded”
}
}