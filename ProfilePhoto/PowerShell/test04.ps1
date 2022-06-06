####################### Microsoft365 Profile photo download setting
$Basedir = "Z:\test01"
$TargetEmpCSVFile = "$basedir\target\Users.csv"
$OutputDir = "$basedir\downloadPhoto"
$LogDir = "$basedir\log"

Write-Output "BaseDir = $Basedir"
Write-Output "TargetEmpCSVFile = $TargetEmpCSVFile"
Write-Output "OutputDir = $OutputDir"
Write-Output "LogDir = $LogDir"

####################ログの取得
Start-Transcript $LogDir\log.txt -Append
####################GraphAPI接続 start
$clientid = "b0dc427a-4baf-4545-8e30-04f1cd69b81d"
$tenantName = "b1380395-f290-46e0-9390-cd896bd4ae16"
$clientSecret = "tQ5MNnM+T2foO/SuOOvsFRs80ZNgW0R5IXqB1649oOQ="

# Webサンプル
$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = ${clientID}
    Client_Secret = ${clientSecret}
} 

#PostmanのときのHTTPパラメータ
# $ReqTokenBody2 = @{
    # Grant_Type    = "client_credentials"
    # resource      = "https://graph.microsoft.com/"
    # client_Id     = ${clientID}
    # Client_Secret = ${clientSecret}
# } 

Write-Output "--- Azure Login ---"
try {
	$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method POST -Body ${ReqTokenBody}
}
catch {
	//HTTPステータスコード
	Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__
	Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
}
Write-Output ${TokenResponse}
#######################GraphAPI接続 end

#######################GraphAPIで情報取得
$csv = Import-Csv $TargetEmpCSVFile

Write-Output "--- Get Profile Photo ---"

foreach ($user in $csv) {
	$EmpID=($user.UserPrincipalName -Split "@")[0] 
	Write-Output "================================="
	Write-Output "Downloading user = $user"
	Write-Output "Downloading EmpID = $EmpID"

	$apiUrl = "https://graph.microsoft.com/beta/users/$($user.UserPrincipalName)/photos/504x504/`$value"
	Write-Output ${apiUrl}

	try {
		$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get -ContentType "image/jpeg" -OutFile $OutputDir\$EmpID.jpeg
	}
	catch {
		Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__
		Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
	}
		$Data
}

