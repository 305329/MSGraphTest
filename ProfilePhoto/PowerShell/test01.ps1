#######################Target User
$TargetEmp = "305329"
# $TargetEmp = "304231" #角さん
#######################GraphAPI接続 start
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

$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method POST -Body ${ReqTokenBody}
# $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/${tenantName}/oauth2/token" -Method POST -Body ${ReqTokenBody}
# $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($tenantName)/oauth2/token" -Method POST -Body ${ReqTokenBody2}
# Write-Output ${TokenResponse}
#######################GraphAPI接続 end

#######################GraphAPIで情報取得
Write-Output $TargetEmp
$apiUrl = "https://graph.microsoft.com/beta/users/${TargetEmp}@kao.com/photo/$value"
# $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $(Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
# $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get -OutFile z:\test01\$targetEmp.jpg
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get -ContentType "image/jpg" -OutFile z:\test01\$targetEmp.jpg

$Data
