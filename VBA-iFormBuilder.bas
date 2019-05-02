Attribute VB_Name = "VBA-iFormBuilder_export"
Option Compare Database
Option Explicit
'##############################################################################'
'######################## PRIVATE FUNCTIONS ###################################'
' @name Base64_URL
' @param Text {String} Base64 encoded text to make URL safe
' @return The input Base64 encoded text converted into URL safe characters.
Private Function Base64_URL(Text As String)
    Base64_URL = Replace(Text, "+", "-")
    Base64_URL = Replace(Base64_URL, "/", "_")
    Base64_URL = Replace(Base64_URL, "=", "")
End Function
' @name GetUnixEpoch
' @return The current Unix epoch (seconds since midnight January 1, 1970 GMT)
Private Function GetUnixEpoch() As Long
    Dim dt As Object
    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate Now
    GetUnixEpoch = DateDiff("S", "1/1/1970", dt.GetVarDate(False))
End Function
' @name Base_URL
' @param Server_Name {String} Name of the iFormBuilder server
' @return {String} Base URL to iFormBuilder
Private Function Base_URL(ByVal Server_Name As String) As String
    Base_URL = "https://" & Server_Name & ".iformbuilder.com"
End Function
' @name API_v60_URL
' @param Server_Name {String} Name of the iFormBuilder server
' @return {String} Base URL to iFormBuilder API
Private Function API_v60_URL(ByVal Server_Name As String) As String
    API_v60_URL = Base_URL(Server_Name) & "/exzact/api/v60/profiles/"
End Function
'##############################################################################'
'######################## TOKEN RESOURCES # ###################################'

' @name Get_iForm_Access_Token
' @author Bill DeVoe, MaineDMR, william.devoe@maine.gov
' @description Generates an access token for the iFormBuilder API.
'
' @dependencies -
'       - VBA-WEB from VBA-TOOLS (includes VBA-JSON and UTC functions) from
'               https://github.com/VBA-tools/VBA-Web/
' @param Server_Name {String} iFormBuilder server name
' @param Client_Key {String} Client API key
' @param Client_Secret {String} Client API secret
' @return {String} iForm API access token
Function Get_iForm_Access_Token(ByVal Server_Name As String, _
                                ByVal Client_Key As String, _
                                ByVal Client_Secret As String) As String
    ' Build URL to get token
    Dim token_url As String
    token_url = "https://" & Server_Name & ".iformbuilder.com/exzact/api/oauth/token"

    ' Build JWT header
    Dim jwt_header As Object
    Set jwt_header = ParseJson("{}")
    jwt_header("alg") = "HS256"
    jwt_header("typ") = "JWT"
    
    ' Build JWT payload
    Dim jwt_claim As Object
    Set jwt_claim = ParseJson("{}")
    jwt_claim("iss") = Client_Key
    jwt_claim("aud") = token_url
    ' Current Unix time
    jwt_claim("iat") = GetUnixEpoch()
    ' Expires in 10 minutes
    jwt_claim("exp") = GetUnixEpoch() + 600
    
    ' Base sign
    Dim base_sign As String
    base_sign = Base64_URL(Base64Encode(ConvertToJson(jwt_header))) & _
                "." & Base64_URL(Base64Encode(ConvertToJson(jwt_claim)))
    ' Sign the base with the client secret
    Dim signature As String
    signature = Base64_URL(HMACSHA256(base_sign, Client_Secret, "Base64"))
    ' Then add signature to base sign to make JWT
    Dim jwt As String
    jwt = base_sign & "." & signature
    ' Build API call
    Dim xhr As Object
    Dim body_str As String
    Dim json_Text As String
    Dim json As Object
    ' Build query body - uses form encoding, NOT JSON encoding
    body_str = "grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=" & jwt
    ' Use late binding
    Set xhr = CreateObject("MSXML2.serverXMLHTTP")
    xhr.Open "POST", token_url, False
    xhr.SetRequestHeader "Content-Type", "x-www-form-urlencoded"
    xhr.Send (body_str)
    ' If request successful
    If xhr.Status = 200 Then
        json_Text = xhr.ResponseText
    ' Could not connect to service, 404 error or 400 bad request -
    ' ie, no internet connection or invalid url
    Else
        MsgBox xhr.Status & ": " & xhr.StatusText
        MsgBox "Invalid Response From Server"
        Exit Function
    End If
    Set xhr = Nothing
    ' Parse response
    Set json = ParseJson(json_Text)
    ' Get the token
    Get_iForm_Access_Token = json("access_token")
End Function

'##############################################################################'
'######################## RECORD RESOURCES ####################################'

' @name Delete_Records
' @author Bill DeVoe, MaineDMR, william.devoe@maine.gov
' @description Deletes a list of up to 100 record IDs from an IFB page.
'
' @dependencies -
'       - VBA-WEB from VBA-TOOLS (includes VBA-JSON and UTC functions) from
'               https://github.com/VBA-tools/VBA-Web/
' @param Server_Name {String} iFormBuilder server name
' @param Profile_ID {Integer} iFormBuilder profile ID
' @param Access_Token {String} Access token produced by Get_iForm_Access_Token
' @param Page_ID {Integer} ID of the page from which to delete the records
' @param Record_IDs {Variant} Integer array of the record IDs to delete.
' @return {String} JSON array of the deleted record IDs.
Function Delete_Records(ByVal Server_Name As String, _
                        ByVal Profile_ID As Long, _
                        ByVal Access_Token As String, _
                        ByVal Page_ID As Long, _
                        ByVal Record_IDs As Variant) As String
    ' Offset for URL call
    Dim offset As Integer
    offset = 100
    ' Build URL for API call
    Dim request_url As String
    request_url = API_v60_URL(Server_Name) & Profile_ID & "/pages/" _
        & Page_ID & "/records?fields=&limit=100&offset=" & offset
    ' Convert record IDs array to JSON array
    Dim body_str As String
    Dim i As Integer
    body_str = "["
    For i = LBound(Record_IDs) To UBound(Record_IDs)
        If i <> UBound(Record_IDs) Then
            body_str = body_str & "{""id"": " & Record_IDs(i) & "},"
        Else
            body_str = body_str & "{""id"": " & Record_IDs(i) & "}]"
        End If
    Next i
    ' Bearer and call
    Dim bearer As String
    bearer = "Bearer " & Access_Token
    Dim xhr As Object
    ' Use late binding
    Set xhr = CreateObject("MSXML2.serverXMLHTTP")
    ' Use DELETE method
    xhr.Open "DELETE", request_url, False
    xhr.SetRequestHeader "Content-Type", "application/json"
    xhr.SetRequestHeader "Authorization", bearer
    xhr.Send (body_str)
    Dim json_Text As String
    ' If request successful
    If xhr.Status = 200 Then
        json_Text = xhr.ResponseText
    ' Could not connect to service, 404 error or 400 bad request -
    ' ie, no internet connection or invalid url
    Else
        MsgBox xhr.Status & ": " & xhr.StatusText
        MsgBox "Invalid Response From Server"
        Delete_Records = "[]"
        Exit Function
    End If
    Set xhr = Nothing
    ' Parse response
    Delete_Records = json_Text
End Function



