# VBA-iFormBuilder
Resources for utilizing the Zerion iFormBuilder API in VBA.

## Dependencies

Communicating with the iFormBuilder APIs requires parsing JSON data, in addition to encoding text in Base64 and hashing with SHA256. All of these capabilities are provided by [VBA-Web](https://github.com/VBA-tools/VBA-Web/) from VBA-Tools.

## Obtaining an API token

```visual basic
' Load variables
    Dim Client_Key As String
    Client_Key = "myapikey"
    Dim Client_Secret As String
    Client_Secret = "myapisecret"
    Dim Server_Name As String
    Server_Name = "myservername"
    
' Get an access token to the IFB API
    Dim Access_Token As String
    Access_Token = Get_iForm_Access_Token(Server_Name, Client_Key, Client_Secret)
    If Access_Token = "" Then
        MsgBox "An access token to the iFormBuilder API could not be generated."
    End If 
```
