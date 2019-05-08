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
' Function to determine if a value is in an array
' See: https://wellsr.com/vba/2016/excel/check-if-value-is-in-array-vba/
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError: 'array is empty
        For Each element In arr
            If element = valToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function
' Function to check if a field exists
Private Function DoesFieldExist(field As Variant, Table As Variant) As Boolean
   Dim exists As Boolean
   exists = False
   On Error Resume Next
   exists = CurrentDb.TableDefs(Table).Fields(field).Name = field
   DoesFieldExist = exists
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

'##############################################################################'
'######################## OTHER RESOURCES #####################################'
' @name TimeFromIFBGPS
' @author Bill DeVoe, MaineDMR, william.devoe@maine.gov
' @description iFormBuilder currently contains no method for parsing time from the Location
' widget; additionally, the time data is excluded from the JSON of form data, but is present
' in the Excel feeds. For projects needing to capture location and time in one element, this
' function allows the time to be parsed from the Location widget field in an Excel feed.
'
' @param GPStext {String} String of the GPS field.
' @return {String} The time component of the GPS field formatted as 24 hour time HH:MM:SS
Public Function TimeFromIFBGPS(GPStext As String) As String
    If IsEmpty(GPStext) Then
        TimeFromIFBGPS = ""
        Exit Function
    End If
    Dim substrings() As String
    ' Split by ,
    substrings = Split(GPStext, ",")
    Dim time As String
    ' Time field is the 7th element split by comma, take 12 characters to the right to get the time
    time = Right(Trim(substrings(6)), 12)
    ' Split by whitespace and take the first value to get rid of "EDT"
    substrings = Split(time, " ")
    TimeFromIFBGPS = Trim(substrings(0))
End Function

' @name DownloadIFBData
' @author Bill DeVoe, MaineDMR, william.devoe@maine.gov
' @description Downloads data from a given iFormBuilder form, including child forms, and creates/appends
' each form into a table in the source Access database. The first time the function is run, the tables
' will be created. Subsequent calls to the function will append data to the tables. Existing data should
' be cleared to avoid primary key violations. When the tables are created, the ID column will be set as the
' primary key. Relates/referential integrity can be added to the destination tables. Modifications to the
' schema of form data in IFB will result in an error, unless the schema is modified to be identical in the
' destination Access tables. Destination tables can be recreated by deleting them and recalling the function.
'
' IMPORTANT: At present, the username and password must be an account belonging to the profile of the parent
' form; server admin accounts spanning profiles will generate an error.
'
' For more information on how the feed URL is constructed, see here:
' https://iformbuilder.zendesk.com/hc/en-us/articles/202168664-How-do-I-export-my-data-to-XLSX-
'
' @dependencies -
'       - Module modArraySupport from www.cpearson.com
'       - Function IsInArray from wellsr.com
'       - Reference to reference to the Microsoft Excel x.x Object Library
'
' @param PageID - *Long Integer* - Page ID of the parent form
' @param User - *String* - Username belonging to the form profile with view rights to the form
' @param Password - *String* - Password for the username provided
' @param SplitName - *Optional Boolean* - If True, the form name will be split. If False, the form name will be used
'   for the destination table. Defaults to False.
' @param SplitBy - *Optional String* - If SplitBy is True, the character provided will be used to split the form name.
'   Defaults to "_"
' @param SplitSubscript - *Optional Integer* - If SplitBy is True, this index (starting with 0) of the array resulting
'   from the form name split will be used as the destination table name. Defaults to 1 (2nd item in array)
' @param Feed - *Optional Boolean* - If True, returns record metadata, column names, and option key values, versus
'   option list labels and no metadata. Defaults to True.
' @param Flatten - *Optional Boolean* - If True, related parent and child data is output to a single row in the same
'   worksheet. Defaults to False.
' @param ServerName - *Optional String* - Server name the form data is on. Defaults to "mainedmr"
' @return - *Boolean* - True if function successful with no errors, else False
Public Function DownloadIFBData(ByVal PageID As Long, _
                                ByVal User As String, _
                                ByVal Password As String, _
                                Optional ByVal SplitName As Boolean = False, _
                                Optional ByVal SplitBy As String = "_", _
                                Optional ByVal SplitSubscript As Integer = 1, _
                                Optional ByVal Feed As Boolean = True, _
                                Optional ByVal Flatten As Boolean = False, _
                                Optional ByVal ServerName As String = "mainedmr") As Boolean
    ' Function wide error handler
    On Error GoTo ErrorHandler
    ' Shut warnings off
    DoCmd.SetWarnings False
    ' Build URL to Excel feed of project data using settings in PSN table (TRIPPAGEID, IFBUSER, IFBPWD)
    Dim Url As String
    ' Build URL
    Url = "https://" & ServerName & ".iformbuilder.com/exzact/dataExcelViewV2.php?PAGE_ID=" _
        & PageID _
        & "&USERNAME=" & User _
        & "&PASSWORD=" & Password
    ' If feed is true, Returns record metadata, column names, and option key values
    If Feed Then
        Url = Url & "&FEED=1"
    End If
    ' If flatten is true, related parent and child data is output to a single row in the same worksheet
    If Flatten Then
        Url = Url & "&FLAT_VIEW=1"
    End If
    ' Debug output URL
    Debug.Print "URL to IFB Excel File: " & Url
    ' Path to a safe place to download the Excel file
    Dim SavePath As String
    ' Find the location of the users temp dir
    Dim temp_dir As String
    temp_dir = Environ("temp")
    ' Initial filename to save the download
    SavePath = temp_dir & "\IFB_Download.xlsx"
    ' Check if the file exists; if it does, try to delete it; if it cannot be deleted, try a different file name
    Dim filenum As Integer
    filenum = 1
    Dim repeat As Boolean
    repeat = True
    While repeat = True
        ' If the file exists
        If (Dir(SavePath) <> "") Then
            ' Try to delete it - sometimes it cannot be deleted (locks, etc)
            On Error Resume Next
            Kill SavePath
            On Error GoTo ErrorHandler
            ' If the file still exists
            If (Dir(SavePath) <> "") Then
                ' Make a new file name incremented sequentially
                filenum = filenum + 1
                SavePath = temp_dir & "\IFB_Download" & Str(filenum) & ".xlsx"
            End If
        Else
            ' File does not exist (success!!), exit the while loop
            repeat = False
        End If
    Wend
    
    ' Download file to the empty save path
    Dim RetVal As Integer
    RetVal = URLDownloadToFile(0, Url, SavePath, 0, 0)
    ' Output file location to debugger
    Debug.Print "IFB Excel File Successfully Saved To: " & SavePath
  
    ' Load sheets from Excel file to Access tbls - Requires reference to the Microsoft Excel x.x Object Library
    Dim objXL As New Excel.Application
    Dim Workbook As Excel.Workbook
    Dim Sheet As Object

    Set Workbook = objXL.Workbooks.Open(SavePath, Notify:=False, ReadOnly:=True)
    ' Array to hold the names of destination tables; used to make sure two sheets from the Excel file
    ' do not get sent to the same destination table
    Dim tblNames() As String
    For Each Sheet In Workbook.Worksheets
        Debug.Print "Loaded sheet: " & Sheet.Name
        Dim DestTable As String
        Dim DestTable_unsplit As String
        DestTable = Sheet.Name
        ' First figure out the form name without all the numbers on the end in the Excel sheet
        ' ie, we want my_awesome_form instead of my_awesome_form_123456789
        ' Split apart the form name by _ into an array
        Dim subs() As String
        subs = Split(Sheet.Name, "_")
        ' Then create a second array to slice the first array element to one less than the last element
        Dim subs2() As String
        Dim Result As Boolean
        ' Iterate over first array up to the 2nd to last item, adding it to the second array
        Dim x As Integer
        For x = 0 To UBound(subs) - 1
            ReDim Preserve subs2(x)
            subs2(x) = subs(x)
        Next x
        ' Collapse the new array to a string, separating each value by "_"
        DestTable_unsplit = Join(subs2, "_")
        Debug.Print "Loaded sheet from Excel file: " & DestTable_unsplit
        ' If Splitname then split form name by SplitBy and take SplitSubscipt value from array
        If SplitName Then
            ' Split sheet name by SplitBy
            subs = Split(Sheet.Name, SplitBy)
            ' SplitSubscript part of form name is table name, unless subscript is out of bounds then keep as is
            If SplitSubscript <= UBound(subs) Then
                DestTable = subs(SplitSubscript)
            Else
                DestTable = DestTable_unsplit
            End If
        Else
            ' Use the unsplit form name
            DestTable = DestTable_unsplit
        End If
        ' Check if the dest table has already been loaded
        If IsInArray(DestTable, tblNames) Then
            MsgBox "The table name resulting from the specified form name split resulted in a duplicate destination table name. " & _
                "The source sheet will be inserted into a table of the same name: " & DestTable_unsplit
            ' Redirect destination table
            DestTable = DestTable_unsplit
        End If
        ' Add dest table to tblNames array
        Dim new_len As Integer
        ' If the array is empty
        If IsArrayEmpty(arr:=tblNames) = True Then
            new_len = 0
        Else
            new_len = UBound(tblNames) + 1
        End If
        ReDim Preserve tblNames(new_len)
        tblNames(new_len) = DestTable
        ' Load the sheet to a table of the table name
        ' First check if the table already exists
        Dim DestExists As Boolean
        DestExists = False
        If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & DestTable & "'")) Then DestExists = True
        ' Load the data into the table
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
                                    DestTable, SavePath, True, Sheet.Name & "$"
        ' If the table was just created, make the ID column the primary key and set the PARENT_RECORD_ID column as integer not double
        If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & DestTable & "'")) And DestExists = False Then
            ' Make the ID column the primary key; cast ID and PARENT_RECORD_ID as integer
            DoCmd.SetWarnings False
            DoCmd.RunSQL "ALTER TABLE " & DestTable & " ALTER COLUMN ID INTEGER CONSTRAINT PK_" & DestTable & " PRIMARY KEY"
            DoCmd.RunSQL "ALTER TABLE " & DestTable & " ALTER COLUMN PARENT_RECORD_ID INTEGER"
            DoCmd.SetWarnings True
        End If
        Debug.Print "Excel sheet " & Sheet.Name & " loaded into table " & DestTable
' Next sheet in the workbook
    Next Sheet

    ' Cleanup, close the workbook, etc
    Workbook.Close
    Set Workbook = Nothing
    Set Sheet = Nothing
    objXL.Quit
    Set objXL = Nothing
    ' Return true
    DownloadIFBData = True
    Exit Function
' Error handler
ErrorHandler:
    DownloadIFBData = False
    Debug.Print Err.Description
    MsgBox Err.Description
End Function
