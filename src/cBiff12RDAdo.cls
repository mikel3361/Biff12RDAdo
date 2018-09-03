VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiff12RDAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "cBiff12RDAdo"

Private m_sFilename    As String
Private m_sSheetName   As String
Private m_lSheetIndex  As Long
Private m_oConn        As Object
Private m_oRS          As Object
Private m_sSheets      As String
Private m_sFields      As String
Private m_lFieldCount  As Long
Private m_oRecords     As cBiff12RDRows
Private m_oFields      As cBiff12RDFields
Private m_oSheets      As cBiff12RDSheets
Private m_lSheetCount  As Long
Private m_bHasHeader   As Boolean
Private m_bNoEmptyRows As Boolean
Private m_sLastError   As String
'''''''''''''''''''''''''''''''''''
' Used by FormatMessage
'''''''''''''''''''''''''''''''''''
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE  As Long = &H800
Private Const FORMAT_MESSAGE_FROM_STRING  As Long = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN  As Long = &HA0 ' from VC++ ERRORS.H file

'''''''''''''''''''''''''''''''''''
' Windows API Declare
'''''''''''''''''''''''''''''''''''
Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long

Private Function GetSystemErrorMessageText(ErrorNumber As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
'
' This function gets the system error message text that corresponds
' to the error code parameter ErrorNumber. This value is the value returned
' by Err.LastDLLError or by GetLastError, or occasionally as the returned
' result of a Windows API function.
'
' These are NOT the error numbers returned by Err.Number (for these
' errors, use Err.Description to get the description of the error).
'
' In general, you should use Err.LastDllError rather than GetLastError
' because under some circumstances the value of GetLastError will be
' reset to 0 before the value is returned to VBA. Err.LastDllError will
' always reliably return the last error number raised in an API function.
'
' The function returns vbNullString is an error occurred or if there is
' no error text for the specified error number.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ErrorText As String
Dim TextLen As Long
Dim FormatMessageResult As Long
Dim LangID As Long

''''''''''''''''''''''''''''''''
' Initialize the variables
''''''''''''''''''''''''''''''''
LangID = 0&   ' Default language
ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
TextLen = FORMAT_MESSAGE_TEXT_LEN

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Call FormatMessage to get the text of the error message text
' associated with ErrorNumber.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FormatMessageResult = FormatMessage( _
                        dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS, _
                        lpSource:=0&, _
                        dwMessageId:=ErrorNumber, _
                        dwLanguageId:=LangID, _
                        lpBuffer:=ErrorText, _
                        nSize:=TextLen, _
                        Arguments:=0&)

If FormatMessageResult = 0& Then
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' An error occured. Display the error number, but
    ' don't call GetSystemErrorMessageText to get the
    ' text, which would likely cause the error again,
    ' getting us into a loop.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    MsgBox "An error occurred with the FormatMessage" & _
           " API function call." & vbCrLf & _
           "Error: " & CStr(Err.LastDllError) & _
           " Hex(" & Hex(Err.LastDllError) & ")."
    GetSystemErrorMessageText = "An internal system error occurred with the" & vbCrLf & _
        "FormatMessage API function: " & CStr(Err.LastDllError) & ". No futher information" & vbCrLf & _
        "is available."
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If FormatMessageResult is not zero, it is the number
' of characters placed in the ErrorText variable.
' Take the left FormatMessageResult characters and
' return that text.
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ErrorText = Left$(ErrorText, FormatMessageResult)
'''''''''''''''''''''''''''''''''''''''''''''
' Get rid of the trailing vbCrLf, if present.
'''''''''''''''''''''''''''''''''''''''''''''
If Len(ErrorText) >= 2 Then
    If Right$(ErrorText, 2) = vbCrLf Then
        ErrorText = Left$(ErrorText, Len(ErrorText) - 2)
    End If
End If

''''''''''''''''''''''''''''''''
' Return the error text as the
' result.
''''''''''''''''''''''''''''''''
GetSystemErrorMessageText = ErrorText

End Function

Public Property Get FileName() As String
    FileName = m_sFilename
End Property
Public Property Let FileName(NewVal As String)
   Dim fs As Object, fle As Object
   Err.Clear
   m_sLastError = ""
   On Error Resume Next
   Set fs = CreateObject("Scripting.FileSystemObject")
   If Not fs Is Nothing Then
     Set fle = fs.GetFile(NewVal)
   End If
   If Err.Number <> 0 Then
     m_sLastError = "cBiff12Ado.FileName: " & Err.Description
   End If
   Set fle = Nothing
   Set fs = Nothing
   On Error GoTo 0
   m_sFilename = NewVal
   m_sSheetName = ""
   m_sSheets = ""
   m_sFields = ""
   m_oFields.Clear
   m_lFieldCount = 0
   m_oSheets.Clear
   m_lSheetCount = 0
   Set m_oRecords = New cBiff12RDRows
   If m_sFilename <> "" And m_sLastError = "" Then
     Call GetSheets
   End If
End Property

Public Property Get SheetName() As String
    SheetName = Replace(m_sSheetName, "$", "")
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

Public Property Get SheetIndex() As Long
    SheetIndex = m_lSheetIndex
End Property
Public Property Let SheetIndex(ByVal NewVal As Long)
    If NewVal > 0 And NewVal <= m_lSheetCount Then
      m_lSheetIndex = NewVal
    Else
      m_lSheetIndex = 0
    End If
    m_sLastError = ""
    m_sSheetName = GetSheetName(m_lSheetIndex)
    m_sSheets = ""
    m_sFields = ""
    m_sLastError = ""
    m_oFields.Clear
    m_lFieldCount = 0
    Set m_oRecords = New cBiff12RDRows
    If m_sSheetName <> "" Then
      Call GetFields
      If m_lFieldCount > 0 Then
        Call GetRecords
      End If
    End If
End Property

Public Property Get SheetCount() As Long
    SheetCount = m_lSheetCount
End Property

Public Function GetSheetName(ByVal Index As Long) As String
  If Index > 0 And Index <= m_lSheetCount Then
    GetSheetName = m_oSheets.Sheets(Index).SheetName
  Else
    GetSheetName = ""
  End If
End Function

Public Function GetSheetIndex(ByVal SheetName As String) As Long
  Dim ret As Long, idx As Long
  If SheetName <> "" And m_lSheetCount > 0 Then
    ret = 0
    For idx = 1 To m_lSheetCount
      If LCase(Replace(m_oSheets.Sheets(idx).SheetName, "$", "")) = LCase(SheetName) Then ret = idx: Exit For
    Next
    GetSheetIndex = ret
  Else
    GetSheetIndex = 0
  End If
End Function

Public Function GetFieldName(ByVal Index As Long) As String
  If Index > 0 And Index <= m_lFieldCount Then
    GetFieldName = m_oFields.Fields(Index).FieldName
  Else
    GetFieldName = ""
  End If
End Function

Public Function GetFieldType(ByVal Index As Long) As String
  If Index > 0 And Index <= m_lFieldCount Then
    GetFieldType = m_oFields.Fields(Index).FieldType
  Else
    GetFieldType = ""
  End If
End Function

Public Function GetFieldIndex(ByVal FieldName As String) As Long
  Dim ret As Long, idx As Long
  If FieldName <> "" And m_lFieldCount > 0 Then
    ret = 0
    For idx = 1 To m_lFieldCount
      If LCase(m_oFields.Fields(idx).FieldName) = LCase(FieldName) Then ret = idx: Exit For
    Next
    GetFieldIndex = ret
  Else
    GetFieldIndex = 0
  End If
End Function

Public Property Get FieldCount() As Long
    FieldCount = m_lFieldCount
End Property

Public Property Get HasHeader() As Boolean
    HasHeader = m_bHasHeader
End Property
Public Property Let HasHeader(ByVal NewVal As Boolean)
    m_sLastError = ""
    m_sFields = ""
    m_oFields.Clear
    m_lFieldCount = 0
    Set m_oRecords = New cBiff12RDRows
    m_bHasHeader = NewVal
    If m_sSheetName <> "" Then
      Call GetFields
      Call GetRecords
    End If
End Property

Public Property Get NoEmptyRows() As Boolean
    HasHeader = m_bNoEmptyRows
End Property
Public Property Let NoEmptyRows(ByVal NewVal As Boolean)
    m_sLastError = ""
    Set m_oRecords = New cBiff12RDRows
    m_bNoEmptyRows = NewVal
    If m_sSheetName <> "" Then
      Call GetRecords
    End If
End Property

Public Property Get Records() As cBiff12RDRows
   Set Records = m_oRecords
End Property

Private Sub Class_Initialize()
    m_sFilename = ""
    m_sSheetName = ""
    m_sSheets = ""
    m_sFields = ""
    m_lFieldCount = 0
    m_lSheetCount = 0
    Set m_oRecords = New cBiff12RDRows
    Set m_oFields = New cBiff12RDFields
    Set m_oSheets = New cBiff12RDSheets
    m_bHasHeader = False
    m_bNoEmptyRows = False
    m_sLastError = ""
End Sub

Private Sub Class_Terminate()
    m_sFilename = ""
    m_sSheetName = ""
    m_sSheets = ""
    m_sFields = ""
    m_lFieldCount = 0
    m_lSheetCount = 0
    Set m_oRecords = Nothing
    Set m_oFields = Nothing
    Set m_oSheets = Nothing
    m_bHasHeader = False
    m_bNoEmptyRows = False
    m_sLastError = ""
End Sub

Private Sub GetSheets()
   Err.Clear
   m_sLastError = ""
   m_sSheets = ""
   m_oSheets.Clear
   m_lSheetCount = 0
   On Error Resume Next
   Set m_oConn = CreateObject("ADODB.Connection")
   If Err.Number <> 0 Then
     m_sLastError = "ADODB.Connection: " & Err.Description
   End If
   On Error GoTo 0
   If Not m_oConn Is Nothing Then
     'Connection für Datatype und Schema
     m_oConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source='" & m_sFilename & "';Extended Properties='Excel 12.0;HDR=NO'; "
     On Error Resume Next
     m_oConn.Mode = 16  'adModeShareDenyNone
     m_oConn.Open
     If Err.Number <> 0 Then
       m_sLastError = "Microsoft.ACE.OLEDB.12.0: " & Err.Description
     End If
     On Error GoTo 0
     If (m_oConn.State And 1) = 1 Then
       ' Use OpenSchema.
       Set m_oRS = m_oConn.OpenSchema(20)   'adSchemaSheets
       While Not m_oRS.EOF
         If m_sSheets <> "" Then m_sSheets = m_sSheets & "¦"
         m_sSheets = m_sSheets & Replace(m_oRS!TABLE_NAME, "'", "")
         m_oRS.MoveNext
       Wend
       m_oRS.Close
       Set m_oRS = Nothing
       m_oConn.Close
       Set m_oConn = Nothing
     End If
   End If
   If m_sSheets <> "" Then
     m_oSheets.Add Replace(m_sSheets, "¦", vbTab)
     m_lSheetCount = m_oSheets.Count
   Else
     If m_sLastError = "" Then m_sLastError = "cBiff12Ado.GetSheets: Wrong or corrupt file!"
   End If
End Sub

Private Function SheetNameQuote(ByVal sName As String) As String
   SheetNameQuote = IIf(InStr(sName, " ") > 0, "'", "") & sName & IIf(InStr(sName, " ") > 0, "'", "")
End Function

Private Sub GetFields()
   Dim i As Integer, tmp() As String, tmp1() As String, tmp2() As String
   Dim aFields() As String
   Err.Clear
   m_sLastError = ""
   m_sFields = ""
   m_oFields.Clear
   m_lFieldCount = 0
   On Error Resume Next
   Set m_oConn = CreateObject("ADODB.Connection")
   If Err.Number <> 0 Then
     m_sLastError = "ADODB.Connection: " & Err.Description
   End If
   On Error GoTo 0
   If Not m_oConn Is Nothing Then
     'Connection für Datatype und Schema
     m_oConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source='" & m_sFilename & "';Extended Properties='Excel 12.0" & IIf(Not m_bHasHeader, ";HDR=NO", "") & "'; "
     On Error Resume Next
     m_oConn.Mode = 16  'adModeShareDenyNone
     m_oConn.Open
     If Err.Number <> 0 Then
       m_sLastError = "Microsoft.ACE.OLEDB.12.0: " & Err.Description
     End If
     On Error GoTo 0
     If (m_oConn.State And 1) = 1 Then
       ' Use OpenSchema.
       Set m_oRS = m_oConn.OpenSchema(4, Array(Empty, Empty, SheetNameQuote(m_sSheetName)))   'adSchemaColumns
       If Err.Number <> 0 Then
         m_sLastError = "Microsoft.ACE.OLEDB.12.0: " & Err.Description
       End If
     While Not m_oRS.EOF
'       For i = 0 To m_oRS.Fields.Count - 1
'         Debug.Print m_oRS(i).Name & " = " & m_oRS(i).Value
'       Next
       Select Case m_oRS!DATA_TYPE
         Case 9, 133, 134, 135
           If m_sFields <> "" Then m_sFields = m_sFields & "¦"
           m_sFields = m_sFields & Format(m_oRS!ORDINAL_POSITION, "000") & "," & m_oRS!COLUMN_NAME & "=DAT," & Format(m_oRS!DATA_TYPE, "00000")
         Case 11
           If m_sFields <> "" Then m_sFields = m_sFields & "¦"
           m_sFields = m_sFields & Format(m_oRS!ORDINAL_POSITION, "000") & "," & m_oRS!COLUMN_NAME & "=BOL," & Format(m_oRS!DATA_TYPE, "00000")
         Case 2, 3, 4, 5, 6, 12, 14, 16, 17, 18, 19, 20, 21, 131
           If m_sFields <> "" Then m_sFields = m_sFields & "¦"
           m_sFields = m_sFields & Format(m_oRS!ORDINAL_POSITION, "000") & "," & m_oRS!COLUMN_NAME & "=NUM," & Format(m_oRS!DATA_TYPE, "00000")
         Case 0, 9, 10, 13, 74, 72, 128, 132, 136, 138, 139, 200, 204, 205
           'Nicht ausgeben
         Case Else
           If m_sFields <> "" Then m_sFields = m_sFields & "¦"
           m_sFields = m_sFields & Format(m_oRS!ORDINAL_POSITION, "000") & "," & m_oRS!COLUMN_NAME & "=TXT," & Format(m_oRS!DATA_TYPE, "00000")
       End Select
     m_oRS.MoveNext
     Wend
     m_oRS.Close
     Set m_oRS = Nothing
     m_oConn.Close
     Set m_oConn = Nothing
     End If
   End If
   If m_sFields <> "" Then
     If InStr(m_sFields, "¦") > 0 Then
       aFields = Split(m_sFields, "¦")
     Else
       ReDim aFields(0) As String
       aFields(0) = m_sFields
     End If
     If UBound(aFields) > 0 Then
       'Sort Fields by Number
       ReDim tmp(UBound(aFields)) As String
       For i = 0 To UBound(aFields)
         tmp1() = Split(aFields(i), "=")
         tmp2() = Split(tmp1(0), ",")
         tmp(Val(tmp2(0)) - 1) = aFields(i)
       Next
       aFields = tmp
       m_sFields = Join(aFields, "¦")
       Erase aFields
       Erase tmp
       Erase tmp1
       Erase tmp2
     End If
     m_oFields.Add Replace(m_sFields, "¦", vbTab)
     m_lFieldCount = m_oFields.Count
   Else
     If m_sLastError = "" Then m_sLastError = "cBiff12RDAdo.GetFields: Sheet contains no fields."
   End If
End Sub

Private Sub GetRecords()
   Dim sRecordRow As String, i As Long
   Err.Clear
   m_sLastError = ""
   Set m_oRecords = New cBiff12RDRows
   On Error Resume Next
   Set m_oConn = CreateObject("ADODB.Connection")
   If Err.Number <> 0 Then
     m_sLastError = "ADODB.Connection: " & Err.Description
   Else
     If Err.LastDllError <> 0 Then
       m_sLastError = "ADODB.Connection: " & GetSystemErrorMessageText(Err.LastDllError)
     End If
   End If
   On Error GoTo 0
   If Not m_oConn Is Nothing Then
     'Connection für Datatype und Schema
     m_oConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source='" & m_sFilename & "';Extended Properties='Excel 12.0;IMEX=1" & IIf(Not m_bHasHeader, ";HDR=NO", "") & "'; "
     On Error Resume Next
     m_oConn.Mode = 16  'adModeShareDenyNone
     m_oConn.Open
     If Err.Number <> 0 Then
       m_sLastError = "Microsoft.ACE.OLEDB.12.0: " & Err.Description
     End If
     On Error GoTo 0
     If (m_oConn.State And 1) = 1 Then
       ' Use OpenSchema.
       Set m_oRS = m_oConn.Execute("SELECT * FROM [" & m_sSheetName & "]; ")
       While Not m_oRS.EOF
         sRecordRow = ""
         For i = 0 To m_oRS.Fields.Count - 1
           If Not IsNull(m_oRS(i).Value) Then
             Select Case m_oRS(i).Type
               Case 9, 133, 134, 135
                 If sRecordRow <> "" Then sRecordRow = sRecordRow & "¦"
                 sRecordRow = sRecordRow & Trim(Replace(Replace(CStr(m_oRS(i).Value), vbCrLf, vbLf), vbLf, vbCrLf))
               Case 11
                 If sRecordRow <> "" Then sRecordRow = sRecordRow & "¦"
                 sRecordRow = sRecordRow & Trim(Replace(Replace(CStr(m_oRS(i).Value), vbCrLf, vbLf), vbLf, vbCrLf))
               Case 2, 3, 4, 5, 6, 12, 14, 16, 17, 18, 19, 20, 21, 131
                 If sRecordRow <> "" Then sRecordRow = sRecordRow & "¦"
                 sRecordRow = sRecordRow & Trim(Replace(Replace(CStr(m_oRS(i).Value), vbCrLf, vbLf), vbLf, vbCrLf))
               Case 0, 9, 10, 13, 74, 72, 128, 132, 136, 138, 139, 200, 204, 205
                 'Nicht ausgeben
               Case Else
                 If sRecordRow <> "" Then sRecordRow = sRecordRow & "¦"
                 sRecordRow = sRecordRow & Trim(Replace(Replace(CStr(m_oRS(i).Value), vbCrLf, vbLf), vbLf, vbCrLf))
             End Select
           Else
             'Leeres Feld
             If sRecordRow <> "" Then sRecordRow = sRecordRow & "¦"
             sRecordRow = sRecordRow & "¤"
           End If
         Next
         sRecordRow = Replace(sRecordRow, "¤", "")       'Leere Inhalte Platzhalter entfernen
         sRecordRow = Replace(sRecordRow, vbTab, " ")       'Keine Tabulatoren im Text
         If (m_bNoEmptyRows And Trim(Replace(sRecordRow, "¦", "")) <> "") Or Not m_bNoEmptyRows Then
           m_oRecords.Add Replace(m_sFields, "¦", vbTab), Replace(sRecordRow, "¦", vbTab)
         End If
         m_oRS.MoveNext
       Wend
       m_oRS.Close
       Set m_oRS = Nothing
       m_oConn.Close
       Set m_oConn = Nothing
     End If
   End If
End Sub
