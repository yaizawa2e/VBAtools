VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Iconv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const defaultCode As String = "UTF-16LE"

Private from_code As String
Private to_code As String

Public Function iconv(ByRef indata() As Byte) As Byte()
    Dim reader As ADODB.Stream
    Dim writer As ADODB.Stream
    Dim bomcheck() As Byte
    Dim offset As Integer
    
    offset = 0
    
    ' Initialize Reader Stream
    Set reader = New ADODB.Stream
    reader.Open
    reader.Type = adTypeText
    reader.Charset = normalizeCodeName(from_code)
    
    ' Set Input Data to Reader Stream
    reader.Type = adTypeBinary
    reader.Write indata
    
    ' Rewind Reader Stream and Prepare to Read
    reader.Position = 0
    reader.Type = adTypeText
    
    ' Initialize Writer Stream
    Set writer = New ADODB.Stream
    writer.Open
    writer.Type = adTypeText
    writer.Charset = normalizeCodeName(to_code)
    
    ' Copy Input String to Writer Stream
    writer.WriteText reader.ReadText
    
    ' Check Byte Order Mark
    writer.Position = 0
    writer.Type = adTypeBinary
    Select Case LCase(to_code)
        Case "utf-8n"
            bomcheck = writer.Read(3)
            If bomcheck(0) = 239 And bomcheck(1) = 187 And bomcheck(2) = 191 Then offset = 3
        Case "utf-16be"
            bomcheck = writer.Read(2)
            If bomcheck(0) = 254 And bomcheck(1) = 255 Then offset = 2
        Case "utf-16le"
            bomcheck = writer.Read(2)
            If bomcheck(0) = 255 And bomcheck(1) = 254 Then offset = 2
    End Select
    
    ' Rewind Writer Stream and Read Binary
    writer.Position = offset
    iconv = writer.Read
End Function

Public Sub init(Optional ByRef fromcode As String = defaultCode, Optional ByRef tocode As String = defaultCode)
    from_code = fromcode
    to_code = tocode
End Sub

Private Function normalizeCodeName(Optional ByRef codename As String = defaultCode)
    Select Case LCase(codename)
        Case "cp932"
            normalizeCodeName = "Shift_JIS"
        Case "ms932"
            normalizeCodeName = "Shift_JIS"
        Case "utf-8n"
            normalizeCodeName = "UTF-8"
        Case Else
            normalizeCodeName = codename
    End Select
End Function

Private Sub Class_Initialize()
    '
End Sub
