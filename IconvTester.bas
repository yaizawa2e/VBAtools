Attribute VB_Name = "IconvTester"
Sub iconvTest()
    Dim iconv As iconv
    Dim testdat() As Byte
    
    testdat = "a1‚ ‚¢‚¤"
    
    Set iconv = New iconv
    iconv.init "UTF-16LE", "Shift_JIS"
    hexDump iconv.iconv(testdat)
    iconv.init "UTF-16LE", "UTF-16LE"
    hexDump iconv.iconv(testdat)
    iconv.init "UTF-16LE", "UTF-16BE"
    hexDump iconv.iconv(testdat)
    iconv.init "UTF-16LE", "UTF-8"
    hexDump iconv.iconv(testdat)
    iconv.init "UTF-16LE", "UTF-8N"
    hexDump iconv.iconv(testdat)
    iconv.init "UTF-16LE", "UTF-7"
    hexDump iconv.iconv(testdat)
    
    iconv.init "UTF-16LE", "UTF-16LE"
    Debug.Print iconv.iconv(testdat)
End Sub

Sub hexDump(ByRef arg() As Byte)
    Dim i As Long
    Dim out As String
    
    For i = 0 To UBound(arg)
        If 0 < Len(out) Then out = out + " "
        If arg(i) < 16 Then out = out + "0"
        out = out + Hex(arg(i))
    Next i
    
    Debug.Print out
End Sub
