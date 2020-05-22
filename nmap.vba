' Copyright (c) 2020 Nathanael Wettstein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'
' -----------------------------------------------------------------------------
'
' Nmap for the poor, implemented in VBA
'
'
' Features:
'   * Scan an IP range for HTTP(S) responses
'   * Slow by design to avoid detection by IDS
'     (honestly, I just could make it work faster)
'
' Does not support (yet):
'   * Other protocols and ports
'   * IPv6
'   * CIDR notation
'   * Faster scanning
'
' Partly inspired by:
'   * https://github.com/andreafortuna-org/VBAIPFunctions
'
' -----------------------------------------------------------------------------
 
' Set case insensitive string search (globally)
Option Compare Text
 
'IP v4
Type Ipv4
    x(1 To 4) As Byte
End Type
 
' -----------------------------------------------------------------------------
' Conversion Tools for IP v4
' -----------------------------------------------------------------------------
 
' Text IP address (like "1.2.3.4") to Ipv4 type
Function StrToIpv4(ByVal ip As String) As Ipv4
    Dim pos As Integer
    Dim i As Integer
    Dim returnIp As Ipv4
    ip = ip + "."
    For i = 1 To 4
        pos = InStr(ip, ".")
        returnIp.x(i) = Val(Left(ip, pos - 1))
        ip = Mid(ip, pos + 1)
    Next
    StrToIpv4 = returnIp
End Function
 
'Ipv4 type to text IP
Function Ipv4ToStr(ip As Ipv4) As String
    Ipv4ToStr = CStr(ip.x(1)) + "." + _
                CStr(ip.x(2)) + "." + _
                CStr(ip.x(3)) + "." + _
                CStr(ip.x(4))
End Function
 
'Four bytes to Ipv4 type
Function BytesToIpv4(ByVal w As Byte, _
                     ByVal x As Byte, _
                     ByVal y As Byte, _
                     ByVal z As Byte) As Ipv4
    Dim returnIpv4 As Ipv4
    returnIpv4.x(1) = w
    returnIpv4.x(2) = x
    returnIpv4.x(3) = y
    returnIpv4.x(4) = z
    BytesToIpv4 = returnIpv4
End Function
 
' "Unit tests" to check at least "good" cases with
' well-formatted IPs
Sub runUnitTests()
    Dim allTestsOk As Boolean
    allTestsOk = True
   
    allTestsOk = allTestsOk And test_StrToIpv4()
    allTestsOk = allTestsOk And test_Ipv4ToStr()
    allTestsOk = allTestsOk And test_BytesToIpv4()
   
    If allTestsOk Then
        MsgBox "All tests ok :-)"
    Else
        MsgBox "Some tests failed :-("
    End If
End Sub
 
Function test_StrToIpv4()
    test_StrToIpv4 = True
    Dim ip As Ipv4
    ip = StrToIpv4("1.234.56.178")
    If Not (1 = ip.x(1) And _
            234 = ip.x(2) And _
            56 = ip.x(3) And _
            178 = ip.x(4)) Then
        MsgBox "StrToIpv4() does not work as expected"
        test_StrToIpv4 = False
    End If
End Function
 
Function test_Ipv4ToStr()
    test_Ipv4ToStr = True
    Dim ip As Ipv4
    ip = StrToIpv4("12.34.56.78")
    If Not ("12.34.56.78" = Ipv4ToStr(ip)) Then
        MsgBox "Ipv4ToStr() does not work as expected"
        test_Ipv4ToStr = False
    End If
End Function
 
Function test_BytesToIpv4()
    test_BytesToIpv4 = True
    Dim w As Integer, x As Integer, y As Integer, z As Integer
    w = 21
    x = 31
    y = 41
    z = 51
    Dim ip As Ipv4
    ip = BytesToIpv4(w, x, y, z)
    If Not (21 = ip.x(1) And _
            31 = ip.x(2) And _
            41 = ip.x(3) And _
            51 = ip.x(4)) Then
        MsgBox "BytesToIpv4() does not work as expected"
        test_BytesToIpv4 = False
    End If
End Function
 
' -----------------------------------------------------------------------------
' Scanning Logic
' -----------------------------------------------------------------------------
 
Function callURL(url As String, _
                 ByRef returnStatus As String, _
                 ByRef returnHtml As String) As String
    On Error GoTo EndNow
    Dim request As New MSXML2.ServerXMLHTTP60
   
    Dim timeout As Integer
    timeout = 1000
    request.setTimeouts timeout, timeout, timeout, timeout
   
    request.Open "GET", url, False
    request.send
    On Error GoTo OnTimeout
   
    returnStatus = request.statusText
    returnHtml = request.responseText
   
    Set request = Nothing
    callURL = ""
   
    Exit Function
OnTimeout:
    checkURL = "Timeout"
EndNow:
End Function
 
Function scanUrls(ByVal fromUrlStr As String, _
                  ByVal toUrlStr As String, _
                  ByVal row As Integer) As String
    Dim url As String, rc As String
    Dim html As String, status As String
    Dim w As Integer, x As Integer, y As Integer, z As Integer
    Dim numIps As Long
    Dim fromUrl As Ipv4, toUrl As Ipv4
   
    numIps = 0
    fromUrl = StrToIpv4(fromUrlStr)
    toUrl = StrToIpv4(toUrlStr)
 
    For w = fromUrl.x(1) To toUrl.x(1)
        For x = fromUrl.x(2) To toUrl.x(2)
            For y = fromUrl.x(3) To toUrl.x(3)
                For z = fromUrl.x(4) To toUrl.x(4)
                    url = "http://" + Ipv4ToStr(BytesToIpv4(w, x, y, z))
                    status = ""
                    html = ""
                    rc = callURL(url, status, html)
                    If (rc = "" And status <> "") Or 0 = z Then
                        'Cells(row, 1).Value = url
                        ActiveSheet.Hyperlinks.Add Cells(row, 1), _
                                                   Address:=url, _
                                                   TextToDisplay:=url
                        Cells(row, 2).Value = status
                        Cells(row, 3).Value = "Checked"
                        Cells(row, 4).Value = extractTitleFromHtml(html)
                        Cells(row, 3).Select
                        row = row + 1
                    End If
                    numIps = numIps + 1
                Next z
                ActiveWorkbook.Save
            Next y
        Next x
    Next w
    scanUrls = "Done processing " + CStr(numIps) + " IPs."
End Function
 
Function extractTitleFromHtml(html As String) As String
    Dim titleStart As Integer, titleEnd As Integer
    extractTitleFromHtml = ""
    If (html <> "") Then
        titleStart = InStr(html, "<title>")
        titleEnd = InStr(html, "</title>")
        If (titleStart <> 0) Then
            extractTitleFromHtml = RTrim(Replace(Mid(html, _
                titleStart + 7, titleEnd - titleStart - 7), vbLf, ""))
        Else
            extractTitleFromHtml = Left(rc, 60)
        End If
    End If
End Function
 
Sub scanUrlRange()
    Dim returnValue As String
    returnValue = scanUrls("10.0.0.0", "10.0.0.10", 2)
    MsgBox (returnValue)
End Sub
