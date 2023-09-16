Attribute VB_Name = "Module1"
Option Explicit
Dim sEnginePath As String, sPDF
Const APPNAME = "Change Bluebeam Revu Markups"

'Choose PDF
Sub Btn3_Click()
    Dim sTmp As String
    On Error Resume Next
    If Range("A3").Value = "" Then
        sTmp = ThisWorkbook.Path
    Else
        sTmp = Range("A3").Value
        sTmp = Mid(sTmp, 1, InStrRev(sTmp, "\") - 1)
    End If
    
    If InStr(1, sTmp, "http") > 0 Then
        sTmp = Environ("OneDrive") & Mid(ActiveWorkbook.Path, Application.Find("@", Application.Substitute(ActiveWorkbook.Path, "/", "@", 4)), 999)
    End If
    ChDir sTmp
    sPDF = Application.GetOpenFilename("PDF File,*.pdf")
    If VarType(sPDF) = 11 Then
        MsgBox "Please select the PDF file.", vbOKOnly, APPNAME
        sPDF = ""
    Else
        Range("A3").Value = sPDF
    End If
End Sub

'Find ScriptEngine.exe
Sub Btn2_Click()
    On Error Resume Next
    ChDir Environ("ProgramFiles") + "\Bluebeam Software\Bluebeam Revu\20\Revu\"
    sEnginePath = Application.GetOpenFilename("ScriptEngine.exe,*.exe")
    If InStr(1, sEnginePath, "ScriptEngine.exe") < 1 Then
        MsgBox "Please select the ScriptEngine.exe.", vbOKOnly, APPNAME
        sEnginePath = ""
    End If
    Range("A2").Value = sEnginePath
End Sub

'Change Markups
Sub Btn1_Click()
    Dim oShell As Object, oExec As Object, oOutput As Object, sResult
    Dim aID() As String, aR() As String, ScriptEngine As String, sCMD As String
    Dim i As Integer, aSubject() As String, aSubjectID() As String, iTmp As Integer, iID As Integer
    Dim iStart As Integer, iEnd As Integer, aSubjectSort() As String, sTmp As String
    Dim sMsg As String, sNewFile As String, iCount As Integer
    Const CountMax = 100
    
    Range("C2").Value = ""
    Range("D4:D1000").Select
    Selection.ClearContents
    Selection.Interior.Pattern = xlNone
    
    If sEnginePath = "" Then
        If Range("A2").Value <> "" Then
            sEnginePath = Range("A2").Value
        Else
            MsgBox "Please select the ScriptEngine.exe first.", vbOKOnly, APPNAME
            Exit Sub
        End If
    End If
    
    If sPDF = "" Then
        If Range("A3").Value <> "" Then
            sPDF = Range("A3").Value
        Else
            MsgBox "Please select the PDF file.", vbOKOnly, APPNAME
            Exit Sub
        End If
    End If
    
    sNewFile = Mid(sPDF, 1, InStr(1, sPDF, ".pdf") - 1)
    sNewFile = sNewFile & "_" & DatePart("yyyy", Date) & DatePart("M", Date) & DatePart("D", Date) & ".pdf"
    
    Range("C2").Value = "Reading Markup ID list..."
    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(sEnginePath + " Open('" + sPDF + "') MarkupList(1) Close()")
    Sleep (0.5)
    sResult = oExec.StdOut.ReadAll
    aID = Split(sResult, vbCrLf)
    
    If IsEmpty(aID) Then
        MsgBox "Can't find any Markup ID in this file.", vbOKOnly, APPNAME
        Exit Sub
    End If
    sMsg = "Found: ID*" + Trim(UBound(aID)) + ";"
    Range("C2").Value = sMsg
    
    iCount = 0
    sResult = ""
    sCMD = ""
    Range("C2").Value = "Reading Markup subjects..."
    For i = 1 To UBound(aID)
        If aID(i) <> "" Then
            sCMD = sCMD + "MarkupGetEx(1, '" + aID(i) + "','subject') "
            
            iCount = iCount + 1
            If iCount = CountMax Then 'seperate the CMD in case of too long
                sCMD = " Open('" + sPDF + "') " + sCMD + "Close()"
                Range("C2").Value = "Reading Markup subjects to ID:" & Trim(i) & "/" & Trim(UBound(aID)) & "..."
                Set oExec = oShell.Exec(sEnginePath + sCMD)
                Sleep (0.5)
                sResult = sResult + oExec.StdOut.ReadAll
                iCount = 0
                sCMD = ""
            End If
        End If
    Next i
    
    If iCount > 0 Then
        sCMD = " Open('" + sPDF + "') " + sCMD + "Close()"
        Range("C2").Value = "Reading Markup subjects to ID:" & Trim(i) & "/" & Trim(UBound(aID)) & "..."
        Set oExec = oShell.Exec(sEnginePath + sCMD)
        Sleep (0.5)
        sResult = sResult + oExec.StdOut.ReadAll
        iCount = 0
    End If
    aR = Split(sResult, vbCrLf)
    
    iTmp = 0
    iID = 0
    For i = 0 To UBound(aR)
        Select Case aR(i)
            Case "0"
                'without subject
                iID = iID + 1
            Case "1"
                'with subject
            Case ""
            
            Case Else
                'subject
                iID = iID + 1
                iTmp = iTmp + 1
                ReDim Preserve aSubject(iTmp)
                ReDim Preserve aSubjectID(iTmp)
                iStart = InStr(aR(i), "'subject':")
                iEnd = InStr(aR(i), "'}")
                aSubject(iTmp) = Mid(aR(i), iStart + 11, iEnd - iStart - 11)
                aSubjectID(iTmp) = aID(iID)
        End Select
    Next i
    
    If IsEmpty(aSubject) Then
        MsgBox "Can't find any Markup Subject in this file.", vbOKOnly, APPNAME
        Exit Sub
    End If
    sMsg = sMsg + " Markup*" + Trim(UBound(aSubject)) + ";"
    Range("C2").Value = sMsg
    
    iTmp = 0
    iEnd = 0
    iCount = 0
    sCMD = ""
    Range("C2").Value = "Changing Markup subjects..."
    For i = 1 To UBound(aSubject)
        If iTmp = 0 Then
            iTmp = iTmp + 1
            ReDim Preserve aSubjectSort(iTmp)
            aSubjectSort(iTmp) = aSubject(i)
        Else
            On Error Resume Next
            iStart = WorksheetFunction.Match(aSubject(i), aSubjectSort, 0)
            If Err <> 0 Then iStart = -1
            If iStart <= 0 Then
                iTmp = iTmp + 1
                ReDim Preserve aSubjectSort(iTmp)
                aSubjectSort(iTmp) = aSubject(i)
            End If
        End If
        
        On Error Resume Next
        sTmp = ""
        sTmp = WorksheetFunction.VLookup(aSubject(i), Range("A4:B100"), 2, 0)
        If sTmp <> "" Then
            iEnd = iEnd + 1
            sCMD = sCMD + "MarkupSet(1,'" + aSubjectID(i) + "',\""{'subject':'" + sTmp + "'}\"") "
            
            iCount = iCount + 1
            If iCount = CountMax Then 'seperate the CMD in case of too long
                If iEnd > iCount Then 'First time open from the origin PDF, the rest from the saved PDF
                    sCMD = " Open('" + sNewFile + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
                Else
                    sCMD = " Open('" + sPDF + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
                End If
                Range("C2").Value = "Changing Markup subjects to ID:" & Trim(i) & "/" & Trim(UBound(aSubject)) & "..."
                Set oExec = oShell.Exec(sEnginePath + sCMD)
                Sleep (0.5)
                sResult = oExec.StdOut.ReadAll
                iCount = 0
                sCMD = ""
            End If
        End If
    Next i
    sMsg = sMsg + " Paired*" + Trim(iEnd) + ";"
    Range("C2").Value = sMsg
    
    For i = 1 To UBound(aSubjectSort)
        Range("D" + Trim(3 + i)).Value = aSubjectSort(i)
        sTmp = ""
        sTmp = WorksheetFunction.VLookup(aSubjectSort(i), Range("A4:B100"), 2, 0)
        If sTmp = "" Then
            sTmp = WorksheetFunction.VLookup(aSubjectSort(i), Range("B4:B100"), 1, 0)
            Range("D" + Trim(3 + i)).Select
            If sTmp = "" Then
                Selection.Interior.ThemeColor = xlThemeColorDark2
            Else
                Selection.Interior.ThemeColor = xlThemeColorAccent6
                Selection.Interior.TintAndShade = 0.799981688894314
            End If
        End If
    Next i
    
    If iEnd > iCount Then
        sCMD = " Open('" + sNewFile + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
    Else
        sCMD = " Open('" + sPDF + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
    End If
    
    If iEnd > 0 Then
        Set oExec = oShell.Exec(sEnginePath + sCMD)
        Sleep (0.5)
        sResult = oExec.StdOut.ReadAll
        MsgBox "PDF saved as: " + vbCrLf + sNewFile, vbOKOnly, APPNAME
    Else
        MsgBox "No paired markup, Please add the pair list in Col A&B.", vbOKOnly, APPNAME
    End If
    
    
    Set oShell = Nothing
    Set oExec = Nothing
    
End Sub

Sub Sleep(T As Single)  ' T:Seconds
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents
    Loop While Timer - time1 < T
End Sub

Function IsEmpty(ByVal sArray As Variant) As Boolean
    Dim i As Long
    IsEmpty = False
    On Error GoTo lerr:
    i = UBound(sArray)
    Exit Function
lerr:
        IsEmpty = True
End Function
