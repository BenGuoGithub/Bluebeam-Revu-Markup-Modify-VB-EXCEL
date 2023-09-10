Attribute VB_Name = "Module1"
Option Explicit
Dim sEnginePath As String, sPDF
Const APPNAME = "Change Bluebeam Revu Markups"

'Choose PDF
Sub Btn3_Click()
    Dim sTmp As String
    On Error Resume Next
    sTmp = ThisWorkbook.Path
    If InStr(1, sTmp, "http") >= 0 Then
        sTmp = Environ("OneDrive") & Mid(ActiveWorkbook.Path, Application.Find("@", Application.Substitute(ActiveWorkbook.Path, "/", "@", 4)), 999)
    End If
    ChDir sTmp
    sPDF = Application.GetOpenFilename("PDF File,*.pdf")
    If VarType(sPDF) = 11 Then
        MsgBox "Please select the PDF file.", vbOKOnly, APPNAME
        sPDF = ""
    End If
    Range("A3").Value = sPDF
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
    Dim I As Integer, aSubject() As String, aSubjectID() As String, iTmp As Integer, iID As Integer
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
    
    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(sEnginePath + " Open('" + sPDF + "') MarkupList(1) Close()")
    sResult = oExec.StdOut.ReadAll
    aID = Split(sResult, vbCrLf)
    sMsg = "Found: ID*" + Trim(UBound(aID)) + ";"
    
    iCount = 0
    sResult = ""
    sCMD = ""
    For I = 1 To UBound(aID)
        If aID(I) <> "" Then
            sCMD = sCMD + "MarkupGetEx(1, '" + aID(I) + "','subject') "
            
            iCount = iCount + 1
            If iCount = CountMax Then 'seperate the CMD in case of too long
                sCMD = " Open('" + sPDF + "') " + sCMD + "Close()"
                Set oExec = oShell.Exec(sEnginePath + sCMD)
                sResult = sResult + oExec.StdOut.ReadAll
                iCount = 0
                sCMD = ""
            End If
        End If
    Next I
    
    If iCount > 0 Then
        sCMD = " Open('" + sPDF + "') " + sCMD + "Close()"
        Set oExec = oShell.Exec(sEnginePath + sCMD)
        sResult = sResult + oExec.StdOut.ReadAll
        iCount = 0
    End If
    aR = Split(sResult, vbCrLf)
    
    iTmp = 0
    iID = 0
    For I = 0 To UBound(aR)
        Select Case aR(I)
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
                iStart = InStr(aR(I), "'subject':")
                iEnd = InStr(aR(I), "'}")
                aSubject(iTmp) = Mid(aR(I), iStart + 11, iEnd - iStart - 11)
                aSubjectID(iTmp) = aID(iID)
        End Select
    Next I
    sMsg = sMsg + " Markup*" + Trim(UBound(aSubject)) + ";"
    
    iTmp = 0
    iEnd = 0
    iCount = 0
    sCMD = ""
    For I = 1 To UBound(aSubject)
        If iTmp = 0 Then
            iTmp = iTmp + 1
            ReDim Preserve aSubjectSort(iTmp)
            aSubjectSort(iTmp) = aSubject(I)
        Else
            On Error Resume Next
            iStart = WorksheetFunction.Match(aSubject(I), aSubjectSort, 0)
            If Err <> 0 Then iStart = -1
            If iStart <= 0 Then
                iTmp = iTmp + 1
                ReDim Preserve aSubjectSort(iTmp)
                aSubjectSort(iTmp) = aSubject(I)
            End If
        End If
        
        On Error Resume Next
        sTmp = ""
        sTmp = WorksheetFunction.VLookup(aSubject(I), Range("A4:B100"), 2, 0)
        If sTmp <> "" Then
            iEnd = iEnd + 1
            sCMD = sCMD + "MarkupSet(1,'" + aSubjectID(I) + "',\""{'subject':'" + sTmp + "'}\"") "
            
            iCount = iCount + 1
            If iCount = CountMax Then 'seperate the CMD in case of too long
                If iEnd > iCount Then 'First time open from the origin PDF, the rest from the saved PDF
                    sCMD = " Open('" + sNewFile + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
                Else
                    sCMD = " Open('" + sPDF + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
                End If
                Set oExec = oShell.Exec(sEnginePath + sCMD)
                Sleep (2)
                iCount = 0
                sCMD = ""
            End If
        End If
    Next I
    sMsg = sMsg + " Paired*" + Trim(iEnd) + ";"
    Range("C2").Value = sMsg
    
    For I = 1 To UBound(aSubjectSort)
        Range("D" + Trim(3 + I)).Value = aSubjectSort(I)
        sTmp = ""
        sTmp = WorksheetFunction.VLookup(aSubjectSort(I), Range("A4:B100"), 2, 0)
        If sTmp = "" Then
            Range("D" + Trim(3 + I)).Select
            Selection.Interior.ThemeColor = xlThemeColorDark2
        End If
    Next I
    
    If iEnd > iCount Then
        sCMD = " Open('" + sNewFile + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
    Else
        sCMD = " Open('" + sPDF + "') " + sCMD + "Save('" + sNewFile + "',1) Close()"
    End If
    
    If iEnd > 0 Then
        Set oExec = oShell.Exec(sEnginePath + sCMD)
        MsgBox "PDF saved as: " + vbCrLf + sNewFile, vbOKCancel, APPNAME
    Else
        MsgBox "No paired markup, Please add the pair list in Col A&B.", vbOKCancel, APPNAME
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
