Attribute VB_Name = "Module1"
Option Explicit
Dim sEnginePath As String, sPDF
Const APPNAME = "Change Bluebeam Revu Markups"

'Choose PDF
Sub Btn3_Click()
    Dim sTmp As String
    On Error Resume Next
    If ThisWorkbook.Worksheets(1).Range("A3").Value = "" Then
        sTmp = ThisWorkbook.Path
    Else
        sTmp = ThisWorkbook.Worksheets(1).Range("A3").Value
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
        ThisWorkbook.Worksheets(1).Range("A3").Value = sPDF
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
    ThisWorkbook.Worksheets(1).Range("A2").Value = sEnginePath
End Sub

'Change Markups
Sub Btn1_Click()
    Dim oShell As Object, oExec As Object, oOutput As Object, sResult
    Dim aID() As String, aR() As String, ScriptEngine As String, sCMD As String
    Dim I As Integer, aSubject() As String, aSubjectID() As String, iTmp As Integer, iID As Integer
    Dim iStart As Integer, iEnd As Integer, aSubjectSort() As String, sTmp As String
    Dim sMsg As String, sNewFile As String, iCount As Integer
    Const CountMax = 100
    
    ThisWorkbook.Worksheets(1).Range("C2").Value = ""
    ThisWorkbook.Worksheets(1).Range("D4:D1000").Select
    Selection.ClearContents
    Selection.Interior.Pattern = xlNone
    
    If sEnginePath = "" Then
        If ThisWorkbook.Worksheets(1).Range("A2").Value <> "" Then
            sEnginePath = ThisWorkbook.Worksheets(1).Range("A2").Value
        Else
            MsgBox "Please select the ScriptEngine.exe first.", vbOKOnly, APPNAME
            Exit Sub
        End If
    End If
    
    If sPDF = "" Then
        If ThisWorkbook.Worksheets(1).Range("A3").Value <> "" Then
            sPDF = ThisWorkbook.Worksheets(1).Range("A3").Value
        Else
            MsgBox "Please select the PDF file.", vbOKOnly, APPNAME
            Exit Sub
        End If
    End If
    
    sNewFile = Mid(sPDF, 1, InStr(1, sPDF, ".pdf") - 1)
    sNewFile = sNewFile & "_" & DatePart("yyyy", Date) & DatePart("M", Date) & DatePart("D", Date) & ".pdf"
    
    ThisWorkbook.Worksheets(1).Range("C2").Value = "Reading Markup ID list..."
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
    ThisWorkbook.Worksheets(1).Range("C2").Value = sMsg
    
    iCount = 0
    sResult = ""
    sCMD = ""
    ThisWorkbook.Worksheets(1).Range("C2").Value = "Reading Markup subjects..."
    For I = 1 To UBound(aID)
        If aID(I) <> "" Then
            sCMD = sCMD + "MarkupGetEx(1, '" + aID(I) + "','subject') "
            
            iCount = iCount + 1
            If iCount = CountMax Then 'seperate the CMD in case of too long
                sCMD = " Open('" + sPDF + "') " + sCMD + "Close()"
                ThisWorkbook.Worksheets(1).Range("C2").Value = "Reading Markup subjects to ID:" & Trim(I) & "/" & Trim(UBound(aID)) & "..."
                Set oExec = oShell.Exec(sEnginePath + sCMD)
                Sleep (0.5)
                sResult = sResult + oExec.StdOut.ReadAll
                iCount = 0
                sCMD = ""
            End If
        End If
    Next I
    
    If iCount > 0 Then
        sCMD = " Open('" + sPDF + "') " + sCMD + "Close()"
        ThisWorkbook.Worksheets(1).Range("C2").Value = "Reading Markup subjects to ID:" & Trim(I) & "/" & Trim(UBound(aID)) & "..."
        Set oExec = oShell.Exec(sEnginePath + sCMD)
        Sleep (0.5)
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
                aSubject(iTmp) = Trim(Mid(aR(I), iStart + 11, iEnd - iStart - 11))
                aSubjectID(iTmp) = aID(iID)
        End Select
    Next I
    
    If IsEmpty(aSubject) Then
        MsgBox "Can't find any Markup Subject in this file.", vbOKOnly, APPNAME
        Exit Sub
    End If
    sMsg = sMsg + " Markup*" + Trim(UBound(aSubject)) + ";"
    ThisWorkbook.Worksheets(1).Range("C2").Value = sMsg
    
    For I = 4 To 100
        sTmp = Trim(Range("A" & Trim(I)).Value)
        If sTmp <> "" And Range("A" & Trim(I)).Value <> sTmp Then
            Range("A" & Trim(I)).Value = sTmp
        End If
    Next I
    
    iTmp = 0
    iEnd = 0
    iCount = 0
    sCMD = ""
    ThisWorkbook.Worksheets(1).Range("C2").Value = "Changing Markup subjects..."
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
        sTmp = WorksheetFunction.VLookup(aSubject(I), ThisWorkbook.Worksheets(1).Range("A4:B100"), 2, 0)
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
                ThisWorkbook.Worksheets(1).Range("C2").Value = "Changing Markup subjects to ID:" & Trim(I) & "/" & Trim(UBound(aSubject)) & "..."
                Set oExec = oShell.Exec(sEnginePath + sCMD)
                Sleep (0.5)
                sResult = oExec.StdOut.ReadAll
                iCount = 0
                sCMD = ""
            End If
        End If
    Next I
    sMsg = sMsg + " Paired*" + Trim(iEnd) + ";"
    ThisWorkbook.Worksheets(1).Range("C2").Value = sMsg
    
    For I = 1 To UBound(aSubjectSort)
        ThisWorkbook.Worksheets(1).Range("D" + Trim(3 + I)).Value = aSubjectSort(I)
        sTmp = ""
        sTmp = WorksheetFunction.VLookup(aSubjectSort(I), ThisWorkbook.Worksheets(1).Range("A4:B100"), 2, 0)
        If sTmp = "" Then
            sTmp = WorksheetFunction.VLookup(aSubjectSort(I), ThisWorkbook.Worksheets(1).Range("B4:B100"), 1, 0)
            ThisWorkbook.Worksheets(1).Range("D" + Trim(3 + I)).Select
            If sTmp = "" Then
                Selection.Interior.ThemeColor = xlThemeColorDark2
            Else
                Selection.Interior.ThemeColor = xlThemeColorAccent6
                Selection.Interior.TintAndShade = 0.799981688894314
            End If
        End If
    Next I
    
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
    Dim I As Long
    IsEmpty = False
    On Error GoTo lerr:
    I = UBound(sArray)
    Exit Function
lerr:
        IsEmpty = True
End Function
