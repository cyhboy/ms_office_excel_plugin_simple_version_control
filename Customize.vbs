Option Explicit

Public Sub SynMZ()
    If testing Then Exit Sub
    
    Dim fso As Object
    Dim fileObject, cFileObject, mFileObject, obj As Object
    
    Dim cPath, mPath As String
    Dim iRet As Integer
    
    Dim activeName As String
    activeName = ActiveWorkbook.FullName
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)
    'MsgBox fso.Drives.count
    If fso.Drives.Count > 1 Then
        For Each obj In fso.Drives
            On Error GoTo ErrorHandler
            If obj.DriveType = 3 Then
                If obj.path = theDrive Or obj.path = "Z:" Then
                    If InStr(activeName, "C:") > 0 Then
                        mPath = Replace(activeName, "C:", obj.path)
                        
                        If fso.FileExists(mPath) Then
                            
                            Set mFileObject = fso.GetFile(mPath)
                            
                            If fileObject.DateLastModified > mFileObject.DateLastModified Then
                                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                                    iRet = MsgBox("You are NOT the author of this document, Do U want to manually update " & obj.path & " as well? ", vbYesNo, "Question")
                                    If iRet = vbYes Then
                                        fso.copyfile activeName, mPath, True
                                    End If
                                End If
                            End If
                        Else
                            If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                                iRet = MsgBox("You are the author of this document, Do U want to manually append " & obj.path & " as well? ", vbYesNo, "Question")
                                If iRet = vbYes Then
                                    fso.copyfile activeName, mPath, True
                                End If
                            End If
                        End If
                    ElseIf InStr(activeName, obj.path) > 0 Then
                        cPath = Replace(activeName, obj.path, "C:")
                        
                        If fso.FileExists(cPath) Then
                            Set cFileObject = fso.GetFile(cPath)
                            
                            If fileObject.DateLastModified > cFileObject.DateLastModified Then
                                
                                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                                    iRet = MsgBox("You are the author of this document, Do U want to manually update C: as well? ", vbYesNo, "Question")
                                    If iRet = vbYes Then
                                        fso.copyfile activeName, cPath, True
                                    End If
                                End If
                            End If
                        Else
                            If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                                iRet = MsgBox("You are NOT the author of this document, Do U want to manually append C: as well? ", vbYesNo, "Question")
                                If iRet = vbYes Then
                                    fso.copyfile activeName, cPath, True
                                End If
                            End If
                        End If
                        
                    End If
                End If
            End If
            
ErrorHandler:
            If Err.Number <> 0 Then
                MyMsgBox Err.Number & " " & Err.Description, 30
            End If
        Next
    Else
        If InStr(activeName, "C:") > 0 Then
            mPath = Replace(activeName, "C:", "\\192.168.0.73\tmp")
            If fso.FileExists(mPath) Then
                Set mFileObject = fso.GetFile(mPath)
                If fileObject.DateLastModified > mFileObject.DateLastModified Then
                    If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                        MyQuestionBox "You are NOT the author of this document, Do U want to manually update \\192.168.0.73\tmp as well? ", "No", "Yes", 10
                        If confirmation = "Yes" Then
                            fso.copyfile activeName, mPath, True
                        End If
                    End If
                End If
            Else
                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                    MyQuestionBox "You are the author of this document, Do U want to manually append \\192.168.0.73\tmp as well? ", "Yes", "No", 10
                    If confirmation = "Yes" Then
                        fso.copyfile activeName, mPath, True
                    End If
                End If
            End If
        End If
    End If
    Set fso = Nothing
End Sub

Public Sub CommonCheckSynWithAllAvailableDrive()
    If testing Then Exit Sub
    
    If InStr(ActiveWorkbook.Sheets("Info").Range("A1"), "#") > 0 Then Exit Sub
    Dim fso As Object
    Dim fileObject, cFileObject, mFileObject, obj As Object
    Dim cPath, mPath As String
    
    Dim activeName As String
    activeName = ActiveWorkbook.FullName
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)
    
    If fso.Drives.Count > 2 Then
        For Each obj In fso.Drives
            On Error GoTo ErrorHandler
            If obj.DriveType = 3 Then
                If InStr(activeName, "C:") > 0 Then
                    mPath = Replace(activeName, "C:", obj.path)
                    'If Dir(mPath) <> "" Then
                    If fso.FileExists(mPath) Then
                        Set mFileObject = fso.GetFile(mPath)
                        If fileObject.DateLastModified > mFileObject.DateLastModified Then
                            
                            If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                                MyQuestionBox "You are the author of this document, Do U want to update " & obj.path & " as well? ", "Yes", "No", 10
                                If confirmation = "Yes" Then
                                    fso.copyfile activeName, mPath, True
                                End If
                            End If
                        End If
                        
                    End If
                ElseIf InStr(activeName, obj.path) > 0 Then
                    cPath = Replace(activeName, obj.path, "C:")
                    'If Dir(cPath) <> "" Then
                    If fso.FileExists(cPath) Then
                        Set cFileObject = fso.GetFile(cPath)
                        If fileObject.DateLastModified > cFileObject.DateLastModified Then
                            'If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) >= 0 Then
                            MyQuestionBox "You are not the author of this document, Do U want to update C: as well? ", "Yes", "No", 10
                            If confirmation = "Yes" Then
                                fso.copyfile activeName, cPath, True
                            End If
                            'End If
                        End If
                    End If
                End If
            End If
ErrorHandler:
            If Err.Number <> 0 Then
                MyMsgBox Err.Number & " " & Err.Description, 30
            End If
        Next
    Else
        If Len(theDrive) = 2 Then
            If InStr(activeName, "C:") > 0 Then
                mPath = Replace(activeName, "C:", theDrive)
                'If Dir(mPath) <> "" Then
                If fso.FileExists(mPath) Then
                    Set mFileObject = fso.GetFile(mPath)
                    If fileObject.DateLastModified > mFileObject.DateLastModified Then
                        If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                            MyQuestionBox "You are the author of this document, Do U want to update " & theDrive & " as well? ", "Yes", "No", 10
                            If confirmation = "Yes" Then
                                fso.copyfile activeName, mPath, True
                            End If
                        End If
                    End If
                    
                End If
            End If
        End If
    End If
    
    Set fso = Nothing
End Sub

Public Sub ShellRunHide(cmd As String)
    If testing Then Exit Sub
    'On Error GoTo ErrorHandler
    Shell cmd, vbHide
    'ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description, 30
    '    End If
End Sub

Public Sub TestVBA()
    testing = True
    On Error GoTo ErrorHandler
    Dim objProject As VBIDE.VBProject
    Dim objComponent As VBIDE.VBComponent
    Dim objCode As VBIDE.CodeModule
    
    ' Declare other miscellaneous variables.
    Dim iLine As Integer
    Dim sProcName As String
    Dim pk As VBIDE.vbext_ProcKind
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Dim i As Integer
    Dim comm As String
    Dim codeOfLine As String
    
    Set objProject = ThisWorkbook.VBProject
    Dim subCount0 As Integer
    Dim subCount1 As Integer
    Dim subCount2 As Integer
    Dim subCount3 As Integer
    Dim subCount4 As Integer
    Dim subCount5 As Integer
    Dim subCount6 As Integer
    Dim subCount7 As Integer
    Dim subCount8 As Integer
    Dim subCountX As Integer
    
    Dim funcCount0 As Integer
    Dim funcCount1 As Integer
    Dim funcCount2 As Integer
    Dim funcCount3 As Integer
    Dim funcCount4 As Integer
    Dim funcCount5 As Integer
    Dim funcCount6 As Integer
    Dim funcCount7 As Integer
    Dim funcCount8 As Integer
    Dim funcCountX As Integer
    
    Dim xObj As Variant
    'Iterate through each component in the project.
    For Each objComponent In objProject.VBComponents
        'If InStr(objComponent.Name, "All") > 0 Or InStr(objComponent.Name, "SubParam") > 0 Or InStr(objComponent.Name, "FuncNoParam") > 0 Or InStr(objComponent.Name, "FuncParam") Then
        'Find the code module for the project.
        Set objCode = objComponent.CodeModule
        'Scan through the code module, looking for procedures.
        iLine = 1
        Do While iLine < objCode.CountOfLines
            
            codeOfLine = objCode.Lines(iLine, 1)
            If Trim(codeOfLine) <> "" And Not StartsWith(Trim(codeOfLine), "'") Then
                sProcName = objCode.ProcOfLine(iLine, pk)
                If sProcName <> "" And sProcName <> "Ver" And sProcName <> "Test" And sProcName <> "TestVBA" And sProcName <> "CountRegx" And sProcName <> "ListNodes" And sProcName <> "CntOfficeUI" And sProcName <> "TestCall" And sProcName <> "StartsWith" And sProcName <> "EndsWith" Then
                    comm = ""
                    If testing Then
                        If InStr(Trim(codeOfLine), "Public Sub " & sProcName & "()") > 0 Then
                            'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
                            'RobotRunByParam objComponent.Name & "." & sProcName
                            comm = "'Svc.xlam'!" & objComponent.Name & "." & sProcName
                            Application.Run comm
                            subCount0 = subCount0 + 1
                        Else
                            If InStr(Trim(codeOfLine), "Public Sub " & sProcName) > 0 Then
                                'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
                                comm = "'Svc.xlam'!" & objComponent.Name & "." & sProcName
                                If CountRegx(Trim(codeOfLine), ", ") = 0 Then
                                    Application.Run comm, "0"
                                    subCount1 = subCount1 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 1 Then
                                    Application.Run comm, "0", "0"
                                    subCount2 = subCount2 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 2 Then
                                    Application.Run comm, "0", "0", "0"
                                    subCount3 = subCount3 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 3 Then
                                    Application.Run comm, "0", "0", "0", "0"
                                    subCount4 = subCount4 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 4 Then
                                    Application.Run comm, "0", "0", "0", "0", "0"
                                    subCount5 = subCount5 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 5 Then
                                    Application.Run comm, "0", "0", "0", "0", "0", "0"
                                    subCount6 = subCount6 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 6 Then
                                    Application.Run comm, "0", "0", "0", "0", "0", "0", "0"
                                    subCount7 = subCount7 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 7 Then
                                    Application.Run comm, "0", "0", "0", "0", "0", "0", "0", "0"
                                    subCount8 = subCount8 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") >= 8 Then
                                    'Application.Run comm, "0", "0", "0", "0", "0", "0", "0", "0", "0"
                                    MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine) & " Not Test As Too Many Param"
                                    subCountX = subCountX + 1
                                End If
                            End If
                        End If
                        
                        If InStr(Trim(codeOfLine), "Public Function " & sProcName & "()") > 0 Then
                            'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
                            'RobotRunByParam objComponent.Name & "." & sProcName
                            comm = "'Svc.xlam'!" & objComponent.Name & "." & sProcName
                            Application.Run comm
                            funcCount0 = funcCount0 + 1
                        Else
                            If InStr(Trim(codeOfLine), "Public Function " & sProcName) > 0 Then
                                'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
                                comm = "'Svc.xlam'!" & objComponent.Name & "." & sProcName
                                If CountRegx(Trim(codeOfLine), ", ") = 0 Then
                                    Application.Run comm, xObj
                                    funcCount1 = funcCount1 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 1 Then
                                    Application.Run comm, xObj, xObj
                                    funcCount2 = funcCount2 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 2 Then
                                    Application.Run comm, xObj, xObj, xObj
                                    funcCount3 = funcCount3 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 3 Then
                                    Application.Run comm, xObj, xObj, xObj, xObj
                                    funcCount4 = funcCount4 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 4 Then
                                    Application.Run comm, xObj, xObj, xObj, xObj, xObj
                                    funcCount5 = funcCount5 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 5 Then
                                    Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj
                                    funcCount6 = funcCount6 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 6 Then
                                    Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj, xObj
                                    funcCount7 = funcCount7 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") = 7 Then
                                    Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj, xObj, xObj
                                    funcCount8 = funcCount8 + 1
                                End If
                                If CountRegx(Trim(codeOfLine), ", ") >= 8 Then
                                    'Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj, xObj, xObj, xObj
                                    MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine) & " Not Test As Too Many Param"
                                    funcCountX = funcCountX + 1
                                End If
                            End If
                        End If
                        
                        'Exit Sub
                    End If
                    
                End If
                'iLine = iLine + objCode.ProcCountLines(sProcName, pk) - 2
            End If
            iLine = iLine + 1
        Loop
        Set objCode = Nothing
        Set objComponent = Nothing
        'End If
    Next
    Set objProject = Nothing
    Dim resultStr As String
    resultStr = resultStr & "Total Sub0 Testing Count " & subCount0 & vbCrLf
    resultStr = resultStr & "Total Sub1 Testing Count " & subCount1 & vbCrLf
    resultStr = resultStr & "Total Sub2 Testing Count " & subCount2 & vbCrLf
    resultStr = resultStr & "Total Sub3 Testing Count " & subCount3 & vbCrLf
    resultStr = resultStr & "Total Sub4 Testing Count " & subCount4 & vbCrLf
    resultStr = resultStr & "Total Sub5 Testing Count " & subCount5 & vbCrLf
    resultStr = resultStr & "Total Sub6 Testing Count " & subCount6 & vbCrLf
    resultStr = resultStr & "Total Sub7 Testing Count " & subCount7 & vbCrLf
    resultStr = resultStr & "Total Sub8 Testing Count " & subCount8 & vbCrLf
    resultStr = resultStr & "Total SubX Not Testing Count " & subCountX & vbCrLf
    
    resultStr = resultStr & "Total Sub Count " & (subCount0 + subCount1 + subCount2 + subCount3 + subCount4 + subCount5 + subCount6 + subCount7 + subCount8 + subCountX) & vbCrLf & vbCrLf
    
    resultStr = resultStr & "Total Func0 Testing Count " & funcCount0 & vbCrLf
    resultStr = resultStr & "Total Func1 Testing Count " & funcCount1 & vbCrLf
    resultStr = resultStr & "Total Func2 Testing Count " & funcCount2 & vbCrLf
    resultStr = resultStr & "Total Func3 Testing Count " & funcCount3 & vbCrLf
    resultStr = resultStr & "Total Func4 Testing Count " & funcCount4 & vbCrLf
    resultStr = resultStr & "Total Func5 Testing Count " & funcCount5 & vbCrLf
    resultStr = resultStr & "Total Func6 Testing Count " & funcCount6 & vbCrLf
    resultStr = resultStr & "Total Func7 Testing Count " & funcCount7 & vbCrLf
    resultStr = resultStr & "Total Func8 Testing Count " & funcCount8 & vbCrLf
    resultStr = resultStr & "Total FuncX Not Testing Count " & funcCountX & vbCrLf
    
    resultStr = resultStr & "Total Func Count " & (funcCount0 + funcCount1 + funcCount2 + funcCount3 + funcCount4 + funcCount5 + funcCount6 + funcCount7 + funcCount8 + funcCountX) & vbCrLf & vbCrLf
    
    MsgBox resultStr
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description & " " & objComponent.Name & "." & sProcName, 30
    End If
    testing = False
End Sub

Public Sub MyQuestionBox(detail As String, answer1 As String, answer2 As String, duration As Long)
    If testing Then Exit Sub
    nexttime = Now() + TimeSerial(0, 0, duration)
    Application.OnTime nexttime, "MyQuestionBoxHide"
    confirmation = ""
    'UserForm2.CommandButton1.Caption = answer1
    'UserForm2.CommandButton2.Caption = answer2
    'UserForm2.TextBox1.text = detail
    'UserForm2.TextBox1.SetFocus
    'UserForm2.Show
    
    Set uf2 = New UserForm2
    uf2.CommandButton1.Caption = answer1
    uf2.CommandButton2.Caption = answer2
    uf2.TextBox1.text = detail
    uf2.TextBox1.SetFocus
    uf2.Show
End Sub

Public Sub MyMsgBox(detail As String, duration As Long)
    If testing Then Exit Sub
    nexttime = Now() + TimeSerial(0, 0, duration)
    Application.OnTime nexttime, "MyMsgBoxHide"
    
    'UserForm1.TextBox1.text = detail
    'UserForm1.TextBox1.SetFocus
    'UserForm1.Show
    Set uf1 = New UserForm1
    uf1.TextBox1.text = detail
    uf1.TextBox1.SetFocus
    uf1.Show
End Sub

Public Sub CpFil2FilBk(filePath1 As String, filePath2 As String, displayFlag As Boolean)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim result
    result = fso.copyfile(filePath1, filePath2)
    Set fso = Nothing
    
    If displayFlag Then
        If result = "" Then
            MyMsgBox filePath1 & " to " & filePath2 & " copied", 5
        End If
    End If
    Application.Workbooks.Open filePath2

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

Public Sub CpFil2Fil(filePath1 As String, filePath2 As String, displayFlag As Boolean)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.copyfile filePath1, filePath2
    Set fso = Nothing
    If displayFlag Then
        MyMsgBox filePath1 & " to " & filePath2 & " copied", 5
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

