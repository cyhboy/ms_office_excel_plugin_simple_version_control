Option Explicit

Public Function EndsWith(str As String, ending As String) As Boolean
    'If testing Then Exit Function
    Dim endingLen As Integer
    endingLen = Len(ending)
    EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Public Function StartsWith(str As String, start As String) As Boolean
    'If testing Then Exit Function
    Dim startLen As Integer
    startLen = Len(start)
    StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

Public Function CommonCheckSynWithM_OPEN()
    If testing Then Exit Function
    Dim closeFlag As Boolean
    closeFlag = False
    
    Dim activeName As String
    activeName = ActiveWorkbook.FullName
    
    If InStr(ActiveWorkbook.Sheets("Info").Range("A1"), "#") > 0 Then
        CommonCheckSynWithM_OPEN = closeFlag
        Exit Function
    End If
    
    If InStr(activeName, "http") = 1 Then
        CommonCheckSynWithM_OPEN = closeFlag
        Exit Function
    End If
    
    Dim fso As Object
    'Dim md As Object
    Dim fileObject, cFileObject, mFileObject As Object
    
    Dim cPath, mPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)
    'Set md = fso.GetDrive(theDrive)
    
    Dim path As String
    Dim parameter As String
    'MsgBox Len(theDrive)
    'MsgBox theUser
    If Len(theDrive) = 2 Then
        If InStr(activeName, "C:") > 0 Then
            mPath = Replace(activeName, "C:", theDrive)
            'If md.IsReady Then
            'If Dir(mPath) <> "" Then
            If fso.FileExists(mPath) Then
                Set mFileObject = fso.GetFile(mPath)
                
                If fileObject.DateLastModified < mFileObject.DateLastModified Then
                    If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                        MyQuestionBox "You are not the author of this document and found another updated verion in share drive, Do U want to override from " & theDrive & " after close? Last modifier is " & GetWorkbookProperties(mPath, "Last Author"), "Yes", "No", 10
                        If confirmation = "Yes" Then
                            nexttime = Now() + TimeSerial(0, 0, 5)
                            Application.OnTime nexttime, "'CpFil2FilBk """ & mPath & """, """ & activeName & """, True'"
                            
                            closeFlag = True
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            Else
                closeFlag = False
            End If
            
            'End If
        End If
        
        If InStr(activeName, theDrive) > 0 Then
            cPath = Replace(activeName, theDrive, "C:")
            
            'If Dir(cPath) <> "" Then
            If fso.FileExists(cPath) Then
                Set cFileObject = fso.GetFile(cPath)
                
                If fileObject.DateLastModified < cFileObject.DateLastModified Then
                    If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                        MyQuestionBox "You are not the author of this document and found another updated verion in C drive, Do U want to override from C: after close? Last modifier is " & GetWorkbookProperties(cPath, "Last Author"), "Yes", "No", 10
                        If confirmation = "Yes" Then
                            nexttime = Now() + TimeSerial(0, 0, 5)
                            Application.OnTime nexttime, "'CpFil2FilBk """ & cPath & """, """ & activeName & """, True'"
                            
                            closeFlag = True
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            Else
                MyQuestionBox "You didn't update a local copy of this document yet, Do U want to proceed now? ", "Yes", "No", 10
                If confirmation = "Yes" Then
                    fso.copyfile activeName, cPath, True
                End If
                
                closeFlag = False
            End If
            
        End If
    End If
    Set fso = Nothing
    
    CommonCheckSynWithM_OPEN = closeFlag
End Function

Public Function CommonCheckSynWithAllAvailableDrive_OPEN()
    If testing Then Exit Function
    Dim closeFlag As Boolean
    closeFlag = False
    
    Dim activeName As String
    activeName = ActiveWorkbook.FullName
    
    If InStr(ActiveWorkbook.Sheets("Info").Range("A1"), "#") > 0 Then
        CommonCheckSynWithAllAvailableDrive_OPEN = closeFlag
        Exit Function
    End If
    
    If InStr(activeName, "http") = 1 Then
        CommonCheckSynWithAllAvailableDrive_OPEN = closeFlag
        Exit Function
    End If
    
    Dim fso As Object
    Dim fileObject, cFileObject, mFileObject As Object
    
    Dim cPath, mPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)
    
    Dim path As String
    Dim parameter As String
    
    Dim obj As Object
    For Each obj In fso.Drives
        If obj.DriveType = 3 Then
            If InStr(activeName, "C:") > 0 Then
                mPath = Replace(activeName, "C:", obj.path)
                
                'If Dir(mPath) <> "" Then
                If fso.FileExists(mPath) Then
                    Set mFileObject = fso.GetFile(mPath)
                    
                    If fileObject.DateLastModified < mFileObject.DateLastModified Then
                        If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                            MyQuestionBox "You are not the author of this document and found another updated verion in share drive, Do U want to override from " & obj.path & " after close? Last modifier is " & GetWorkbookProperties(mPath, "Last Author"), "Yes", "No", 10
                            If confirmation = "Yes" Then
                                nexttime = Now() + TimeSerial(0, 0, 5)
                                Application.OnTime nexttime, "'CpFil2FilBk """ & mPath & """, """ & activeName & """, True'"
                                closeFlag = True
                                Exit For
                            Else
                                closeFlag = False
                                Exit For
                            End If
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            ElseIf InStr(activeName, obj.path) > 0 Then
                cPath = Replace(activeName, obj.path, "C:")
                
                'If Dir(cPath) <> "" Then
                If fso.FileExists(cPath) Then
                    Set cFileObject = fso.GetFile(cPath)
                    
                    If fileObject.DateLastModified < cFileObject.DateLastModified Then
                        If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) < 0 Then
                            MyQuestionBox "You are not the author of this document and found another updated verion in C drive, Do U want to override from C: after close? Last modifier is " & GetWorkbookProperties(cPath, "Last Author"), "Yes", "No", 10
                            If confirmation = "Yes" Then
                                nexttime = Now() + TimeSerial(0, 0, 5)
                                Application.OnTime nexttime, "'CpFil2FilBk """ & cPath & """, """ & activeName & """, True'"
                                closeFlag = True
                            Else
                                closeFlag = False
                            End If
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
                
            End If
        End If
        
    Next
    
    Set fso = Nothing
    CommonCheckSynWithAllAvailableDrive_OPEN = closeFlag
End Function

Public Function CommonGetTheDrive()
    If testing Then Exit Function
    Dim fso As Object
    Dim obj As Object
    Dim retDrive As String
    retDrive = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each obj In fso.Drives
        If obj.DriveType = 3 Then
            'MsgBox "Testing Drive: " & obj.path
            'If Dir(obj.path & "\AppFiles\SupportSetup\Svc.xlam") <> "" Then
            If fso.FileExists(obj.path & "\AppFiles\SupportSetup\Svc.xlam") Then
                retDrive = obj.path
                If obj.path = "M:" Then
                    Exit For
                End If
            End If
        End If
    Next
    Set fso = Nothing
    
    If retDrive = "" Then
        retDrive = "\\192.168.0.73\tmp"
    End If
    
    CommonGetTheDrive = retDrive
End Function

Public Function GetWorkbookProperties(ByVal filePath As String, ByVal propName As String)
    If testing Then Exit Function
    Dim retvalue As String
    Dim appOffice As New Application
    Dim richFile As Workbook
    Set richFile = appOffice.Workbooks.Open(filePath)
    retvalue = richFile.BuiltinDocumentProperties(propName)
    richFile.Saved = True
    'richFile.Close
    appOffice.Workbooks.Close
    appOffice.Quit
    Set appOffice = Nothing
    GetWorkbookProperties = retvalue
End Function

Public Function CountRegx(text As String, patt As String) As Long
    On Error GoTo ErrorHandler
    Dim RE As New RegExp
    RE.Pattern = patt
    RE.Global = True
    RE.IgnoreCase = False
    RE.MultiLine = True
    'Retrieve all matches
    Dim Matches As MatchCollection
    Set Matches = RE.Execute(text)
    'Return the corrected count of matches
    CountRegx = Matches.Count
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Function

