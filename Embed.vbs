Option Explicit

Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim cFileObject As Object
    Dim mFileObject As Object
    Dim obj As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim thePath As String
    Dim thePathUI As String
    Dim thePathVer As String
    
    Dim theFolder As String
    
    Dim scriptPath As String
    Dim scriptParameter As String
    
    Dim uiPath As String
    uiPath = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    
    Dim macroPath As String
    macroPath = ThisWorkbook.FullName
    
    theDrive = CommonGetTheDrive()
    'MsgBox theDrive
    'theUser = RespExtMail(Environ$("username"), "EXTERNAL_MAIL")
    theUser = Environ$("username")
    'MsgBox theUser
    'MsgBox ReadEnv("%PROGRAMFILES%")
    'MsgBox Environ("AppData")
    'MsgBox Environ("USERPROFILE")
    'MsgBox ThisWorkbook.FullName
    Dim cFile As String
    cFile = "C:\AppFiles\Svc.xlam"
    Set cFileObject = fso.GetFile(cFile)
    
    Dim mFileDate As Date
    Dim cFileDate As Date
    
    mFileDate = DateAdd("yyyy", -5, Now)
    cFileDate = DateAdd("yyyy", -5, Now)
    
    cFileDate = cFileObject.DateLastModified
    
    For Each obj In fso.Drives()
        'MsgBox obj.path & " " & obj.DriveType
        If obj.DriveType = 3 Then
            'If Dir(obj.path & "\AppFiles\SupportSetup\Svc.xlam") <> "" Then
            If fso.FileExists(obj.path & "\AppFiles\SupportSetup\Svc.xlam") Then
                Set mFileObject = fso.GetFile(obj.path & "\AppFiles\SupportSetup\Svc.xlam")
                If mFileObject.DateLastModified > mFileDate Then
                    mFileDate = mFileObject.DateLastModified
                    theFolder = mFileObject.ParentFolder
                    thePath = obj.path & "\AppFiles\SupportSetup\Svc.xlam"
                    thePathUI = obj.path & "\AppFiles\SupportSetup\Excel.officeUI"
                    thePathVer = obj.path & "\AppFiles\SupportSetup\" & "Excel_" & Environ$("username") & ".officeUI"
                End If
            End If
        End If
    Next
    
    Set mFileObject = Nothing
    Set cFileObject = Nothing
    Set fso = Nothing
    
    If (mFileDate - cFileDate > 0) Then
        MyMsgBox "Dear CST Users, Thanks for choosing common support toolkits for your daily work. You now was recommended to upgrade to a new CST version, Please free 1 min to close your office suites and double click " & theFolder & "\install.bat. Thanks very much in deep. ", 10
        'ShellRun "explorer.exe " & theFolder
        CpFil2Fil thePathUI, uiPath, False
        CpFil2Fil uiPath, thePathVer, False
        CpFil2Fil thePath, cFile, False
        
        scriptPath = "WScript.exe C:\AppFiles\WaitThenRunHiddenJob.vbs "
        scriptParameter = """cmd.exe /C copy /Y %22" & thePath & "%22" & " " & "%22" & macroPath & "%22""" & " " & """5000"""
        ShellRunHide scriptPath & scriptParameter
        'Sleep 1000
        ThisWorkbook.Saved = True
        ThisWorkbook.Close
    End If
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox "Dear CST Users, When you see this message, The initialization of Common Support Toolkits may encounter some abnormal, It would not affect your daily excel operation, Be patience and try to dump this screen to CST Support, Thanks much. " & Err.Number & " " & Err.Description, 15
    End If
    
End Sub

