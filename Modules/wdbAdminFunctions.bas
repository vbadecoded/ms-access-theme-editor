Option Compare Database
Option Explicit

Public Sub handleError(modName As String, activeCon As String, errDesc As String, errNum As Long, Optional dataTag As String = "")
On Error Resume Next

Select Case errNum
    Case 70
        MsgBox "Permissions Error - Check if the file is already in use.", vbInformation, "Error Code: " & errNum
    Case 53
        MsgBox "File Not Found", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 3011
        MsgBox "Looks like I'm having issues connecting to SharePoint. Please reopen when you can", vbInformation, "Error Code: " & errNum
    Case 490, 52, 75
        MsgBox "I cannot open this file or location - check if it has been moved or deleted. Or - you do not have proper access to this location", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 3022
        MsgBox "A record with this key already exists. I cannot create another!", vbInformation, "Error Code: " & errNum
    Case 3167
        MsgBox "Looks like you already deleted that record", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 94
        MsgBox "Hmm. Looks like something is missing. Check for an empty field", vbInformation, "Error Code: " & errNum
    Case 3151
        MsgBox "You're not connected to Oracle. Just FYI, Oracle connection does not work outside of VMWare.", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 429
        If modName = "frmCatiaMacros" Then
            MsgBox "Looks like Catia isn't open", vbInformation, "Error Code: " & errNum
            Exit Sub
        Else
            MsgBox errDesc, vbInformation, "Error Code: " & errNum
        End If
    Case 3343
        MsgBox "Error. Please re-open WorkingDB to reset.", vbCritical, "Error Code: " & errNum
    Case Else
        MsgBox errDesc, vbInformation, "Error Code: " & errNum
End Select

End Sub

Function ap_DisableShift()

On Error GoTo errDisableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

db.Properties("AllowByPassKey") = False
Set db = Nothing
Exit Function

errDisableShift:
If Err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, False)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Function
End If

End Function

Function ap_EnableShift()

On Error GoTo errEnableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

db.Properties("AllowByPassKey") = True
Set db = Nothing
Exit Function

errEnableShift:
If Err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, True)
    db.Properties.Append prop
    Resume Next
    Else
    MsgBox "Function 'ap_DisableShift' did not complete successfully."
    Exit Function
End If

End Function