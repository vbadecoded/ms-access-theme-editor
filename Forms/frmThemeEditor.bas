Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function applyThemeChanges()

TempVars.Add "themePrimary", Me.primaryColor.Value
TempVars.Add "themeSecondary", Me.secondaryColor.Value

If Me.darkMode Then
    TempVars.Add "themeMode", "Dark"
Else
    TempVars.Add "themeMode", "Light"
End If

TempVars.Add "themeColorLevels", Me.colorLevels.Value

DoCmd.Hourglass True
Me.Painting = False
DoCmd.Echo False

Call setTheme(Me)
Call setTheme(Me.sfrmThemeEditor.Form)

DoCmd.Hourglass False
Me.Painting = True
DoCmd.Echo True

End Function
Private Sub colorLevels_AfterUpdate()
splitColorArray
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor

End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

applyThemeChanges

splitColorArray
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Function applyLevels()

Select Case ""
    Case Nz(Me.L1), Nz(Me.L2), Nz(Me.L3), Nz(Me.L4)
        Exit Function
    Case Else
        Me.colorLevels = Me.L1 & "," & Me.L2 & "," & Me.L3 & "," & Me.L4
        applyThemeChanges
End Select

End Function

Public Function splitColorArray()

Dim splitIt() As String

splitIt = Split(Me.colorLevels, ",")

Me.L1 = splitIt(0)
Me.L2 = splitIt(1)
Me.L3 = splitIt(2)
Me.L4 = splitIt(3)

End Function

Private Sub L1_AfterUpdate()
applyLevels
End Sub

Private Sub L2_AfterUpdate()
applyLevels
End Sub

Private Sub L3_AfterUpdate()
applyLevels
End Sub

Private Sub L4_AfterUpdate()
applyLevels
End Sub

Private Sub newTheme_Click()
DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub primaryColor_Click()

If Me.Dirty Then Me.Dirty = False
Me.ActiveControl = colorPicker(Me.ActiveControl)

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor

applyThemeChanges

End Sub

Private Sub secondaryColor_Click()

If Me.Dirty Then Me.Dirty = False
Me.ActiveControl = colorPicker(Me.ActiveControl)

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor

applyThemeChanges

End Sub

Private Sub testTheme_Click()
If Me.Dirty Then Me.Dirty = False
applyThemeChanges
End Sub
