Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

'here is the standard VBA call for applying a theme to a form
Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.Name, "Form_Load", Err.DESCRIPTION, Err.Numbe)
End Sub

Private Sub linkArticle_Click()
On Error GoTo Err_Handler

FollowHyperlink "https://www.vbadecoded.com/ms-access-vba/user-themes"

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub linkGithub_Click()
On Error GoTo Err_Handler

FollowHyperlink "https://github.com/vbadecoded/ms-access-theme-editor"

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub sampleTracker_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmTaskTracker_example"

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub themeEditor_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmThemeEditor"

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub
