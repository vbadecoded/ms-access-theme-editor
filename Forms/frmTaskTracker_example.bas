Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function filterIt(controlName As String)
On Error GoTo Err_Handler

Me(controlName).SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Function
Err_Handler:
    Call handleError(Me.Name, "filterIt", Err.DESCRIPTION, Err.Number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.Filter = "completed_date is null"
Me.FilterOn = True

Me.OrderBy = "Due_date"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.Name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub newTask_Click()
On Error GoTo Err_Handler

MsgBox "No sample form here, just a sample button to show an 'action button'"

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmTaskDetails_example", , , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler

Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.Name, Me.ActiveControl.Name, Err.DESCRIPTION, Err.Number)
End Sub
