Option Compare Database
Option Explicit

Function dueDay(dueDate, completeddate) As String
On Error Resume Next

If IsNull(dueDate) Then
    dueDay = "N/A"
    Exit Function
End If

If IsNull(completeddate) Then
    Select Case dueDate
        Case date
            dueDay = "Today"
        Case date + 1
            dueDay = "Tomorrow"
        Case Is < date
            dueDay = "Overdue"
        Case Is < date + 7
            dueDay = WeekdayName(Weekday(dueDate))
        Case date + 7
            dueDay = "1 Week"
        Case Is < date + 14
            dueDay = "<2 Weeks"
        Case date + 14
            dueDay = "2 Weeks"
        Case Is < date + 21
            dueDay = "<3 Weeks"
        Case date + 21
            dueDay = "3 Weeks"
        Case Is < date + 28
            dueDay = "<4 Weeks"
        Case date + 28
            dueDay = "4 Weeks"
        Case Is > date + 28
            dueDay = ">4 Weeks"
        Case Else
            dueDay = dueDate
    End Select
Else
    dueDay = "Complete"
End If

End Function

Function randomNumber(low As Long, high As Long) As Long

Randomize
randomNumber = Int((high - low + 1) * Rnd() + low)

End Function

Function generateValues()

Dim db As Database
Dim rs As Recordset

Set db = CurrentDb()
Set rs = db.OpenRecordset("tblTaskTracker_example")

Do While Not rs.EOF
    rs.Edit
    
    rs!Request_Type = randomNumber(1, 7)
    rs!complexity = randomNumber(1, 3)
    rs!Assignee = randomNumber(1, 10)
    rs!Checker_1 = randomNumber(1, 10)
    rs!Checker_2 = randomNumber(1, 10)
    rs!Customer = randomNumber(1, 3)
    rs!Delay_Reason = randomNumber(1, 3)
    rs!Status = randomNumber(1, 6)
    
    rs.Update
    
    rs.MoveNext
Loop

End Function