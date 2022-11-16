Sub CountTimeSpent()
Dim oOLApp As Outlook.Application
Dim oSelection As Outlook.Selection
Dim oItem As Object
Dim iDuration As Long
Dim iDuration2 As Long
Dim iTotalWork As Long
Dim iMileage As Long
Dim iResult As Integer
Dim bShowiMileage As Boolean
Dim objAttendees As Outlook.Recipients
Dim objAttendee As Outlook.Recipient
Dim lRequiredAttendeeCount, lOptionalAttendeeCount, lResourceCount As Long
 
bShowiMileage = False
 
iDuration = 0
iDuration2 = 0
iTotalWork = 0
iMileage = 0
 
On Error Resume Next
 
    Set oOLApp = CreateObject("Outlook.Application")
Set oSelection = oOLApp.ActiveExplorer.Selection
 
For Each oItem In oSelection
    If oItem.Class = olAppointment Then
        Set objAttendees = oItem.Recipients
        lRequiredAttendeeCount = 0
        lOptionalAttendeeCount = 0
        lResourceCount = 0
        For Each objAttendee In objAttendees
            If objAttendee.Type = olRequired Then
               lRequiredAttendeeCount = lRequiredAttendeeCount + 1
            ElseIf objAttendee.Type = olOptional Then
               lOptionalAttendeeCount = lOptionalAttendeeCount + 1
            ElseIf objAttendee.Type = olResource Then
               lResourceCount = lResourceCount + 1
            End If
        Next
        If lRequiredAttendeeCount < 2 And lOptionalAttendeeCount = 0 Then
            iDuration2 = iDuration2 + oItem.Duration
        Else
            iDuration = iDuration + oItem.Duration
        End If
    Else
        iResult = MsgBox("Please select some Calendar, Task or Journal items at first!", vbCritical, "Items Time Spent")
        Exit Sub
    End If
Next
 
Dim MsgBoxText As String
MsgBoxText = "Total time spent Meetings: " & vbNewLine & iDuration & " minutes"
    If iDuration > 60 Then
        MsgBoxText = MsgBoxText & HoursMsg(iDuration)
    End If
MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total time spent Tasks: " & vbNewLine & iDuration2 & " minutes"
    If iDuration2 > 60 Then
        MsgBoxText = MsgBoxText & HoursMsg(iDuration2)
    End If
MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total time spent Meetings: " & vbNewLine & iDuration + iDuration2 & " minutes"
    If iDuration2 > 60 Then
        MsgBoxText = MsgBoxText & HoursMsg(iDuration + iDuration2)
    End If
MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total time spent Meetings+Tasks: " & vbNewLine & (iDuration + iDuration2) / 60 & " hours"
    If iDuration2 > 60 Then
        MsgBoxText = MsgBoxText & HoursMsg(iDuration + iDuration2)
    End If
 
iResult = MsgBox(MsgBoxText, vbInformation, "Items Time spent")

ExitSub:
Set oItem = Nothing
Set oSelection = Nothing
Set oOLApp = Nothing
End Sub

Function HoursMsg(TotalMinutes As Long) As String
Dim iHours As Long
Dim iMinutes As Long
iHours = TotalMinutes \ 60
iMinutes = TotalMinutes Mod 60
HoursMsg = " (" & iHours & " Hours and " & iMinutes & " Minutes)"
End Function
