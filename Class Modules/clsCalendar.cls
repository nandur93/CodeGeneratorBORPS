VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is
'Copyright (c) 2010, Jonathon English
'All rights reserved, under the terms of the "New BSD License."

Private mDisplayMonth As Date
Private mStartDate As Date
Private mSelectedDate As Date
Public sngX As Single
Public sngY As Single
Dim mForm As Object
'VBA Bug: use "As Object" for late binding; "As MSForms.UserForm" won't work
Private mBoundControl As Object

Public Property Set Form(BoundControl As Object, myUserform As Object)   'tie the class to a specific form
'Check somehow to make sure the form has the required controls...
On Error GoTo Catch
    Dim mask As Integer
    Dim ctrl As Control
    For Each ctrl In myUserform.Controls
        If ctrl.Name = "FrameCalendar" Then
            mask = mask Or 1  '01
        ElseIf ctrl.Name = "LabelClickArea" Then
            mask = mask Or 2  '10 in binary
        End If
        If mask = 3 Then Exit For   'both bits set is 11, or 3 in non-binary
    Next
    If mask <> 3 Then Err.Raise vbObjectError, "clsCalendar > Property Set Form", _
        "'" & myUserform.Name & "' is missing required calendar controls."
'Form looks like it has the right controls, continue:
    Set mForm = myUserform
    Set mBoundControl = BoundControl
    With mForm
        With .LabelUpDown
            With .Font
                .Name = "Marlett"
                .Size = 11
                .Charset = 2
            End With
            .Caption = "v"
        End With
    End With
    
Finally:
    Exit Property
Catch:
    Stop
    Resume
End Property

Public Property Get SelectedDate() As Date
    SelectedDate = mSelectedDate
End Property

Sub CaptureClick()
Dim intRow As Integer, intCol As Integer

'Here we determine which area in the 9 x 7 grid was clicked:
intRow = sngY / 12 + 0.5
intCol = sngX / 12 + 0.5

'now determine what to do about it
Select Case intRow
    Case 1
        Select Case intCol
            Case 1   'go back one month
                mDisplayMonth = DateSerial(Year(mDisplayMonth), Month(mDisplayMonth) - 1, 1)
            Case 6
                'this is the tough one... was it up or down?
                Select Case sngY
                    Case Is <= 5 'go forward one year
                        mDisplayMonth = DateSerial(Year(mDisplayMonth) + 1, Month(mDisplayMonth), 1)
                    Case Else  'go back one year
                        mDisplayMonth = DateSerial(Year(mDisplayMonth) - 1, Month(mDisplayMonth), 1)
                End Select
            Case 7   'go forward one month
                mDisplayMonth = DateSerial(Year(mDisplayMonth), Month(mDisplayMonth) + 1, 1)
        End Select
    Case 3, 4, 5, 6, 7, 8
        'select days...
        mSelectedDate = mStartDate + (intRow - 3) * 7 + (intCol - 1)
        mBoundControl.Value = SelectedDate
        mForm.FrameCalendar.Visible = False
        Exit Sub
    Case 9
        Select Case intCol
            Case 1, 2, 3, 4
                mDisplayMonth = Date
            Case 5, 6, 7
                mForm.FrameCalendar.Visible = False
                Exit Sub
        End Select
End Select

LoadView mDisplayMonth

End Sub


Sub LoadView(dtDate As Date)

Dim frmForm As UserForm
Dim dtToday As Date
Dim lngInactiveTextColor As Long
Dim lngRegularTextColor As Long
Dim lngHighlightBackColor As Long
Dim lngRegularBackColor As Long
mDisplayMonth = dtDate
dtToday = Date
lngInactiveTextColor = &H80000010  'button shadow
lngRegularTextColor = &H80000012  'button text
lngHighlightBackColor = &H80000018  'Tooltip
lngRegularBackColor = &H80000014  'button highlight

'To find out where to start our calendar, I want to find the first
'Sunday BEFORE the start of the month...
If Weekday(dtDate) = 1 Then
    mStartDate = dtDate
Else
    mStartDate = dtDate - Day(dtDate) - Weekday(dtDate - Day(dtDate)) + 1
End If

With mForm
    .LabelMonth = Format(dtDate, "mmm")
    .LabelYear = Format(dtDate, "yyyy")
    With .Label1
        .Caption = Format(mStartDate, "d")
        If Day(mStartDate) > 7 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label2
        .Caption = Format(mStartDate + 1, "d")
        If Day(mStartDate + 1) > 7 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 1 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label3
        .Caption = Format(mStartDate + 2, "d")
        If Day(mStartDate + 2) > 7 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 2 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label4
        .Caption = Format(mStartDate + 3, "d")
        If Day(mStartDate + 3) > 7 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 3 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label5
        .Caption = Format(mStartDate + 4, "d")
        If Day(mStartDate + 4) > 7 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 4 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label6
        .Caption = Format(mStartDate + 5, "d")
        If Day(mStartDate + 5) > 7 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 5 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label7
        .Caption = Format(mStartDate + 6, "d")
        If mStartDate + 6 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label8
        .Caption = Format(mStartDate + 7, "d")
        If mStartDate + 7 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label9
        .Caption = Format(mStartDate + 8, "d")
        If mStartDate + 8 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label10
        .Caption = Format(mStartDate + 9, "d")
        If mStartDate + 9 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label11
        .Caption = Format(mStartDate + 10, "d")
        If mStartDate + 10 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label12
        .Caption = Format(mStartDate + 11, "d")
        If mStartDate + 11 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label13
        .Caption = Format(mStartDate + 12, "d")
        If mStartDate + 12 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label14
        .Caption = Format(mStartDate + 13, "d")
        If mStartDate + 13 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label15
        .Caption = Format(mStartDate + 14, "d")
        If mStartDate + 14 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label16
        .Caption = Format(mStartDate + 15, "d")
        If mStartDate + 15 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label17
        .Caption = Format(mStartDate + 16, "d")
        If mStartDate + 16 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label18
        .Caption = Format(mStartDate + 17, "d")
        If mStartDate + 17 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label19
        .Caption = Format(mStartDate + 18, "d")
        If mStartDate + 18 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label20
        .Caption = Format(mStartDate + 19, "d")
        If mStartDate + 19 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label21
        .Caption = Format(mStartDate + 20, "d")
        If mStartDate + 20 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label22
        .Caption = Format(mStartDate + 21, "d")
        If mStartDate + 21 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label23
        .Caption = Format(mStartDate + 22, "d")
        If mStartDate + 22 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label24
        .Caption = Format(mStartDate + 23, "d")
        If mStartDate + 23 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label25
        .Caption = Format(mStartDate + 24, "d")
        If mStartDate + 24 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label26
        .Caption = Format(mStartDate + 25, "d")
        If mStartDate + 25 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label27
        .Caption = Format(mStartDate + 26, "d")
        If mStartDate + 26 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label28
        .Caption = Format(mStartDate + 27, "d")
        If mStartDate + 27 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label29
        .Caption = Format(mStartDate + 28, "d")
        If Day(mStartDate + 28) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 28 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label30
        .Caption = Format(mStartDate + 29, "d")
        If Day(mStartDate + 29) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 29 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label31
        .Caption = Format(mStartDate + 30, "d")
        If Day(mStartDate + 30) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 30 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label32
        .Caption = Format(mStartDate + 31, "d")
        If Day(mStartDate + 31) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 31 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label33
        .Caption = Format(mStartDate + 32, "d")
        If Day(mStartDate + 32) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 32 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label34
        .Caption = Format(mStartDate + 33, "d")
        If Day(mStartDate + 33) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 33 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label35
        .Caption = Format(mStartDate + 34, "d")
        If Day(mStartDate + 34) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 34 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label36
        .Caption = Format(mStartDate + 35, "d")
        If Day(mStartDate + 35) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 35 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label37
        .Caption = Format(mStartDate + 36, "d")
        If Day(mStartDate + 36) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 36 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label38
        .Caption = Format(mStartDate + 37, "d")
        If Day(mStartDate + 37) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 37 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label39
        .Caption = Format(mStartDate + 38, "d")
        If Day(mStartDate + 38) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 38 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label40
        .Caption = Format(mStartDate + 39, "d")
        If Day(mStartDate + 39) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 39 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label41
        .Caption = Format(mStartDate + 40, "d")
        If Day(mStartDate + 40) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 40 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
    With .Label42
        .Caption = Format(mStartDate + 41, "d")
        If Day(mStartDate + 41) < 15 Then .ForeColor = lngInactiveTextColor Else .ForeColor = lngRegularTextColor
        If mStartDate + 41 = dtToday Then .BackColor = lngHighlightBackColor Else .BackColor = lngRegularBackColor
    End With
End With

End Sub

