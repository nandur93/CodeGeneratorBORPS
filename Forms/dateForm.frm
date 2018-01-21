VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dateForm 
   Caption         =   "Production Code Generator"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   OleObjectBlob   =   "dateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================= kode kalender ===========================================
'This code is
'Copyright (c) 2010, Jonathon English
'All rights reserved, under the terms of the "New BSD License."

Dim clsCal As clsCalendar

Private Sub ButtonAbout_Click()
    Dim msg As Integer
    msg = MsgBox("Developed by NDU", vbInformation, "Code Generator V1.0")
End Sub

Private Sub CheckBoxBesok_Click()
    If LabelTodayNow.Caption = "Sunday" _
        Or LabelTodayNow.Caption = "Monday" _
        Or LabelTodayNow.Caption = "Tuesday" _
        Or LabelTodayNow.Caption = "Wednesday" _
        Or LabelTodayNow.Caption = "Thursday" _
        Or LabelTodayNow.Caption = "Friday" _
        Or LabelTodayNow.Caption = "Saturday" Then
        If CheckBoxBesok.Value = True Then 'jika cek aktif
                If TextBoxDate = "" Then
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 1, TextBoxTanggalBO) 'maka tanggal + 1
                    End With
                Else
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 1, TextBoxDate) 'maka tanggal + 1
                    End With
                End If
            TextBoxTanggalForm = Format(TextBoxTanggalForm, "mm/dd/yy")
        ElseIf CheckBoxBesok.Value = False Then 'jika cek tidak aktif
                If TextBoxDate = "" Then
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 0, TextBoxTanggalBO) 'maka tanggal tetap
                    End With
                Else
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 0, TextBoxDate) 'maka tanggal tetap
                    End With
                End If
            TextBoxTanggalForm = Format(TextBoxTanggalForm, "mm/dd/yy")
        End If
    Else
        If CheckBoxBesok.Value = True Then 'jika cek aktif
                If TextBoxDate = "" Then
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 1, TextBoxTanggalBO) 'maka tanggal + 1
                    End With
                Else
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 1, TextBoxDate) 'maka tanggal + 1
                    End With
                End If
            TextBoxTanggalForm = Format(TextBoxTanggalForm, "dd/mm/yy")
        ElseIf CheckBoxBesok.Value = False Then 'jika cek tidak aktif
                If TextBoxDate = "" Then
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 0, TextBoxTanggalBO) 'maka tanggal tetap
                    End With
                Else
                    With TextBoxTanggalForm
                        .Enabled = True
                        .Value = DateAdd("d", 0, TextBoxDate) 'maka tanggal tetap
                    End With
                End If
            TextBoxTanggalForm = Format(TextBoxTanggalForm, "dd/mm/yy")
        End If
    End If
End Sub

Private Sub ComboBoxLineSachet_Change()
If ComboBoxLineSachet = "A" Then
Cells.Find(What:="LINE A", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "B1" Or ComboBoxLineSachet = "B2" Or ComboBoxLineSachet = "B3" Then
Cells.Find(What:="LINE B", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "C" Then
Cells.Find(What:="LINE C", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "D1" Then
Cells.Find(What:="LINE D1", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "D2" Then
Cells.Find(What:="LINE D2", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "E1" Then
Cells.Find(What:="LINE E1", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "E2" Then
Cells.Find(What:="LINE E2", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "F1" Then
Cells.Find(What:="LINE F1", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "G1" Then
Cells.Find(What:="LINE G1", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "G2" Then
Cells.Find(What:="LINE G2", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "H1" Then
Cells.Find(What:="LINE H1", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "H2" Then
Cells.Find(What:="LINE H2", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "I1" Then
Cells.Find(What:="LINE I1", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "I2" Then
Cells.Find(What:="LINE I2", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "I3" Then
Cells.Find(What:="LINE I3", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
ElseIf ComboBoxLineSachet = "J" Then
Cells.Find(What:="LINE J", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
End If
End Sub

Private Sub LabelDateNow_Click()
    TextBoxDate = Format(Now)
    CheckBoxBesok.Value = False
End Sub

Private Sub LabelTodayNow_Click()
    TextBoxDate = Format(Now)
    CheckBoxBesok.Value = False
End Sub

Private Sub TextBoxDate_Change()
    CheckBoxBesok.Value = False
        If LabelTodayNow.Caption = "Sunday" _
            Or LabelTodayNow.Caption = "Monday" _
            Or LabelTodayNow.Caption = "Tuesday" _
            Or LabelTodayNow.Caption = "Wednesday" _
            Or LabelTodayNow.Caption = "Thursday" _
            Or LabelTodayNow.Caption = "Friday" _
            Or LabelTodayNow.Caption = "Saturday" Then
                TextBoxTanggalBO = Format(TextBoxDate, "mm/dd/yy")
                TextBoxTanggalForm = Format(TextBoxDate, "mm/dd/yy")
                    Else
                TextBoxTanggalBO = Format(TextBoxDate, "dd/mm/yy")
                TextBoxTanggalForm = Format(TextBoxDate, "dd/mm/yy")
        End If
End Sub

'Private Sub TextBoxDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'    CheckBoxBesok.Value = False
'    TextBoxTanggalBO = Format(TextBoxDate, "dd/mm/yy")
'    TextBoxTanggalForm = Format(TextBoxDate, "dd/mm/yy")
'End Sub

Private Sub TextBoxTanggalBO_Change()
If TextBoxTanggalBO = "" Then
    CheckBoxBesok.Enabled = False
            Else
    CheckBoxBesok.Enabled = True
End If
CheckBoxBesok.Value = False
TextBoxTanggalForm = TextBoxTanggalBO
End Sub

Private Sub UserForm_Initialize()
With dateForm
Judul
End With
LabelTodayNow.Caption = Format(Now, "dddd")
LabelDateNow = Format(Now, "dd/mm/yyyy")

    Windows("ALL NEW VERIFIKASI KODE (DILARANG DI COPY).xlsx").Activate
    Sheets("RPS").Select
    
    CheckBoxBesok.Enabled = False

With ComboBoxLineSachet
    .AddItem "A"
    '.AddItem "A2"
    .AddItem "B1"
    .AddItem "B2"
    .AddItem "B3"
    .AddItem "C"
    '.AddItem "C2"
    .AddItem "D1"
    .AddItem "D2"
    '.AddItem "E1"
    '.AddItem "E2"
    .AddItem "F1"
    .AddItem "F2"
    .AddItem "G1"
    .AddItem "G2"
    '.AddItem "H1"
    '.AddItem "H2"
    .AddItem "I1"
    .AddItem "I2"
    .AddItem "I3"
End With
        
    Set clsCal = New clsCalendar
    Set clsCal.Form(Me.TextBoxDate) = Me
End Sub

'I used to use the MouseDown event, but I found that it wasn't responsive to double-clicks, so now I
'use a combination of the MouseMove event to get the coordinates, and both the Click and DblClick events.
Private Sub LabelClickArea_Click()
    clsCal.CaptureClick
End Sub
Private Sub LabelClickArea_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    clsCal.CaptureClick
    Cancel = True
End Sub
Private Sub LabelClickArea_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With clsCal
        .sngX = X
        .sngY = Y
    End With
End Sub

Private Sub ButtonShowCalendar_Click()
With Me.FrameCalendar
    clsCal.LoadView Date
    .Visible = True
End With
End Sub

Private Sub ButtonGenerate_Click()
'kita buat time swap untuk menghandle error ketika beda timezone
'=======pseucode=======
'jika senin maka dd/mm/yy
'jika monday maka mm/dd/yy
'Dim TextBoxTanggalBO, TextBoxTanggalForm As Date
Dim msg As Integer
    If ComboBoxLineSachet = vbNullString Then
        msg = MsgBox("Line Sachet Tidak Boleh Kosong", vbRetryCancel + vbExclamation, "Peringatan!")
            Else
            Windows("ALL NEW VERIFIKASI KODE (DILARANG DI COPY).xlsx").Activate
            Sheets("MASTER").Select
                        If LabelTodayNow.Caption = "Sunday" _
                            Or LabelTodayNow.Caption = "Monday" _
                            Or LabelTodayNow.Caption = "Tuesday" _
                            Or LabelTodayNow.Caption = "Wednesday" _
                            Or LabelTodayNow.Caption = "Thursday" _
                            Or LabelTodayNow.Caption = "Friday" _
                            Or LabelTodayNow.Caption = "Saturday" Then
                    Range("D6") = Format(TextBoxTanggalBO, "mm/dd/yyyy")
                                Else
                            Range("D6") = Format(TextBoxTanggalBO, "dd/mm/yyyy")
                        End If
            Range("D10") = ComboBoxLineSachet
                If TextBoxNoBO = vbNullString Then
                    Range("D26").ClearContents
                        Else
                    Range("D26") = CDbl(TextBoxNoBO) 'convert to double
                End If
                If TextBoxProdukSebelum = vbNullString Then
                    Range("D30").ClearContents
                        Else
                    Range("D30") = CDbl(TextBoxProdukSebelum) 'convert to double
                End If
            Range("D31") = TextBoxChangeOver
            Range("E31") = TextBoxMaterial
                        If LabelTodayNow.Caption = "Sunday" _
                            Or LabelTodayNow.Caption = "Monday" _
                            Or LabelTodayNow.Caption = "Tuesday" _
                            Or LabelTodayNow.Caption = "Wednesday" _
                            Or LabelTodayNow.Caption = "Thursday" _
                            Or LabelTodayNow.Caption = "Friday" _
                            Or LabelTodayNow.Caption = "Saturday" Then
                            Range("D32") = Format(TextBoxTanggalForm, "mm/dd/yyyy")
                                Else
                            Range("D32") = Format(TextBoxTanggalForm, "dd/mm/yyyy")
                        End If
            Unload Me
    End If
End Sub
Private Sub UserForm_Terminate()
    'MsgBox "Closed by Terminate"
    Set clsCal = Nothing
End Sub
'============================= kode kalender ===========================================

Private Sub ButtonStart_Click()
    BoSebelum
    ChangeOver
    Material
    NoBo
End Sub
Sub BoSebelum()
Dim rRange As Range
    On Error Resume Next

        Application.DisplayAlerts = False

            Set rRange = Application.InputBox(Prompt:= _
                "Klik BO Sebelumnya Pada RPS", _
                    Title:="Produk Sebelum", Type:=8)
    On Error GoTo 0

        Application.DisplayAlerts = True

        If rRange Is Nothing Then
           Exit Sub
        Else
TextBoxProdukSebelum = rRange
        End If
End Sub

Sub NoBo()
Dim rRange As Range
    On Error Resume Next

        Application.DisplayAlerts = False

            Set rRange = Application.InputBox(Prompt:= _
                "Klik NO BO Pada RPS" & vbNewLine & "(Kolom Wajib Diisi)", _
                    Title:="BO Yang Ready", Type:=8)
    On Error GoTo 0

        Application.DisplayAlerts = True

        If rRange Is Nothing Then
           Exit Sub
        Else
TextBoxNoBO = rRange
        End If
End Sub

Sub ChangeOver()
Dim rRange As Range
    On Error Resume Next

        Application.DisplayAlerts = False

            Set rRange = Application.InputBox(Prompt:= _
                "Klik Change Over Pada RPS", _
                    Title:="Perlakuan", Type:=8)
    On Error GoTo 0

        Application.DisplayAlerts = True

        If rRange Is Nothing Then
           Exit Sub
        Else
TextBoxChangeOver = rRange
        End If
End Sub

Sub Material()
Dim rRange As Range
    On Error Resume Next

        Application.DisplayAlerts = False

            Set rRange = Application.InputBox(Prompt:= _
                "Klik Material RPS", _
                    Title:="Material", Type:=8)
    On Error GoTo 0

        Application.DisplayAlerts = True

        If rRange Is Nothing Then
           Exit Sub
        Else
TextBoxMaterial = rRange
        End If
End Sub

