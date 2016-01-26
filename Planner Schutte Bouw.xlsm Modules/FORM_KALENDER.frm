VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_KALENDER 
   Caption         =   "DATUM SELECTEREN"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3330
   OleObjectBlob   =   "FORM_KALENDER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_KALENDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private aanpassingen As Boolean
Public CKalender As Collection
Public vandaagdatum As Date
Public datumingeladen As Date




Private Sub CheckBox1_Click()
Dim d As datum
If IsDate(TextBoxDatum) = True Then
    If CheckBox1.Value Then TextBoxOpmerking.Enabled = True Else TextBoxOpmerking.Enabled = False
    For Each d In CKalender
        If d.datum = CDate(TextBoxDatum) Then
            If d.Feestdag <> CheckBox1.Value Then
                d.Feestdag = CheckBox1.Value
                d.update
                aanpassingen = True
                updatekalender
            End If
        End If
    Next d
End If
End Sub


Private Sub CheckBox3_Change()
Dim d As datum
If IsDate(TextBoxDatum) = True Then
    For Each d In CKalender
        If d.datum = CDate(TextBoxDatum) Then
            If d.Zichtbaar <> CheckBox3.Value Then
                d.Zichtbaar = CheckBox3.Value
                d.update
                aanpassingen = True
                updatekalender
            End If
        End If
    Next d
End If
End Sub


Private Sub CommandButton1_Click()

Dim Startdatum As Date
Dim datum As Date
Dim vandaag As Date: vandaag = FormatDateTime(Now(), vbShortDate)
Dim lbl As MSForms.label

ThisWorkbook.inladen = False
Call UserForm_Initialize

Startdatum = CDate(LabelStartdatum)
LabelVandaag = vandaag
For x = 1 To 42
    Set lbl = Me.Controls("Label" & x)
    datum = DateAdd("d", x - 1, Startdatum)
    
    If datum = vandaag Then

        'lbl.BorderStyle = fmBorderStyleSingle
        setData ("Label" & x)
    Else
        'lbl.BorderStyle = fmBorderStyleNone
    End If
Next x
End Sub

Private Sub CommandButton2_Click()
Dim datum As Date
If TextBoxDatum = "" Then ThisWorkbook.inladen = False

If IsDate(TextBoxDatum) = True Then
    datum = CDate(TextBoxDatum)
    ThisWorkbook.datum = datum
    ThisWorkbook.inladen = True
End If

Unload FORM_KALENDER
End Sub

Private Sub Label1_Click()
setData (Label1.Name)
End Sub

Private Sub Label2_Click()
setData (Label2.Name)
End Sub

Private Sub Label3_Click()
setData (Label3.Name)
End Sub

Private Sub Label4_Click()
setData (Label4.Name)
End Sub

Private Sub Label5_Click()
setData (Label5.Name)
End Sub

Private Sub Label6_Click()
setData (Label6.Name)
End Sub

Private Sub Label7_Click()
setData (Label7.Name)
End Sub

Private Sub Label8_Click()
setData (Label8.Name)
End Sub

Private Sub Label9_Click()
setData (Label9.Name)
End Sub

Private Sub Label10_Click()
setData (Label10.Name)
End Sub

Private Sub Label11_Click()
setData (Label11.Name)
End Sub

Private Sub Label12_Click()
setData (Label12.Name)
End Sub

Private Sub Label13_Click()
setData (Label13.Name)
End Sub

Private Sub Label14_Click()
setData (Label14.Name)
End Sub

Private Sub Label15_Click()
setData (Label15.Name)
End Sub

Private Sub Label16_Click()
setData (Label16.Name)
End Sub

Private Sub Label17_Click()
setData (Label17.Name)
End Sub

Private Sub Label18_Click()
setData (Label18.Name)
End Sub

Private Sub Label19_Click()
setData (Label19.Name)
End Sub

Private Sub Label20_Click()
setData (Label20.Name)
End Sub

Private Sub Label21_Click()
setData (Label21.Name)
End Sub

Private Sub Label22_Click()
setData (Label22.Name)
End Sub

Private Sub Label23_Click()
setData (Label23.Name)
End Sub

Private Sub Label24_Click()
setData (Label24.Name)
End Sub

Private Sub Label25_Click()
setData (Label25.Name)
End Sub

Private Sub Label26_Click()
setData (Label26.Name)
End Sub

Private Sub Label27_Click()
setData (Label27.Name)
End Sub

Private Sub Label28_Click()
setData (Label28.Name)
End Sub

Private Sub Label29_Click()
setData (Label29.Name)
End Sub

Private Sub Label30_Click()
setData (Label30.Name)
End Sub

Private Sub Label31_Click()
setData (Label31.Name)
End Sub

Private Sub Label32_Click()
setData (Label32.Name)
End Sub

Private Sub Label33_Click()
setData (Label33.Name)
End Sub

Private Sub Label34_Click()
setData (Label34.Name)
End Sub

Private Sub Label35_Click()
setData (Label35.Name)
End Sub

Private Sub Label36_Click()
setData (Label36.Name)
End Sub

Private Sub Label37_Click()
setData (Label37.Name)
End Sub

Private Sub Label38_Click()
setData (Label38.Name)
End Sub

Private Sub Label39_Click()
setData (Label39.Name)
End Sub

Private Sub Label40_Click()
setData (Label40.Name)
End Sub

Private Sub Label41_Click()
setData (Label41.Name)
End Sub

Private Sub Label42_Click()
setData (Label42.Name)
End Sub


Private Sub TextBoxOpmerking_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim d As datum
If TextBoxOpmerking <> "" Or IsDate(TextBoxDatum) = True Then
    For Each d In CKalender
        If d.datum = CDate(TextBoxDatum) Then
            If d.Omschrijving <> TextBoxOpmerking.Value Then
                d.Omschrijving = TextBoxOpmerking.Value
                d.update
                updatekalender
            End If
        End If
    Next d
End If


End Sub


Private Sub UserForm_Initialize()
Dim vandaag As Date
Dim Startdatum As Date
Dim Einddatum As Date
Dim synergy As String
vandaag = FormatDateTime(Now(), vbShortDate)
datumingeladen = ThisWorkbook.datum
Me.vandaagdatum = FormatDateTime(Now(), vbShortDate)
If ThisWorkbook.inladen Then
    Startdatum = eersteMaandagInDeMaand(ThisWorkbook.datum)
    Einddatum = laatstezondagindemaand(ThisWorkbook.datum)
    labelJaar = Year(ThisWorkbook.datum)
    labelMaand = Month(ThisWorkbook.datum)
    LabelTextMaand = DatumNaarMaand(ThisWorkbook.datum)
    LabelVandaag = ThisWorkbook.datum
Else
    Startdatum = eersteMaandagInDeMaand(vandaag)
    Einddatum = laatstezondagindemaand(vandaag)
    labelJaar = Year(vandaag)
    labelMaand = Month(vandaag)
    LabelTextMaand = DatumNaarMaand(vandaag)
    LabelVandaag = vandaag
End If

Set CKalender = Lijsten.KalenderGeheel

LabelStartdatum = Startdatum
updatekalender

    LabelOmschrijivng = ThisWorkbook.infokalender

End Sub

Function eersteMaandagInDeMaand(datum As Date) As Date
eersteMaandagInDeMaand = DateAdd("d", _
    1 - Weekday(DateSerial(Year(datum), Month(datum), 1), vbMonday), _
    DateSerial(Year(datum), Month(datum), 1))
End Function

Function laatstezondagindemaand(datum As Date) As Date
laatstezondagindemaand = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, datum)), Month(DateAdd("m", 1, datum)), 1))
laatstezondagindemaand = DateAdd("d", 7 - Weekday(laatstezondagindemaand, vbMonday), laatstezondagindemaand)
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'If FORM_PROJECT_AANMAKEN.Visible = True Then Exit Sub
If aanpassingen = True Then
DetailPlanning.MaakDetailPlanning
End If

End Sub

Private Sub Volgende_Jaar_Click()
Dim datum As Date

datum = DateAdd("yyyy", 1, CDate(LabelVandaag))
    
    
    LabelVandaag = datum
    LabelStartdatum = eersteMaandagInDeMaand(datum)
    labelJaar = Year(datum)
    labelMaand = Month(datum)
    LabelTextMaand = DatumNaarMaand(datum)
    
updatekalender
End Sub

Private Sub Volgende_Maand_Click()
Dim datum As Date

datum = DateAdd("m", 1, CDate(LabelVandaag))
    LabelVandaag = datum
    LabelStartdatum = eersteMaandagInDeMaand(datum)
    labelJaar = Year(datum)
    labelMaand = Month(datum)
    LabelTextMaand = DatumNaarMaand(datum)
    
updatekalender
End Sub
Private Sub Vorige_Maand_Click()
Dim datum As Date

datum = DateAdd("m", -1, CDate(LabelVandaag))
    
    LabelVandaag = datum
    LabelStartdatum = eersteMaandagInDeMaand(datum)
    labelJaar = Year(datum)
    labelMaand = Month(datum)
    LabelTextMaand = DatumNaarMaand(datum)
    
updatekalender
End Sub
Private Sub Vorige_Jaar_Click()
Dim datum As Date

datum = DateAdd("yyyy", -1, CDate(LabelVandaag))
    
    LabelVandaag = datum
    LabelStartdatum = eersteMaandagInDeMaand(datum)
    labelJaar = Year(datum)
    labelMaand = Month(datum)
    LabelTextMaand = DatumNaarMaand(datum)
    
updatekalender
End Sub



Function DatumNaarMaand(datum As Date) As String

Select Case Month(datum)
    Case 1
        DatumNaarMaand = "Januari"
    Case 2
        DatumNaarMaand = "Februari"
    Case 3
        DatumNaarMaand = "Maart"
    Case 4
        DatumNaarMaand = "April"
    Case 5
        DatumNaarMaand = "Mei"
    Case 6
        DatumNaarMaand = "Juni"
    Case 7
        DatumNaarMaand = "Juli"
    Case 8
        DatumNaarMaand = "Augustus"
    Case 9
        DatumNaarMaand = "September"
    Case 10
        DatumNaarMaand = "Oktober"
    Case 11
        DatumNaarMaand = "November"
    Case 12
        DatumNaarMaand = "December"
End Select

End Function

Function updatekalender()
Dim datum As Date
Dim lbl As MSForms.label
Dim nietzichtbaar As Boolean
Dim txtdatum As Date
If ThisWorkbook.inladen = True Then
    ThisWorkbook.inladen = False
    datum = ThisWorkbook.datum
    Startdatum = eersteMaandagInDeMaand(datum)
    Einddatum = laatstezondagindemaand(datum)
Else
datum = CDate(LabelVandaag)
Startdatum = eersteMaandagInDeMaand(datum)
Einddatum = laatstezondagindemaand(datum)
End If

For x = 1 To 42
    txtdatum = DateAdd("d", x - 1, Startdatum)
    
    Set lbl = Me.Controls("Label" & x)
    If txtdatum = ThisWorkbook.datum Then CallByName lbl, "BorderStyle", VbLet, 1
    lbl.Caption = Day(txtdatum)
    updateDatumTextBox CStr(x), txtdatum
    
        If x = 36 And lbl.Caption < 15 And lbl.Caption <> 0 Then
        nietzichtbaar = True
       
        Else
        LabelWeek6.Visible = True
        End If
        
        If nietzichtbaar = True Then
            lbl.Visible = False
        Else
            lbl.Visible = True
        End If
        
        If x = 1 Then LabelWeek1 = IsoWeekNumber(txtdatum)
        If x = 8 Then LabelWeek2 = IsoWeekNumber(txtdatum)
        If x = 15 Then LabelWeek3 = IsoWeekNumber(txtdatum)
        If x = 22 Then LabelWeek4 = IsoWeekNumber(txtdatum)
        If x = 29 Then LabelWeek5 = IsoWeekNumber(txtdatum)
        If x = 36 Then LabelWeek6 = IsoWeekNumber(txtdatum)
        
        
        
Next x
If nietzichtbaar = True Then
 LabelWeek6.Visible = False
 Else
 LabelWeek6.Visible = True
End If
 
End Function

Function updateDatumTextBox(x As String, datum As Date)
Dim d As datum
Dim lbl As MSForms.label: Set lbl = Me.Controls("Label" & x)
Dim aanwezig As Boolean
Dim vandaag As Date: vandaag = FormatDateTime(Now(), vbShortDate)

lbl.BackColor = vbButtonFace
lbl.ForeColor = 1
For Each k In Me.CKalender
    If k.datum = datum Then
        If k.Feestdag = True Then lbl.BackColor = 255
        aanwezig = k.Zichtbaar
        Exit For

    End If
Next k

If aanwezig = False Then lbl.ForeColor = -2147483632 Else lbl.ForeColor = 1

If vandaag = datum Then lbl.BackColor = 255255

End Function

Function setData(label As String)
Dim datum As Date
Dim d As datum
Dim lbl As MSForms.label



x = CStr(Mid(label, 6))
datum = DateAdd("d", x - 1, CDate(LabelStartdatum))
TextBoxDatum = datum
For Each d In Me.CKalender
    If d.datum = datum Then
    CheckBox1 = d.Feestdag
    CheckBox3 = d.Zichtbaar
    TextBoxOpmerking = d.Omschrijving
    End If
Next d

For x = 1 To 42
Set lbl = Me.Controls("Label" & x)
If "Label" & x = label Then
    If lbl.ForeColor = -2147483632 Then TextBoxDatum = ""
    CallByName lbl, "BorderStyle", VbLet, 1
Else
    CallByName lbl, "BorderStyle", VbLet, 0
End If

Next x
End Function

Sub test()
Dim t As MSForms.TextBox
For Each t In Me

Next t

End Sub
