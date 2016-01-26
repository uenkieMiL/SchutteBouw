Attribute VB_Name = "Lijsten"
Option Explicit

Public Function KalenderGeheel() As Collection
Dim strSQL As String
Dim lijst As Variant
Dim d As datum
Dim db As New db
Dim r As Long
Dim a As Long

Set KalenderGeheel = New Collection

lijst = db.getLijstBySQL("Select * from Kalender")

If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        Set d = New datum
        d.FromList r, lijst
        If d.Zichtbaar = False Then
            a = a + 1
            d.Kolomnummer = -1
        Else
            d.Kolomnummer = r - a
        End If
      
        KalenderGeheel.Add d, CStr(d.datum)
    Next r
End If
Set db = Nothing
Set lijst = Nothing
End Function
