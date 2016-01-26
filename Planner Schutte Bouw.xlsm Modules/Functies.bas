Attribute VB_Name = "Functies"
Option Explicit

Sub errorhandler_MsgBox(error As String)
MsgBox error, vbCritical, "FOUT"
End Sub

Sub Turbo_AAN()
'Get current state of various Excel settings; put this at the beginning of your code
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting
End Sub

Sub turbo_UIT()
'after your code runs, restore state; put this at the end of your code
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.Calculation = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True 'note this is a sheet-level setting
End Sub
Public Function IsoWeekNumber(d1 As Date) As Integer
   'Attributed to Daniel Maher
   Dim d2 As Long
   d2 = DateSerial(Year(d1 - Weekday(d1 - 1) + 4), 1, 3)
   IsoWeekNumber = Int((d1 - d2 + Weekday(d2) + 5) / 7)
End Function

Function MaakRaster(rng As Range)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Function


