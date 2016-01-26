Attribute VB_Name = "Module1"
Option Explicit

Function testPersoon()
Dim p As New Persoon

p.afkorting = "ABD"
p.GetByafkorting
p.Print_r
End Function
