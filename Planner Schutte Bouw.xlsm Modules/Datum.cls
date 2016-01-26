VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Datum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDatum As Date
Private pFeestdag As Boolean
Private pOmschrijving As String
Private pZichtbaar As Boolean
Private pKolomnummer As Long



'''''''''''''''''''''''''
' Datum property
'''''''''''''''''''''''''

Public Property Get datum() As Date
    datum = pDatum
End Property

Public Property Let datum(Value As Date)
    pDatum = Value
End Property


'''''''''''''''''''''''''
' Feestdag property
'''''''''''''''''''''''''

Public Property Get Feestdag() As Boolean
    Feestdag = pFeestdag
End Property

Public Property Let Feestdag(Value As Boolean)
    pFeestdag = Value
End Property


'''''''''''''''''''''''''
' Omschrijving property
'''''''''''''''''''''''''

Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property

Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property


'''''''''''''''''''''''''
' Zichtbaar property
'''''''''''''''''''''''''

Public Property Get Zichtbaar() As Boolean
    Zichtbaar = pZichtbaar
End Property

Public Property Let Zichtbaar(Value As Boolean)
    pZichtbaar = Value
End Property

'''''''''''''''''''''''''
' Kolomnummer property
'''''''''''''''''''''''''

Public Property Get Kolomnummer() As Long
    Kolomnummer = pKolomnummer
End Property

Public Property Let Kolomnummer(Value As Long)
    pKolomnummer = Value
End Property


Public Function insert() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim conn As New ADODB.connection
Dim strSQL As String
Dim db As New db
Dim maxid As Variant

'Instert Statement
strSQL = "INSERT INTO [KALENDER] (" & _
         "[Datum], " & _
         "[Feestdag], " & _
         "[Omschrijving], " & _
         "[Zichtbaar]) " & _
"VALUES (p1, p2, p3, p4);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("p1", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
    .Parameters.Append .CreateParameter("p2", adBoolean, adParamInput, Len(CStr(Me.Feestdag)), Me.Feestdag)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Zichtbaar)), Me.Zichtbaar)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

'Clean up the connection to the database
Set rst = Nothing
Set cmd = Nothing
db.Disconnect

Exit Function

Err_Insert:
Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
insert = False

End Function


Public Function update() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim db As New db
Dim con As ADODB.connection
Dim lRecordsAffected As Long

'update Statement
strSQL = "UPDATE [KALENDER] " & _
         "SET " & _
         "[Feestdag] = ?, " & _
         "[Omschrijving] = ?, " & _
         "[Zichtbaar] = ? " & _
         "WHERE [Datum] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adBoolean, adParamInput, Len(CStr(Me.Feestdag)), Me.Feestdag)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Zichtbaar)), Me.Zichtbaar)
    .Parameters.Append .CreateParameter("p1", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)

Set rst = .Execute(lRecordsAffected)
End With
    
If lRecordsAffected = 0 Then
    update = False
Else
    'Return a good return code
    update = True
End If

Exit_Insert:
'Clean up the connection to the database
Set rst = Nothing
Set cmd = Nothing
db.Disconnect

Exit Function
    
Err_Insert:
Call errorhandler_MsgBox("Class: " & TypeName(Me) & " , Function: Insert()")
update = False
Resume Exit_Insert
End Function


Function ToString() As String
Dim tekst As String
tekst = "Datum {"
tekst = tekst & vbNewLine & vbTab & "Datum = " & Me.datum
tekst = tekst & vbNewLine & vbTab & "Feestdag = " & Me.Feestdag
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Zichtbaar = " & Me.Zichtbaar
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function

Function FromList(r As Long, lijst As Variant)
    Me.datum = lijst(0, r)
    Me.Feestdag = lijst(1, r)
    If IsNull(lijst(2, r)) = False Then Me.Omschrijving = lijst(2, r)
    Me.Zichtbaar = lijst(3, r)
End Function

Function GetByDatum()
Dim db As New db
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM KALENDER WHERE Datum = p1;", pr)

    FromList 0, lijst

End Function
