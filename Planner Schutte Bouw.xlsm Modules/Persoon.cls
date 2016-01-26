VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Persoon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pAfkorting As String
Private pVoornaam As String
Private pTussenvoegsel As String
Private pAchternaam As String
Private pEmailadres As String
Private pPercentage As Double
Private pIspm As Boolean
Private pIscalc As Boolean
Private pIsbim As Boolean
Private pIspl As Boolean
Private pIswvb As Boolean
Private pIsuitv As Boolean




'''''''''''''''''''''''''
' afkorting property
'''''''''''''''''''''''''

Public Property Get afkorting() As String
    afkorting = pAfkorting
End Property

Public Property Let afkorting(Value As String)
    pAfkorting = Value
End Property


'''''''''''''''''''''''''
' voornaam property
'''''''''''''''''''''''''

Public Property Get voornaam() As String
    voornaam = pVoornaam
End Property

Public Property Let voornaam(Value As String)
    pVoornaam = Value
End Property


'''''''''''''''''''''''''
' tussenvoegsel property
'''''''''''''''''''''''''

Public Property Get tussenvoegsel() As String
    tussenvoegsel = pTussenvoegsel
End Property

Public Property Let tussenvoegsel(Value As String)
    pTussenvoegsel = Value
End Property


'''''''''''''''''''''''''
' achternaam property
'''''''''''''''''''''''''

Public Property Get achternaam() As String
    achternaam = pAchternaam
End Property

Public Property Let achternaam(Value As String)
    pAchternaam = Value
End Property


'''''''''''''''''''''''''
' emailadres property
'''''''''''''''''''''''''

Public Property Get emailadres() As String
    emailadres = pEmailadres
End Property

Public Property Let emailadres(Value As String)
    pEmailadres = Value
End Property


'''''''''''''''''''''''''
' percentage property
'''''''''''''''''''''''''

Public Property Get percentage() As Double
    percentage = pPercentage
End Property

Public Property Let percentage(Value As Double)
    pPercentage = Value
End Property


'''''''''''''''''''''''''
' isPM property
'''''''''''''''''''''''''

Public Property Get isPM() As Boolean
    isPM = pIspm
End Property

Public Property Let isPM(Value As Boolean)
    pIspm = Value
End Property


'''''''''''''''''''''''''
' isCalc property
'''''''''''''''''''''''''

Public Property Get isCalc() As Boolean
    isCalc = pIscalc
End Property

Public Property Let isCalc(Value As Boolean)
    pIscalc = Value
End Property


'''''''''''''''''''''''''
' isBIM property
'''''''''''''''''''''''''

Public Property Get isBIM() As Boolean
    isBIM = pIsbim
End Property

Public Property Let isBIM(Value As Boolean)
    pIsbim = Value
End Property


'''''''''''''''''''''''''
' isPL property
'''''''''''''''''''''''''

Public Property Get isPL() As Boolean
    isPL = pIspl
End Property

Public Property Let isPL(Value As Boolean)
    pIspl = Value
End Property


'''''''''''''''''''''''''
' isWVB property
'''''''''''''''''''''''''

Public Property Get isWVB() As Boolean
    isWVB = pIswvb
End Property

Public Property Let isWVB(Value As Boolean)
    pIswvb = Value
End Property


'''''''''''''''''''''''''
' isUitv property
'''''''''''''''''''''''''

Public Property Get isUitv() As Boolean
    isUitv = pIsuitv
End Property

Public Property Let isUitv(Value As Boolean)
    pIsuitv = Value
End Property

Public Function insert() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim conn As New ADODB.connection
Dim strSQL As String
Dim db As New db
Dim maxid As Variant

'Instert Statement
strSQL = "INSERT INTO [PERSONEN] (" & _
         "[afkorting], " & _
         "[voornaam], " & _
         "[tussenvoegsel], " & _
         "[achternaam], " & _
         "[emailadres], " & _
         "[percentage], " & _
         "[isPM], " & _
         "[isCalc], " & _
         "[isBIM], " & _
         "[isPL], " & _
         "[isWVB], " & _
         "[isUitv]) " & _
"VALUES (p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.afkorting) + 1, Me.afkorting)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.voornaam) + 1, Me.voornaam)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.tussenvoegsel) + 1, Me.tussenvoegsel)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.achternaam) + 1, Me.achternaam)
    .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.emailadres) + 1, Me.emailadres)
    .Parameters.Append .CreateParameter("p6", adDouble, adParamInput, Len(CStr(Me.percentage)), Me.percentage)
    .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.isPM)), Me.isPM)
    .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.isCalc)), Me.isCalc)
    .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, Len(CStr(Me.isBIM)), Me.isBIM)
    .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.isPL)), Me.isPL)
    .Parameters.Append .CreateParameter("p11", adBoolean, adParamInput, Len(CStr(Me.isWVB)), Me.isWVB)
    .Parameters.Append .CreateParameter("p12", adBoolean, adParamInput, Len(CStr(Me.isUitv)), Me.isUitv)
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
Resume 'Exit_Insert

End Function

Public Function update() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim db As New db
Dim con As ADODB.connection
Dim lRecordsAffected As Long

'update Statement
strSQL = "UPDATE [PERSONEN] " & _
         "SET " & _
         "[voornaam] = ?, " & _
         "[tussenvoegsel] = ?, " & _
         "[achternaam] = ?, " & _
         "[emailadres] = ?, " & _
         "[percentage] = ?, " & _
         "[isPM] = ?, " & _
         "[isCalc] = ?, " & _
         "[isBIM] = ?, " & _
         "[isPL] = ?, " & _
         "[isWVB] = ?, " & _
         "[isUitv] = ? " & _
         "WHERE [afkorting] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.voornaam) + 1, Me.voornaam)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.tussenvoegsel) + 1, Me.tussenvoegsel)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.achternaam) + 1, Me.achternaam)
    .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.emailadres) + 1, Me.emailadres)
    .Parameters.Append .CreateParameter("p6", adDouble, adParamInput, Len(CStr(Me.percentage)), Me.percentage)
    .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.isPM)), Me.isPM)
    .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.isCalc)), Me.isCalc)
    .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, Len(CStr(Me.isBIM)), Me.isBIM)
    .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.isPL)), Me.isPL)
    .Parameters.Append .CreateParameter("p11", adBoolean, adParamInput, Len(CStr(Me.isWVB)), Me.isWVB)
    .Parameters.Append .CreateParameter("p12", adBoolean, adParamInput, Len(CStr(Me.isUitv)), Me.isUitv)
    .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.afkorting) + 1, Me.afkorting)

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
Resume Err_Insert
End Function

Public Function delete() As Boolean
On Error GoTo ErrorHandler
Dim sql As String
Dim rst As ADODB.Recordset
Dim db As New db

sql = "SELECT * FROM PERSONEN WHERE afkorting = " & Me.afkorting & ";"
'We are using a select statement that will return only
'one record (TeacherID 7)

db.Connect

Set rst = New ADODB.Recordset
rst.Open sql, db.connection, adOpenDynamic, adLockOptimistic
'Open RecordSet

With rst

If Not .BOF And Not .EOF Then
    'Ensure that the recordset contains records
    'If no records the code inside the if...end if
    'statement won't run
    
    .MoveLast
    .MoveFirst
    'Not necessary but good practice
    
    If .Supports(adDelete) Then
        'It is possible that the record you want to update
        'is locked by another user. If we don't check before
        'updating, we will generate an error
        
        .delete
        'The only command we need!
        delete = True
    End If
End If

.Close
'Make sure you close the recordset...
End With

db.Disconnect

ExitSub:
Set rst = Nothing
Exit Function
ErrorHandler:
Resume ExitSub

End Function
Function FromList(r As Long, lijst As Variant)
    Me.afkorting = lijst(0, r)
    Me.voornaam = lijst(1, r)
    If IsNull(lijst(2, r)) = False Then Me.tussenvoegsel = lijst(2, r)
    Me.achternaam = lijst(3, r)
    Me.emailadres = lijst(4, r)
    Me.percentage = lijst(5, r)
    Me.isPM = lijst(6, r)
    Me.isCalc = lijst(7, r)
    Me.isBIM = lijst(8, r)
    Me.isPL = lijst(9, r)
    Me.isWVB = lijst(10, r)
    Me.isUitv = lijst(11, r)
End Function

Function ToString() As String
Dim tekst As String
tekst = "PERSOON {"
tekst = tekst & vbNewLine & vbTab & "afkorting = " & Me.afkorting
tekst = tekst & vbNewLine & vbTab & "voornaam = " & Me.voornaam
tekst = tekst & vbNewLine & vbTab & "tussenvoegsel = " & Me.tussenvoegsel
tekst = tekst & vbNewLine & vbTab & "achternaam = " & Me.achternaam
tekst = tekst & vbNewLine & vbTab & "emailadres = " & Me.emailadres
tekst = tekst & vbNewLine & vbTab & "percentage = " & Me.percentage
tekst = tekst & vbNewLine & vbTab & "isPM = " & Me.isPM
tekst = tekst & vbNewLine & vbTab & "isCalc = " & Me.isCalc
tekst = tekst & vbNewLine & vbTab & "isBIM = " & Me.isBIM
tekst = tekst & vbNewLine & vbTab & "isPL = " & Me.isPL
tekst = tekst & vbNewLine & vbTab & "isWVB = " & Me.isWVB
tekst = tekst & vbNewLine & vbTab & "isUitv = " & Me.isUitv
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetByafkorting()
Dim db As New db
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adVarChar, adParamInput, Len(Me.afkorting) + 1, Me.afkorting)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM PERSONEN WHERE afkorting = p1;", pr)

    FromList 0, lijst

End Function


