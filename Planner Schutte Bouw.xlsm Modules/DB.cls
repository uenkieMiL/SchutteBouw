VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Compare Database
Const connection_string As String = "Provider = Microsoft.ACE.OLEDB.12.0; data source="
Const filename As String = "SchutteBouw.accdb"

Private WithEvents conn As ADODB.connection
Attribute conn.VB_VarHelpID = -1
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Public Sub Connect()
    If Not conn Is Nothing Then
        Exit Sub
    End If

    Set conn = New ADODB.connection
    conn.Open connection_string & dbbestandlocatie(filename)
End Sub

Public Function connection() As ADODB.connection
    If conn Is Nothing Then
        Set conn = New ADODB.connection
        conn.Open connection_string & dbbestandlocatie(filename)
    End If
    
    Set connection = conn
    End Function

Public Sub Disconnect()
    If Not conn Is Nothing Then
        If conn.State <> 0 Then
            conn.Close
        End If

        Set conn = Nothing
    End If
End Sub

Public Sub Execute(sql As String)
    If conn Is Nothing Then
        Exit Sub
    End If

    conn.Execute (sql)
End Sub

Public Function OpenRecordset(sql As String, Optional CursorLocation As ADODB.CursorLocationEnum = adUseClient, Optional CursorType As ADODB.CursorTypeEnum = adOpenForwardOnly, Optional LockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    If conn Is Nothing Then
        Exit Function
    End If

    If Not rs Is Nothing Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
            
    With rs
        .CursorLocation = CursorLocation
        .CursorType = CursorType
        .LockType = LockType
        .Open sql, conn
    End With

    Set OpenRecordset = rs
End Function

Public Sub BeginTransaction()
    conn.BeginTrans
End Sub

Public Sub CommitTransaction()
    conn.CommitTrans
End Sub

Public Sub RollbackTransaction()
    conn.RollbackTrans
End Sub

Private Sub conn_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.connection)
    Debug.Print "Transaction started."
End Sub

Private Sub conn_CommitTransComplete(ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.connection)
    Debug.Print "Transaction committed."
End Sub


Private Sub conn_RollbackTransComplete(ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.connection)
    Debug.Print "Transaction rolled back."
End Sub

Private Function dbbestandlocatie(databasebestand As String) As String
databasebestand = filename
dbbestandlocatie = ThisWorkbook.Path & "\" & databasebestand
End Function

Public Function getLijstBySQL(sql As String) As Variant
        
    Connect
    
    
    OpenRecordset (sql)

    If rs.EOF = False And rs.EOF = False Then
        getLijstBySQL = rs.GetRows
    End If
    
    Set rs = Nothing
    Disconnect
End Function

Function UpdateQueryUitvoeren(SQLQuery As String) As Long
    Me.connection.Execute SQLQuery, UpdateQueryUitvoeren
    
End Function

Public Function getLijstBySQL_AndParameter(sql As String, param As ADODB.Parameter) As Variant
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim aantal As Long
    Me.Connect

    With cmd
        .ActiveConnection = Me.connection
        .CommandType = adCmdText
        .CommandText = sql
    
        .Parameters.Append param
        
        Set rst = .Execute(aantal)
    End With
    If rst.EOF = False And rst.EOF = False Then
        getLijstBySQL_AndParameter = rst.GetRows
    End If
    
    Set rst = Nothing
    Me.Disconnect
End Function
