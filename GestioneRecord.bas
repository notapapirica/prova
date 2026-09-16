Attribute VB_Name = "GestioneRecord"

Sub ApRiDB(nome As String, Optional Metodo = "DSN")
On Error GoTo Errore
  Set DB = New Connection
  Connessione = nome
  DB.ConnectionString = Connessione
  DB.Open
  DB.CommandTimeout = 600000 '10 MINUTI
Exit Sub
Errore:
  LetSeriale
  If DB.State <> 1 Then frmMDI.ControlloConnessione
  
End Sub
