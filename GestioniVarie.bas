Attribute VB_Name = "GestioniVarie"
Option Explicit
'DICHIARAZIONI PER LA LETTURA DELLA RISOLUZIONE DELLO SCHERMO
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Const HORZRES As Long = 8
Const VERTRES As Long = 10
Const BITSPIXEL As Long = 12
Const VREFRESH As Long = 116
Const DESKTOPVERTRES As Long = 117   'Larghezza Orizzontale del intero desktop in pixels (NT5)
Const DESKTOPHORZRES As Long = 118   'Altezza Orizzontale del intero desktop in pixels (NT5)
Public DiplayResoluz As String

Global IsOpenVenditaDettaglio() As Boolean

Public Function INIRead2(File As String, Section As String, Item As String) As String
Dim sBuf As String
Dim iRC As Integer
Dim i As Long
For i = 0 To 1022
 sBuf = sBuf & " "
Next i
'sBuf = space(1023)
iRC = GetPrivateProfileString(Section, Item, "Not Found", sBuf, 1023, File)
sBuf = Left$(sBuf, iRC)

If StrComp(sBuf, "Not Found") = 0 Then
INIRead2 = vbNullString
Else
INIRead2 = sBuf
End If
End Function

Public Function LeggiDatiDaAxon(Tipo As String)
On Error GoTo Errore
Dim ritorno As String
Dim path As String
Dim doc As New MSXML2.DOMDocument60
Dim nodeList As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim success As Boolean
path = INIRead2(App.path + "\" + "Cassa.ini", "Parametri", "PathLeggi")
'success = doc.load("C:\axonFPiD_Pro_v7\Cartella_LOG\Response_StFisc.xml")
success = doc.Load(path)
Select Case UCase(Tipo)
        Case "CHIUSURA"
            If success = False Then
                MsgBox doc.parseError.reason
            Else
                Set nodeList = doc.selectNodes("/RESPONSE")
                If Not nodeList Is Nothing Then
                    For Each node In nodeList
                        ritorno = node.selectSingleNode("CMD_i_Z_NUMERO").Text
                    Next node
                End If
            End If
        Case "SCONTRINO"
            If success = False Then
                MsgBox doc.parseError.reason
            Else
                Set nodeList = doc.selectNodes("/RESPONSE")
                If Not nodeList Is Nothing Then
                    For Each node In nodeList
                        ritorno = node.selectSingleNode("CMD_9_NUMERO_SCONTRINO").Text
                    Next node
                End If
            End If
        Case "MATRICOLA"
            If success = False Then
                MsgBox doc.parseError.reason
            Else
                Set nodeList = doc.selectNodes("/RESPONSE")
                If Not nodeList Is Nothing Then
                    For Each node In nodeList
                        ritorno = node.selectSingleNode("CMD_a_ECR_MATRICOLA").Text
                    Next node
                End If
            End If
        Case "MODELLO"
            If success = False Then
                MsgBox doc.parseError.reason
            Else
                Set nodeList = doc.selectNodes("/RESPONSE")
                If Not nodeList Is Nothing Then
                    For Each node In nodeList
                        ritorno = node.selectSingleNode("CMD_a_ECR_CODICEMODELLO").Text
                    Next node
                End If
            End If
End Select
LeggiDatiDaAxon = ritorno
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "LeggiDatiDaAxon", Err.Number, Err.Description: Resume Next
End Function

Public Sub NascondiBarraMenu()
On Error GoTo Errore
'If Formisload("frmTotalePerFornitore") Then
  frmMDI.PctrMenu.Width = 435
  frmMDI.Picture3.Visible = False
  frmMDI.ARNascondi.Visible = False
  frmMDI.ARVisualizza.Visible = True
'End If
Exit Sub
Errore:
   GestioneErrori "GestioneVarie", "VisualizzaMenu", Err.Number, Err.Description: Resume Next
End Sub

Public Sub VisualizzaMenu()
On Error GoTo Errore
'If Formisload("frmTotalePerFornitore") Then
  frmMDI.PctrMenu.Width = 2430
  frmMDI.Picture3.Visible = True
  frmMDI.ARNascondi.Visible = True
  frmMDI.ARVisualizza.Visible = False
'End If
Exit Sub
Errore:
   GestioneErrori "GestioneVarie", "VisualizzaMenu", Err.Number, Err.Description: Resume Next
End Sub


Public Sub CaricaComboIndice(Cmb1 As ComboBox, Codice As Long)
On Error GoTo Errore
Dim i As Long
For i = 0 To Cmb1.ListCount - 1
 If Cmb1.ItemData(i) = Codice Then Cmb1.ListIndex = i: Exit For
Next i
Exit Sub
Errore:
  GestioneErrori "GestioneRecord", "CaricaComboIndice", Err.Number, Err.Description: Resume Next
End Sub

Public Function CollezioneDaProdotto(Prodotto As Long) As Long
Dim rs As Recordset
Set rs = New Recordset
rs.CursorLocation = adUseClient
Tabella rs, "SELECT prcollezione FROM Prodotti WHERE prcodice = " & numero(Prodotto)
If rs.RecordCount > 0 Then
 CollezioneDaProdotto = rs!prCollezione
Else
 CollezioneDaProdotto = 0
End If
rs.Close
Set rs = Nothing
End Function

Public Sub ApplicaScontoBenvenuto(Riga As Long)
On Error GoTo Errore
Dim cmd As Command
Set cmd = New Command
cmd.CommandTimeout = 3600000
cmd.ActiveConnection = Dbname
cmd.CommandType = adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("mdmid", adNumeric, adParamInput, , Riga)
cmd.Parameters("mdmid").Precision = 18
cmd.Parameters("mdmid").NumericScale = 0
cmd.CommandText = "ApplicaScontoBenvenuto"
cmd.Execute
Set cmd = Nothing
Exit Sub
Errore:
  GestioneErrori "GestioneVarie", "ApplicaScontoBenvenuto", Err.Number, Err.Description: Resume Next
End Sub

Public Sub InserisciEan(Prodotto As Long, gruppo As String, Taglia As String, Colore As String)
On Error GoTo Errore
Dim cmd As Command
Set cmd = New Command
cmd.CommandTimeout = 3600000
cmd.ActiveConnection = Dbname
cmd.CommandType = adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("Articolo", adNumeric, adParamInput, , Prodotto)
cmd.Parameters("Articolo").Precision = 18
cmd.Parameters("Articolo").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("Gruppo", adVarChar, adParamInput, 5, gruppo)
cmd.Parameters.Append cmd.CreateParameter("Taglia", adVarChar, adParamInput, 50, Taglia)
cmd.Parameters.Append cmd.CreateParameter("Colore", adVarChar, adParamInput, 5, Colore)
cmd.CommandText = "SP_InserisciEan"
cmd.Execute
Set cmd = Nothing
Exit Sub
Errore:
  GestioneErrori "GestioneVarie", "InserisciEan", Err.Number, Err.Description: Resume Next
End Sub

Public Function ControllaMovimentoTCCarico(Riga As Long) As Boolean
'Controllo la corrispondenza tra MovimentiMagazzino e Movimentimagazzinotc se è 1 a 1 torno true altrimenti false
On Error GoTo Errore
Dim rs As Recordset
Set rs = New Recordset
rs.CursorLocation = adUseClient
Tabella rs, "SELECT * FROM MovimentimagazzinotcCarico WHERE tcmdmid =" & numero(Riga)
If rs.RecordCount > 1 Then
 ControllaMovimentoTCCarico = False
Else
 ControllaMovimentoTCCarico = True
End If
rs.Close
Set rs = Nothing
Exit Function
Errore:
   GestioneErrori "GestioneVarie", "ControllaMovimentoTCCarico", Err.Number, Err.Description: Resume Next
End Function

Public Function ControllaMovimentoTC(Riga As Long) As Boolean
'Controllo la corrispondenza tra MovimentiMagazzino e Movimentimagazzinotc se è 1 a 1 torno true altrimenti false
On Error GoTo Errore
Dim rs As Recordset
Set rs = New Recordset
rs.CursorLocation = adUseClient
Tabella rs, "SELECT * FROM Movimentimagazzinotc WHERE tcmdmid =" & numero(Riga)
If rs.RecordCount > 1 Then
 ControllaMovimentoTC = False
Else
 ControllaMovimentoTC = True
End If
rs.Close
Set rs = Nothing
Exit Function
Errore:
   GestioneErrori "GestioneVarie", "ControllaMovimentoTC", Err.Number, Err.Description: Resume Next
End Function

Public Function ControllaMovimentoTCInterno(Riga As Long) As Boolean
'Controllo la corrispondenza tra MovimentiMagazzino e Movimentimagazzinotc se è 1 a 1 torno true altrimenti false
On Error GoTo Errore
Dim rs As Recordset
Set rs = New Recordset
rs.CursorLocation = adUseClient
Tabella rs, "SELECT * FROM MovimentiMagazzinoTCInterno WHERE tcmdmid =" & numero(Riga)
If rs.RecordCount = 1 Then
 ControllaMovimentoTCInterno = True
Else
 ControllaMovimentoTCInterno = False
End If
rs.Close
Set rs = Nothing
Exit Function
Errore:
   GestioneErrori "GestioneVarie", "ControllaMovimentoTCCarico", Err.Number, Err.Description: Resume Next
End Function

Public Sub SalvaLog(T1 As TickLog, Modulo As String, Procedura As String)
On Error GoTo Errore
Dim FileNumber, messaggio As String
  FileNumber = FreeFile
  Open App.path & "\Log\" & Replace(LCase(Modulo), "'", "") & ".log" For Append As #FileNumber
  Print #FileNumber, "*****************************************************************" & Chr(13)
  Print #FileNumber, "Data Log: " & CStr(Now) & Chr(13)
  Print #FileNumber, "Vendita Id: " & T1.IdVendita & Chr(13)
  Print #FileNumber, "Vendita Data: " & T1.DataVendita & Chr(13)
  Print #FileNumber, "Vendita Numero: " & T1.NumVendita & Chr(13)
  Print #FileNumber, "Nome della Maschera: " & Modulo & Chr(13)
  Print #FileNumber, "Nome Procedura: " & Procedura & Chr(13)
  Print #FileNumber, "Descrizione: " & T1.descrizione & Chr(13)
  Print #FileNumber, "Ricerca da Barcode: " & T1.BarcodeVendita & Chr(13)
  Print #FileNumber, "Reso da Scontrino Cortesia(Barcode): " & T1.ResoDaScCortesia & Chr(13)
  Print #FileNumber, "Ricerca da Riferimenti Vendita: " & T1.riferimentivendita & Chr(13)
  Print #FileNumber, "Tipo Operazione: " & T1.TipoOperazione & Chr(13)
  Print #FileNumber, "Totale Manuale: " & T1.TotaleManuale & Chr(13)
  Print #FileNumber, "Totale Controllato: " & T1.TotControllato & Chr(13)
  Print #FileNumber, "Utente: " & T1.Utente & Chr(13)
  Print #FileNumber, "*****************************************************************" & Chr(13)
  Close #FileNumber
Exit Sub
Errore:
  MkDir App.path & "\Log"
  Resume
End Sub


Public Sub CalcolaGiacenzeClient(Collezione As Long, Deposito As Long, DataFine As Date, Utente As String, Prodotto As Long)
Dim cmd As Command
Set cmd = New Command
cmd.CommandTimeout = 3600000
cmd.ActiveConnection = Dbname
cmd.CommandType = adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("ParColl", adNumeric, adParamInput, , Collezione)
cmd.Parameters("ParColl").Precision = 18
cmd.Parameters("ParColl").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("ParDeposito", adNumeric, adParamInput, , Deposito)
cmd.Parameters("ParDeposito").Precision = 18
cmd.Parameters("ParDeposito").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("DataFine", adDate, adParamInput, , DataFine)
cmd.Parameters.Append cmd.CreateParameter("Utente", adVarChar, adParamInput, 20, Utente)
cmd.Parameters.Append cmd.CreateParameter("ParProdotto", adNumeric, adParamInput, , Prodotto)
cmd.Parameters("ParProdotto").Precision = 18
cmd.Parameters("ParProdotto").NumericScale = 0
cmd.CommandText = "GiacenzeTaglieColori"
cmd.Execute
Set cmd = Nothing
End Sub

Public Sub CalcolaGiacenzeServer(Collezione As Long, Deposito As Long, DataFine As Date, Utente As String, Prodotto As Long)
Dim cmd As Command
Set cmd = New Command
cmd.CommandTimeout = 3600000
cmd.ActiveConnection = Dbname
cmd.CommandType = adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("ParColl", adNumeric, adParamInput, , Collezione)
cmd.Parameters("ParColl").Precision = 18
cmd.Parameters("ParColl").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("ParDeposito", adNumeric, adParamInput, , Deposito)
cmd.Parameters("ParDeposito").Precision = 18
cmd.Parameters("ParDeposito").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("DataFine", adDate, adParamInput, , DataFine)
cmd.Parameters.Append cmd.CreateParameter("Utente", adVarChar, adParamInput, 20, Utente)
cmd.Parameters.Append cmd.CreateParameter("ParProdotto", adNumeric, adParamInput, , Prodotto)
cmd.Parameters("ParProdotto").Precision = 18
cmd.Parameters("ParProdotto").NumericScale = 0
cmd.CommandText = "GiacenzeTaglieColori"
cmd.Execute
Set cmd = Nothing
End Sub
Public Sub CalcolaGiacenzeProdotto(Collezione As Long, Deposito As Long, DataFine As Date, Utente As String, Prodotto As Long, Inventario As Long)
Dim cmd As Command
Set cmd = New Command
cmd.CommandTimeout = 3600000
cmd.ActiveConnection = Dbname
cmd.CommandType = adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("ParColl", adNumeric, adParamInput, , Collezione)
cmd.Parameters("ParColl").Precision = 18
cmd.Parameters("ParColl").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("ParDeposito", adNumeric, adParamInput, , Deposito)
cmd.Parameters("ParDeposito").Precision = 18
cmd.Parameters("ParDeposito").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("DataFine", adDate, adParamInput, , DataFine)
cmd.Parameters.Append cmd.CreateParameter("Utente", adVarChar, adParamInput, 20, Utente)
cmd.Parameters.Append cmd.CreateParameter("ParProdotto", adNumeric, adParamInput, , Prodotto)
cmd.Parameters("ParProdotto").Precision = 18
cmd.Parameters("ParProdotto").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("ParInventario", adNumeric, adParamInput, , Inventario)
cmd.Parameters("ParInventario").Precision = 2
cmd.Parameters("ParInventario").NumericScale = 0
cmd.CommandText = "GiacenzeTCSingoloProdotto"
cmd.Execute
Set cmd = Nothing
End Sub
Public Function RisoluzioneSchermo(hdc As Long) As String
  On Error GoTo Errore
  
  Dim currHRes As Long
  Dim currVRes As Long
  Dim currBPP As Long
  Dim currVFreq As Long
  'Dim sBPPtype As String
  'Dim sFreqtype As String
  
  'get the system settings
  currHRes = GetDeviceCaps(hdc, HORZRES)
  currVRes = GetDeviceCaps(hdc, VERTRES)
  currBPP = GetDeviceCaps(hdc, BITSPIXEL)
  currVFreq = GetDeviceCaps(hdc, VREFRESH)
    
  'pretty up the descriptions a tad
'   Select Case currBPP
'      Case 4:      sBPPtype = "(16 Color)"
'      Case 8:      sBPPtype = "(256 Color)"
'      Case 16:     sBPPtype = "(High Color)"
'      Case 24, 32: sBPPtype = "(True Color)"
'   End Select
'
'   Select Case currVFreq
'      Case 0, 1:   sFreqtype = "(Hardware default)"
'      Case Else:   sFreqtype = "(User-selected)"
'   End Select
      
  RisoluzioneSchermo = Trim(currHRes) & "x" & Trim(currVRes)
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "RisoluzioneSchermo", Err.Number, Err.Description: Resume Next
End Function

'FUNZIONE UTILIZZATA PER CONTROLLARE SE GENERARE IL CODICE CLIENTE O FORNITORE IN AUTOMATICO
'MICHELE SCOGLIO DATA: 18/10/2005
Public Function StatoCodiceCliForAutomatico(Tipo As String) As String
On Error GoTo Errore
  Dim AnaAzi As Recordset
  Dim sql As String
  sql = "SELECT AziProgAutoCliente, AziProgAutoFornitore FROM AnaAzienda "
  Set AnaAzi = New Recordset
    Tabella AnaAzi, sql
    If Not AnaAzi.EOF Then
      If Tipo = "C" Then
        If Trim(AnaAzi!AziProgAutoCliente) = "S" Then
          StatoCodiceCliForAutomatico = "S"
        Else
          StatoCodiceCliForAutomatico = "N"
        End If
      Else
        If Trim(AnaAzi!AziProgAutoFornitore) = "S" Then
          StatoCodiceCliForAutomatico = "S"
        Else
          StatoCodiceCliForAutomatico = "N"
        End If
      End If
    Else
      StatoCodiceCliForAutomatico = "N"
    End If
  Set AnaAzi = Nothing
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "StatoCodiceCliForAutomatico", Err.Number, Err.Description: Resume Next
End Function

Function generaCodiceFiscale(Cog, Nom, Ses, Dat, Luo As String) As String
' PAGANO ANDREA
  Dim Cognome, nome, NomCon, NomVoc As String
  Dim pos, Cont As Long
  Const Vocale = "AEIOU"
  Const Consonanti = "BCDFGHJKLMNPQRSTVWXYZ"
  Const Mese = "ABCDEHLMPRST"
  Const Tabe = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Const TabA = "01020304050607080910111213141516171819202122232425"
  Const TabB = "0100050709131517192102041820110306081214161022252423"
  Const Spec1 = "ÀÈÉÌÒÙ '"
  Const Spec2 = "AEEIOU"
  'MI ASSICURO CHE I TESTI SIANO IN MAIUSCOLO
  Cog = UCase(Cog)
  Nom = UCase(Nom)
  'ELIMINAZIONE DI SPAZI, APOSTROFI E ACCENTI
  For Cont = 1 To Len(Cog)
    pos = InStr(Spec1, Mid(Cog, Cont, 1))
    If pos = 0 Then Cognome = Cognome + Mid(Cog, Cont, 1)
    If pos > 0 And pos < 7 Then Cognome = Cognome + Mid(Spec2, pos, 1)
  Next
  For Cont = 1 To Len(Nom)
    pos = InStr(Spec1, Mid(Nom, Cont, 1))
    If pos = 0 Then nome = nome + Mid(Nom, Cont, 1)
    If pos > 0 And pos < 7 Then nome = nome + Mid(Spec2, pos, 1)
  Next
  'CONVERSIONE DEL COGNOME
  For Cont = 1 To Len(Cognome)
    If InStr(Consonanti, Mid(Cognome, Cont, 1)) <> 0 Then generaCodiceFiscale = generaCodiceFiscale + Mid(Cognome, Cont, 1)
    If Len(generaCodiceFiscale) = 3 Then Exit For
  Next
  For Cont = 1 To Len(Cognome)
    If Len(generaCodiceFiscale) = 3 Then Exit For
    If InStr(Vocale, Mid(Cognome, Cont, 1)) <> 0 Then generaCodiceFiscale = generaCodiceFiscale + Mid(Cognome, Cont, 1)
  Next
  If Len(generaCodiceFiscale) < 3 Then
    For Cont = 1 To 3 - Len(generaCodiceFiscale)
      generaCodiceFiscale = generaCodiceFiscale + "X"
    Next
  End If
  'SEPARAZIONE DELLE CONSONANTI E DELLE VOCALI DEL NOME
  For Cont = 1 To Len(nome)
    If InStr(Consonanti, Mid(nome, Cont, 1)) <> 0 Then
      NomCon = NomCon + Mid(nome, Cont, 1)
    Else
      NomVoc = NomVoc + Mid(nome, Cont, 1)
    End If
  Next
  'CONVERSIONE DEL NOME
  If Len(NomCon) = 3 Then
    generaCodiceFiscale = generaCodiceFiscale + NomCon
  Else
    If Len(NomCon) > 3 Then
      generaCodiceFiscale = generaCodiceFiscale + Mid(NomCon, 1, 1) + Mid(NomCon, 3, 2)
    Else
      generaCodiceFiscale = generaCodiceFiscale + NomCon
      If Len(NomVoc) <> 0 Then
        For Cont = 1 To Len(NomVoc)
          generaCodiceFiscale = generaCodiceFiscale + Mid(NomVoc, Cont, 1)
          If Len(generaCodiceFiscale) = 6 Then Exit For
        Next
      End If
    End If
  End If
  If Len(generaCodiceFiscale) < 6 Then
    For Cont = 1 To 6 - Len(generaCodiceFiscale)
      generaCodiceFiscale = generaCodiceFiscale + "X"
    Next
  End If
  'CONVERSIONE DELLA DATA
  generaCodiceFiscale = generaCodiceFiscale + Right(Dat, 2)
  generaCodiceFiscale = generaCodiceFiscale + Mid(Mese, Mid(Dat, 4, 2), 1)
  If Ses = "F" Then
    generaCodiceFiscale = generaCodiceFiscale + Mid(Str(Val(Mid(Dat, 1, 2)) + 40), 2, 2)
  Else
    generaCodiceFiscale = generaCodiceFiscale + Mid(Dat, 1, 2)
  End If
  'CONVERSIONE DEL COMUNE
  generaCodiceFiscale = generaCodiceFiscale + Luo
End Function

Function generaChek(generaCodiceFiscale As String) As String
' PAGANO ANDREA
  Dim controllo, pos, Cont As Long
  Const Tabe = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Const TabB = "0100050709131517192102041820110306081214161022252423"
  'DETERMINAZIONE DELLA LETTERA DI CONTROLLO
  controllo = 0
  For Cont = 1 To 15
    pos = InStr(Tabe, Mid(generaCodiceFiscale, Cont, 1))
    If pos > 10 Then pos = pos - 10
    If Cont Mod 2 <> 0 Then
      controllo = controllo + Val(Mid(TabB, pos * 2 - 1, 2))
    Else
      controllo = controllo + pos - 1
    End If
  Next
  controllo = controllo Mod 26
  generaChek = Mid(Tabe, controllo + 11, 1)
End Function

Function estraiDataNascita(CodFisc As String) As String
' PAGANO ANDREA
  Const Mese = "ABCDEHLMPRST"
  If Mid(CodFisc, 10, 2) > 40 Then estraiDataNascita = Mid(CodFisc, 10, 2) - 40 Else estraiDataNascita = Mid(CodFisc, 10, 2)
  estraiDataNascita = estraiDataNascita & "/" & InStr(1, Mese, UCase(Mid(CodFisc, 9, 1)))
  estraiDataNascita = estraiDataNascita & "/" & Mid(CodFisc, 7, 2)
End Function

Function estraiSesso(CodFisc As String) As String
' PAGANO ANDREA
  If Mid(CodFisc, 10, 2) > 40 Then estraiSesso = "F" Else estraiSesso = "M"
End Function

'MICHELE 19/03/2007  - FUNZIONE RESA DINAMICA PER TABELLA E CAMPOCODICE
Function estraiLuogoNascita(CodFisc As String, StringaConnessione, Optional sTabella As String = "Comuni", Optional sCampoCodice As String = "cCodice", Optional sCampoNome As String = "cNome") As String
' PAGANO ANDREA
  Dim dbComuni As New Connection
  Dim tbComuni As New Recordset
  dbComuni.ConnectionString = StringaConnessione
  dbComuni.Open
    If dbComuni.State = adStateOpen Then
      tbComuni.Source = "SELECT " & Trim(sCampoNome) & " FROM " & Trim(sTabella) & " WHERE " & Trim(sCampoCodice) & "='" & Replace(Mid(CodFisc, 12, 4), "'", "''") & "'"
      tbComuni.ActiveConnection = dbComuni
      tbComuni.CursorType = adOpenKeyset
      tbComuni.LockType = adLockOptimistic
      tbComuni.Open
        If tbComuni.RecordCount <> 0 Then estraiLuogoNascita = tbComuni.Fields(Trim(sCampoNome))  'cNome
      tbComuni.Close
    End If
  dbComuni.Close
End Function

Function estraiProvinciaNascita(CodFisc As String, StringaConnessione) As String
' PAGANO ANDREA
  Dim dbComuni As New Connection
  Dim tbComuni As New Recordset
  dbComuni.ConnectionString = StringaConnessione
  dbComuni.Open
    If dbComuni.State = adStateOpen Then
      tbComuni.Source = "SELECT cProvincia FROM Comuni WHERE cCodice='" & Replace(Mid(CodFisc, 12, 4), "'", "''") & "'"
      tbComuni.ActiveConnection = dbComuni
      tbComuni.CursorType = adOpenKeyset
      tbComuni.LockType = adLockOptimistic
      tbComuni.Open
        If tbComuni.RecordCount <> 0 Then estraiProvinciaNascita = tbComuni!cProvincia
      tbComuni.Close
    End If
  dbComuni.Close
End Function

Function CntlCodiceFiscale(ByVal sz_Codice As String) As Long
' SORACI FRANCESCO
' Controllo il codice fiscale
  Const ALF1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Const CIF1 = "0123456789"
  Const ALF2 = "BAKPLCQDREVOSFTGUHMINJWZYX"
  Const CIF2 = "10   2 3 4   5 6 7 8 9    "
  Dim c_Char As String * 1
  Dim c_Code As String * 1
  Dim n_Count As Long
  Dim lcv As Long
  Dim space(16) As String
  sz_Codice = Left(UCase(sz_Codice) + space(16), 16)
  For lcv = 2 To 14 Step 2
    c_Char = Mid$(sz_Codice, lcv, 1)
    Select Case c_Char
      Case "0" To "9"
        n_Count = n_Count + InStr(CIF1, c_Char)
      Case "A" To "Z"
        n_Count = n_Count + InStr(ALF1, c_Char)
    End Select
  Next lcv
  For lcv = 1 To 15 Step 2
    c_Char = Mid$(sz_Codice, lcv, 1)
    Select Case c_Char
      Case "0" To "9"
        n_Count = n_Count + InStr(CIF2, c_Char)
      Case "A" To "Z"
        n_Count = n_Count + InStr(ALF2$, c_Char)
    End Select
  Next lcv
  n_Count = 1 + ((n_Count - 15) Mod 26) + 64
  c_Code = Chr$(n_Count)
  c_Char = Mid$(sz_Codice, 16, 1)
  If c_Code = c_Char Then
    CntlCodiceFiscale = 0
  Else
    CntlCodiceFiscale = n_Count
  End If
End Function

Function CntlPartitaIva(ByVal sz_Codice As String) As Long
' SORACI FRANCESCO
' Controllo la partita iva
  Dim n_Val As Long
  Dim n_Som1 As Long
  Dim n_Som2 As Long
  Dim lcv As Long
  For lcv = 1 To 9 Step 2
    n_Val = Val(Mid$(sz_Codice, lcv, 1))
    n_Som1 = n_Som1 + n_Val
    n_Val = Val(Mid$(sz_Codice, lcv + 1, 1))
    n_Som1 = n_Som1 + Int((n_Val * 2) / 10) + ((n_Val * 2) Mod 10)
  Next lcv
  n_Som2 = 10 - (n_Som1 Mod 10)
  If n_Som2 = 10 Then n_Som2 = 0
  n_Val = Val(Mid$(sz_Codice, 11, 1))
  If n_Som2 = n_Val Then
    CntlPartitaIva = 0
  Else
    CntlPartitaIva = n_Som2 + 48
  End If
End Function
Public Sub Gestionelog(Modulo As String, Procedura As String, numero, descrizione)
 On Error GoTo Errore
' PAGANO ANDREA
If numero = 5 Then Exit Sub
  Dim FileNumber, messaggio As String
  FileNumber = FreeFile
  messaggio = "Data dell'errore: " & CStr(Now) & Chr(13)
  messaggio = messaggio & "Nome della Maschera: " & Modulo & Chr(13)
  messaggio = messaggio & "Nome Procedura: " & Procedura & Chr(13)
  messaggio = messaggio & "Numero di Errore: " & numero & Chr(13)
  messaggio = messaggio & "Descrizione: " & descrizione & Chr(13)
  Open App.path & "\Log\" & Replace(LCase(Modulo), "'", "") & ".log" For Append As #FileNumber
    Print #FileNumber, "UserName___: " & UCase(strUtente)
    Print #FileNumber, "Data_Time__: " & Format(Date, "dd/mm/yyyy") & "_" & Format(Time, "hh:MM:ss")
    Print #FileNumber, "DbName_____: " & UCase(Dbname)
    Print #FileNumber, "Form_Sub___: " & UCase(Modulo) & "_" & UCase(Procedura)
    Print #FileNumber, "Error______: " & UCase(numero) & "_" & UCase(descrizione)
    Print #FileNumber, "--------------------------------------------------------------------------------"
  Close #FileNumber
   
Exit Sub
Errore:
  MkDir App.path & "\Log"
  Resume
End Sub
Public Sub GestioneErrori(Modulo As String, Procedura As String, numero, descrizione)
 On Error GoTo Errore
' PAGANO ANDREA
If numero = 5 Then Exit Sub
  Dim FileNumber, messaggio As String
  FileNumber = FreeFile
  messaggio = "Data dell'errore: " & CStr(Now) & Chr(13)
  messaggio = messaggio & "Nome della Maschera: " & Modulo & Chr(13)
  messaggio = messaggio & "Nome Procedura: " & Procedura & Chr(13)
  messaggio = messaggio & "Numero di Errore: " & numero & Chr(13)
  messaggio = messaggio & "Descrizione: " & descrizione & Chr(13)
  Open App.path & "\Log\" & Replace(LCase(Modulo), "'", "") & ".log" For Append As #FileNumber
    Print #FileNumber, "UserName___: " & UCase(strUtente)
    Print #FileNumber, "Data_Time__: " & Format(Date, "dd/mm/yyyy") & "_" & Format(Time, "hh:MM:ss")
    Print #FileNumber, "DbName_____: " & UCase(Dbname)
    Print #FileNumber, "Form_Sub___: " & UCase(Modulo) & "_" & UCase(Procedura)
    Print #FileNumber, "Error______: " & UCase(numero) & "_" & UCase(descrizione)
    Print #FileNumber, "--------------------------------------------------------------------------------"
  Close #FileNumber
  Select Case numero
    Case 68, 71, 3051
      MsgBox "Il Drive selezionato potrebbe non essere disponibile" & Chr$(13) & "oppure è vuoto.", vbInformation, "Attenzione!!! Contattare l'amministratore."
    Case 61
      MsgBox "Il disco potrebbe essere pieno o danneggiato.", vbInformation, "Attenzione!!! Contattare l'amministratore."
    Case 52
      'PERCORSO NON VALIDO
    Case 5
      'errori su set focus
    Case 20520
      ' Errore che viene generato quando rilancio lo stesso report mentre ancora il precedente è in elaborazione.
      ' Così facendo non viene più visualizzato il messaggio di crystal report
    Case Else
     MsgBox messaggio, vbCritical, "Attenzione!!! Contattare l'amministratore."
  End Select
Exit Sub
Errore:
  MkDir App.path & "\Log"
  Resume
End Sub

Sub CeNtRa(Maschera As Form, mWidth As Long, mHeight As Long)
' PAGANO ANDREA
On Error GoTo Errore
  Dim altezza As Long
  Maschera.WindowState = vbNormal
  Maschera.Width = mWidth
  Maschera.Height = mHeight
  Maschera.Left = 0
  Maschera.Top = 0
'  maschera.Left = frmMDI.Width / 2 - maschera.Width / 2 + 1
'  maschera.Top = frmMDI.Height / 2 - maschera.Height / 2 - 600  'PER CORREGGERE L'ALTEZZA
' IMPOSTO LA STESSA ICONA DELL'MDI
  Maschera.Icon = frmMDI.Icon
Exit Sub
Errore:
  GestioneErrori Maschera.Name, "Centra", Err.Number, Err.Description: Resume Next
End Sub

Sub BlOcCa(Maschera As Form, Optional Oggetto As String = "TUTTO", Optional stato As Boolean = True)
' PAGANO ANDREA
On Error GoTo Errore
  Dim Contatore As Long
  Dim controllo As Variant
  For Contatore = 0 To (Maschera.Controls.Count - 1)
    Set controllo = Maschera.Controls(Contatore)
    If UCase(controllo.Tag) <> "NO" Then
      Select Case UCase(Oggetto)
        Case "DTPICKER"
          'If TypeOf controllo Is DTPicker Then controllo.Enabled = Not (Stato)
        Case "FRAME"
          If TypeOf controllo Is Frame Then controllo.Enabled = Not (stato)
        Case "LABEL"
          If TypeOf controllo Is Label Then controllo.Enabled = Not (stato)
        Case "CHECKBOX"
          If TypeOf controllo Is CheckBox Then controllo.Enabled = Not (stato)
        Case "COMBOBOX"
          If TypeOf controllo Is ComboBox Then controllo.Enabled = Not (stato)
        Case "OPTIONBUTTON"
          If TypeOf controllo Is OptionButton Then controllo.Enabled = Not (stato)
        Case "TEXTBOX"
          If TypeOf controllo Is TextBox Then controllo.Enabled = Not (stato)
        Case "AMST"
          If TypeOf controllo Is AMST Then controllo.Enabled = Not (stato)
        Case "AMSR"
          If TypeOf controllo Is AMSR Then controllo.Enabled = Not (stato)
        Case "AMSDTP"
          If TypeOf controllo Is AMST Then controllo.Enabled = Not (stato)
        Case "TUTTO"
        '  If TypeOf controllo Is DTPicker Then controllo.Enabled = Not (Stato)
          If TypeOf controllo Is Frame Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is Label Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is CheckBox Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is ComboBox Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is OptionButton Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is TextBox Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is AMST Then controllo.Enabled = Not (stato)
          If TypeOf controllo Is AMSR Then controllo.Enabled = Not (stato)
          'If TypeOf controllo Is AMSDTP Then controllo.Enabled = Not (Stato)
      End Select
    End If
  Next Contatore
Exit Sub
Errore:
  GestioneErrori Maschera.Name, "Blocca", Err.Number, Err.Description: Resume Next
End Sub

Sub PuLiScI(Maschera As Form, Optional Oggetto = "TUTTO")
' PAGANO ANDREA, MARCELLINO DARIO
On Error GoTo Errore
  Dim Contatore As Long
  Dim controllo As Variant
  For Contatore = 0 To (Maschera.Controls.Count - 1)
    Set controllo = Maschera.Controls(Contatore)
    If UCase(controllo.Tag) <> "NO" Then
      Select Case UCase(Oggetto)
        Case "DTPICKER"
       '   If TypeOf controllo Is DTPicker Then controllo.Value = ""
        Case "LABEL"
          If TypeOf controllo Is Label And UCase(controllo.Tag) = "SI" Then controllo.Caption = ""
        Case "CHECKBOX"
          If TypeOf controllo Is CheckBox Then controllo.Value = 0
        Case "COMBOBOX"
           If TypeOf controllo Is ComboBox Then controllo.Text = ""
        '  If TypeOf controllo Is ComboBox Then controllo.ListIndex = -1: controllo.Text = ""
       
        Case "OPTIONBUTTON"
          If TypeOf controllo Is OptionButton Then controllo.Value = 0
        Case "TEXTBOX"
          If TypeOf controllo Is TextBox Then controllo.Text = ""
        Case "AMSR" 'ciccio
            If TypeOf controllo Is AMSR Then If controllo.Bloccato = False Then controllo.Valore = 0 'ciccio
            If TypeOf controllo Is AMSR Then If controllo.Bloccato = False Then controllo.Valore = "" 'ciccio
        Case "AMST" 'ciccio
            If TypeOf controllo Is AMST Then If controllo.Bloccato = False Then controllo.Text = "" 'ciccio
        Case "AMSDTP" 'ciccio
          If TypeOf controllo Is AMSDTP Then If controllo.Bloccato = False Then controllo.Valore = "" 'ciccio
        Case "TUTTO"
     '     If TypeOf controllo Is DTPicker Then controllo.Value = ""
          If TypeOf controllo Is Label And UCase(controllo.Tag) = "SI" Then controllo.Caption = ""
          If TypeOf controllo Is CheckBox Then controllo.Value = 0
          If TypeOf controllo Is ComboBox Then controllo.ListIndex = -1
          If TypeOf controllo Is OptionButton Then controllo.Value = 0
          If TypeOf controllo Is TextBox Then controllo.Text = ""
          If TypeOf controllo Is AMST Then
            If controllo.Bloccato = False Then controllo.Text = ""
          End If
          If TypeOf controllo Is AMSR Then
            If controllo.Bloccato = False Then controllo.Valore = 0
'            If controllo.Bloccato = False Then controllo.Valore = ""
          End If
          If TypeOf controllo Is AMSDTP Then
            If controllo.Bloccato = False Then controllo.Valore = ""
          End If
      End Select
    End If
  Next Contatore
Exit Sub
Errore:
  If Err.Number = 35787 Then Resume Next 'VUOL DIRE CHE IL DTPICKER NON HA IL CHEKBOX
  GestioneErrori Maschera.Name, "PULISCI", Err.Number, Err.Description: Resume Next
End Sub

Sub BloccaSblocca(Maschera As Form, stato As Long, Optional Oggetto = "TUTTO")
' PAGANO ANDREA, MARCELLINO DARIO
On Error GoTo Errore
  Dim Contatore As Long
  Dim controllo As Variant
  For Contatore = 0 To (Maschera.Controls.Count - 1)
    Set controllo = Maschera.Controls(Contatore)
    If stato = 0 Then
         If TypeOf controllo Is AMSDTP Then controllo.Locked = False: controllo.BackColor = &H80000005
         If TypeOf controllo Is ComboBox Then controllo.BackColor = &H80000005
         If TypeOf controllo Is TextBox Then controllo.Locked = False: controllo.BackColor = &H80000005
         If TypeOf controllo Is AMSR Then controllo.Locked = False: controllo.BackColor = &H80000005
         If TypeOf controllo Is AMST Then controllo.Locked = False: controllo.BackColor = &H80000005
'        Case "AMSDTP" 'ciccio
'          If TypeOf controllo Is AMSDTP Then If controllo.Bloccato = False Then controllo.Valore = "" 'ciccio
'        Case "TUTTO"
'     '     If TypeOf controllo Is DTPicker Then controllo.Value = ""
'          If TypeOf controllo Is Label And UCase(controllo.Tag) = "SI" Then controllo.Caption = ""
'          If TypeOf controllo Is CheckBox Then controllo.Value = 0
'          If TypeOf controllo Is ComboBox Then controllo.ListIndex = -1
'          If TypeOf controllo Is OptionButton Then controllo.Value = 0
'          If TypeOf controllo Is TextBox Then controllo.Text = ""
'          If TypeOf controllo Is amst Then
'            If controllo.Bloccato = False Then controllo.Text = ""
'          End If
'          If TypeOf controllo Is amsr Then
'            If controllo.Bloccato = False Then controllo.Valore = 0
'            If controllo.Bloccato = False Then controllo.Valore = ""
'          End If
'          If TypeOf controllo Is AMSDTP Then
'            If controllo.Bloccato = False Then controllo.Valore = ""
'          End If
    Else
     If TypeOf controllo Is AMSDTP Then controllo.Locked = True: controllo.BackColor = &H98EAF8
     If TypeOf controllo Is ComboBox Then controllo.BackColor = &H98EAF8
     If TypeOf controllo Is TextBox Then controllo.Locked = True: controllo.BackColor = &H98EAF8
     If TypeOf controllo Is AMSR Then controllo.Locked = True: controllo.BackColor = &H98EAF8
     If TypeOf controllo Is AMST Then controllo.Locked = True: controllo.BackColor = &H98EAF8
    End If
  Next Contatore
Exit Sub
Errore:
  If Err.Number = 35787 Then Resume Next 'VUOL DIRE CHE IL DTPICKER NON HA IL CHEKBOX
  GestioneErrori Maschera.Name, "PULISCI", Err.Number, Err.Description: Resume Next
End Sub

Public Function Testo(stringa, Optional Maiuscolo As Boolean = False)
' PAGANO ANDREA
On Error GoTo Errore
  If Trim(stringa) <> "" Then
    stringa = Replace(stringa, "'", "''")
    stringa = Trim(stringa)
    If Maiuscolo = True Then stringa = UCase(stringa)
  End If
  Testo = stringa
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "TESTO", Err.Number, Err.Description: Resume Next
End Function

Public Function numero(stringa)
' PAGANO ANDREA
On Error GoTo Errore
  If Trim(stringa) <> "" Then
    If VarType(stringa) = vbDouble Then
      Dim Stringa2 As String
      Stringa2 = stringa
      Stringa2 = Replace(Replace(Stringa2, ".", ""), ",", ".")
      numero = Stringa2
    Else
      stringa = Replace(Replace(stringa, ".", ""), ",", ".")
      numero = stringa
    End If
  Else
    numero = 0
  End If
  If Not IsNumeric(numero) Then numero = 0
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "Numero", Err.Number, Err.Description: Resume Next
End Function

Public Function Data(stringa)
On Error GoTo Errore
' PAGANO ANDREA
  Data = Format(stringa, "dd/mm/yyyy")
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "DATA", Err.Number, Err.Description: Resume Next
End Function

Public Function Ora(stringa)
' PAGANO ANDREA
On Error GoTo Errore
  Ora = Replace(Format(stringa, "HH.MM.ss"), ".", ":")
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "Ora", Err.Number, Err.Description: Resume Next
End Function

Public Function AggiungiZero(Valore As String, Lunghezza As Long) As String
' SORACI FRANCESCO
On Error GoTo Errore
  Dim ind As Long
  Dim campo As String
  Dim i As Long
  If Not IsNumeric(Valore) Or Valore = "" Then Exit Function
  campo = ""
  ind = Lunghezza - Len(Valore)
  For i = 1 To ind
    campo = campo & "0"
  Next i
  campo = campo & Valore
  AggiungiZero = campo
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "AggiungiZero", Err.Number, Err.Description: Resume Next
End Function


Public Function Text2Check(Testo) As Long
' PAGANO ANDREA
On Error GoTo Errore
  Select Case Trim(Testo)
    Case "S"
      Text2Check = 1
    Case Else
      Text2Check = 0
  End Select
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "Text2Check", Err.Number, Err.Description: Resume Next
End Function

Public Function Check2Text(check) As String
' PAGANO ANDREA
On Error GoTo Errore
  If UCase(check) = "VERO" Or UCase(check) = "TRUE" Then check = 1
  If UCase(check) = "FALSO" Or UCase(check) = "FALSE" Then check = 0
  Select Case check
    Case 1
      Check2Text = "S"
    Case Else
      Check2Text = "N"
  End Select
Exit Function
Errore:
  GestioneErrori "GestioneVarie", "Check2Text", Err.Number, Err.Description: Resume Next
End Function

Function Cripta_Pass(LogUte As String, Pass As String) As String
' PAGANO ANDREA
On Error GoTo Errore
  Dim PassC As String   'Componimento della Password Criptata
  Dim PassT             'Componimento della seconda parte della Password Criptata
  Dim Len_Ute           'Lunghezza del Login
  Dim Len_Pass          'Lunghezza della Password
  Dim i As Long      'Variabile contatore
  LogUte = "admin"
  PassC = ""
  Len_Ute = Len(LogUte)
  Len_Pass = Len(Pass)
  Select Case Len_Ute
    Case Is = Len_Pass
      For i = 1 To Len_Ute
        If Asc(Mid(LogUte, i, 1)) + Asc(Mid(Pass, i, 1)) <= 255 Then
          PassC = PassC & Chr(Asc(Mid(LogUte, i, 1)) + Asc(Mid(Pass, i, 1)))
        End If
      Next
    Case Is < Len_Pass
      For i = 1 To Len_Ute
        PassC = PassC & Chr(Asc(Mid(LogUte, i, 1)) + Asc(Mid(Pass, i, 1)))
      Next
      For i = Len_Ute To Len_Pass
        PassT = PassT & Asc(Mid(Pass, i, 1))
      Next
      While PassT > 255
      PassT = PassT / 255
      Wend
      PassC = PassC & Chr(PassT)
    Case Is > Len_Pass
      If Len_Pass > 0 Then
        For i = 1 To Len_Pass
          PassC = PassC & Chr(Asc(Mid(LogUte, i, 1)) + Asc(Mid(Pass, i, 1)))
        Next
        For i = Len_Pass To Len_Ute
          PassT = PassT + Asc(Mid(LogUte, i, 1))
        Next
        If PassT > 255 Then PassT = PassT / 255
        PassC = PassC & Chr(PassT)
      End If
  End Select
  If Len_Pass > 0 Then
    Cripta_Pass = PassC
  Else
    Cripta_Pass = ""
  End If
Exit Function
Errore:
  GestioneErrori "GestioniVarie", "Cripta_Pass", Err.Number, Err.Description: Resume Next
End Function


'R E G I S T R O  -  R E G I S T R O  -  R E G I S T R O  -  R E G I S T R O  -  R E G I S T R O  -  R E G I S T R O
'-------------------------------------------------------------------------------------------------------------------
Public Function isSZKeyExist(szKeyPath As String, szKeyName As String, ByRef szKeyValue As String) As Boolean
On Error GoTo Errore
'PAGANO ANDREA
  Dim bRes As Boolean
  Dim lRes As Long
  Dim hkey As Long
  lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyPath, 0&, KEY_QUERY_VALUE, hkey)
  If lRes <> ERROR_SUCCESS Then
    isSZKeyExist = False
    Exit Function
  End If
  lRes = RegQueryValueEx(hkey, szKeyName, 0&, REG_SZ, ByVal szKeyValue, Len(szKeyValue))
  RegCloseKey (hkey)
  If lRes <> ERROR_SUCCESS Then
    isSZKeyExist = False
    Exit Function
  End If
  isSZKeyExist = True
Exit Function
Errore:
  GestioneErrori "GestioniVarie", "isSZKeyExist", Err.Number, Err.Description: Resume Next
End Function

Public Function checkSqlDriver(ByRef szDriverName As String) As Boolean
'PAGANO ANDREA
  Dim szKeyPath As String
  Dim szKeyName As String
  Dim szKeyValue As String
  Dim bRes As Boolean
  bRes = False
  szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\SQL Server"
  szKeyName = "Driver"
  szKeyValue = String(255, Chr(32))
  If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
    szDriverName = szKeyValue
    bRes = True
  Else
    bRes = False
  End If
  checkSqlDriver = bRes
End Function

Public Function checkWantedSqlDSN(szWantedDSN As String) As Boolean
'PAGANO ANDREA
  Dim szKeyPath As String
  Dim szKeyName As String
  Dim szKeyValue As String
  Dim bRes As Boolean
  szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
  szKeyName = szWantedDSN
  szKeyValue = String(255, Chr(32))
  If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
      bRes = True
  Else
      bRes = False
  End If
  checkWantedSqlDSN = bRes
End Function

Public Function LetSeriale() As String
'PAGANO ANDREA
  Dim hkey As Long
  Dim lRes As Long
  Dim szKeyValue As String
  Dim Parametro As String
  szKeyValue = String(255, Chr(32))
  If isSZKeyExist("SOFTWARE\Ams Consulting s.r.l.\Ams@zienda", "Seriale", szKeyValue) Then
    LetSeriale = Left(szKeyValue, InStr(szKeyValue, Chr(0)) - 1)
  Else
    Parametro = "4791-4021-AERD-NA8O-NAGA-P9YB"
    If RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Ams Consulting s.r.l.\" & "Ams@zienda", hkey) <> ERROR_SUCCESS Then
    Else
      lRes = RegSetValueExString(hkey, "Seriale", 0&, REG_SZ, Parametro, Len(Parametro))
      lRes = RegCloseKey(hkey)
      LetSeriale = "4791-4021-AERD-NA8O-NAGA-P9YB"
    End If
  End If
End Function

Public Function LetBackup(NomeDB As String) As String
'PAGANO ANDREA
  Dim hkey As Long
  Dim lRes As Long
  Dim szKeyValue As String
  Dim Parametro As String
  szKeyValue = String(255, Chr(32))
  If isSZKeyExist("SOFTWARE\Ams Consulting s.r.l.\Ams@zienda", NomeDB, szKeyValue) Then
    LetBackup = Left(szKeyValue, InStr(szKeyValue, Chr(0)) - 1)
  End If
End Function

Public Sub GetBackup(NomeDB As String)
'PAGANO ANDREA
  Dim hkey As Long
  Dim lRes As Long
  Dim szKeyValue As String
  Dim Parametro As String
  szKeyValue = String(255, Chr(32))
  Parametro = "OK"
  If RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Ams Consulting s.r.l.\" & "Ams@zienda", hkey) = ERROR_SUCCESS Then
    lRes = RegSetValueExString(hkey, NomeDB, 0&, REG_SZ, Parametro, Len(Parametro))
    lRes = RegCloseKey(hkey)
  End If
End Sub

'NEW GESTIONE REGISTRO BY PAGANO ANDREA 04/11/2003
'--------------------------------------------------------------------------------------------------------------------
Public Sub RegWrite(Chiave As String, ValoreStringa As String, stringa As String)
On Error GoTo Errore
'PAGANO ANDREA
  Dim hkey As Long
  Dim lRes As Long
  Dim szKeyValue As String
  If Chiave <> "" Then
    Chiave = "SOFTWARE\Ams Consulting s.r.l.\" & App.Title & "\" & Chiave
  Else
    Chiave = "SOFTWARE\Ams Consulting s.r.l.\" & App.Title
  End If
  szKeyValue = String(255, Chr(32))
  If RegCreateKey(HKEY_LOCAL_MACHINE, Chiave, hkey) = ERROR_SUCCESS Then
    lRes = RegSetValueExString(hkey, ValoreStringa, 0&, REG_SZ, UCase(stringa), Len(stringa))
    lRes = RegCloseKey(hkey)
  End If
Exit Sub
Errore:
  GestioneErrori "GestioniVarie", "RegWrite", Err.Number, Err.Description: Resume Next
End Sub

Public Function RegRead(Chiave As String, ValoreStringa As String) As String
On Error GoTo Errore
'PAGANO ANDREA
  Dim szKeyValue As String
  If Chiave <> "" Then
    Chiave = "SOFTWARE\Ams Consulting s.r.l.\" & App.Title & "\" & Chiave
  Else
    Chiave = "SOFTWARE\Ams Consulting s.r.l.\" & App.Title
  End If
  szKeyValue = String(255, Chr(32))
  If isSZKeyExist(Chiave, ValoreStringa, szKeyValue) Then
    If Trim(szKeyValue) <> "" Then
      RegRead = UCase(Left(szKeyValue, InStr(szKeyValue, Chr(0)) - 1))
    Else
      RegRead = ""
    End If
  Else
    RegCreateKey HKEY_LOCAL_MACHINE, Chiave, 0
    RegRead = ""
  End If
Exit Function
Errore:
  GestioneErrori "GestioniVarie", "RegRead", Err.Number, Err.Description: Resume Next
End Function

'NEW GESTIONE DI LOG SU UTILITY DI MAGAZZINO E CONTABILITA' BY SCOGLIO MICHELE 21/02/2005
'-----------------------------------------------------------------------------------------------------------------
Public Sub InsertLog(strNomeDatabase As String, strNomeForm As String, strInfoChiave As String, strTipoOperazione As String)
On Error GoTo Errore
'SCOGLIO MICHELE - IL LOG SONO SITUATI NEL DBLOGIN
  Dim SqlLog As String
    SqlLog = "INSERT INTO Log (Log_DB,Log_NomeForm,Log_InfoChiave,Log_TipoOperazione,UserModifica,DataModifica) VALUES "
    SqlLog = SqlLog & "('" & Testo(strNomeDatabase) & "'"
    SqlLog = SqlLog & ",'" & Testo(strNomeForm) & "'"
    SqlLog = SqlLog & ",'" & Testo(strInfoChiave) & "'"
    SqlLog = SqlLog & ",'" & Testo(strTipoOperazione) & "'"
    SqlLog = SqlLog & ",'" & Testo(strUtente) & "'"
    SqlLog = SqlLog & ",'" & Data(Date) & "')"
    DBStart.Execute SqlLog
Exit Sub
Errore:
  GestioneErrori "GestioniVarie", "InsertLog", Err.Number, Err.Description: Resume Next
End Sub

Public Function RegDel(Chiave As String) As String
'PAGANO ANDREA
  Dim szKeyValue As String
  If Chiave <> "" Then
    Chiave = "SOFTWARE\Ams Consulting s.r.l.\" & App.Title & "\" & Chiave
  Else
    Chiave = "SOFTWARE\Ams Consulting s.r.l.\" & App.Title
  End If
 RegDel = RegDeleteKey(HKEY_LOCAL_MACHINE, Chiave)
End Function
Public Function UtenteConnesso(Utente As String) As Boolean 'CICCCIOLOCK - INIZIO
On Error GoTo Errore
Dim TbUtente As Recordset
Set TbUtente = New Recordset
UtenteConnesso = True
Tabellalogin TbUtente, "Select * from Utenti Where LogFlag=1 and LogUtente='" & Testo(Utente) & "' and LogNomedataBase='" & Trim(Dbname) & "'"
If TbUtente.RecordCount = 0 Then UtenteConnesso = False
Set TbUtente = Nothing
Exit Function
Errore:
  GestioneErrori "GestioniVarie", "UtenteConnesso", Err.Number, Err.Description: Resume Next
End Function 'CICCCIOLOCK - FINE

Public Function CaratteriNonValidi(Codice As String) As Boolean
On Error GoTo Errore
Dim toSearch
Dim Posz As Long
Dim i As Long
toSearch = Array(">", "<", "=", "'", """")
CaratteriNonValidi = False
For i = 0 To UBound(toSearch)
  Posz = InStr(Codice, toSearch(i))
  If Posz > 0 Then CaratteriNonValidi = True: Exit For
Next i
Exit Function
Errore:
  GestioneErrori "GestioniVarie", "CaratteriNonValidi", Err.Number, Err.Description: Resume Next
End Function
Public Sub ApriTastieraVirtuale(hnd As Long)
On Error GoTo Errore
'Const Top_const = 80
'Const left_const = 30
'Dim lngHWnd As Long
'Dim stringa As String
'Dim i As Long
'Dim topposition As Double
'Dim leftposition As Double
'Dim rec As RECT
'Dim FileNumber As String
'Dim height As Long
'Dim altezza As Long
'If Tastiera = 0 Then Exit Sub
'ChiudiTastieraVirtuale
'FileNumber = FreeFile
'
'If GetWindowRect(hnd, rec) = 1 Then
' altezza = 400
' For i = 0 To 8
'  If InStr(1, Dati(i), "TMainFormHeight=", vbTextCompare) <> 0 Then
'    altezza = CLng(Replace(Dati(i), "TMainFormHeight=", ""))
'    Exit For
'  End If
' Next
'   If Abs(rec.Top) + altezza > Screen.height Then
''   MsgBox "2-si"
'      topposition = Abs(rec.Top) - altezza
'   Else
''   MsgBox "2-no"
'      topposition = Abs(rec.Bottom)
'   End If
''   MsgBox "3"
'   leftposition = Abs(rec.Left)
'End If
''MsgBox "4"
'i = 0
'lngHWnd = FindWindow(vbNullString, "AMS Virtual Keyboard")
''MsgBox "5"
'If lngHWnd <> 0 Then
''MsgBox "5-si"
'  SendMessage lngHWnd, WM_CLOSE, 0&, 0&
'End If
''MsgBox "6"
'Kill App.Path & "\AMSVK\AMSVK.ini"
'While Len(Dir(App.Path & "\AMSVK\AMSVK.ini", vbNormal)) <> 0
'Wend
'For i = 0 To 8
'   If InStr(1, Dati(i), "TMainFormTop=", vbTextCompare) <> 0 Then
'      Dati(i) = "TMainFormTop=" & (topposition + Top_const)
'   End If
'   If InStr(1, Dati(i), "TMainFormLeft=", vbTextCompare) <> 0 Then
'      Dati(i) = "TMainFormLeft=" & (leftposition + left_const)
'   End If
'Next i
'Open App.Path & "\AMSVK\AMSVK.ini" For Output As #FileNumber
'    Print #FileNumber, Dati(0)
'    Print #FileNumber, Dati(1)
'    Print #FileNumber, Dati(2)
'    Print #FileNumber, Dati(3)
'    Print #FileNumber, Dati(4)
'    Print #FileNumber, Dati(5)
'    Print #FileNumber, Dati(6)
'    Print #FileNumber, Dati(7)
'    Print #FileNumber, Dati(8)
'Close #FileNumber
'Shell App.Path & "\AMSVK\AMSVK.exe"
Exit Sub
Errore:
  GestioneErrori "GestioniVarie", "TastieraVirtuale", Err.Number, Err.Description: Resume Next
End Sub

Public Function ChiudiTastieraVirtuale()
On Error GoTo Errore
If Tastiera = 0 Then Exit Function
Dim lngHWnd As Long
lngHWnd = FindWindow(vbNullString, "AMS Virtual Keyboard")
If lngHWnd <> 0 Then
  SendMessage lngHWnd, WM_CLOSE, 0&, 0&
End If
Exit Function
Errore:
  GestioneErrori "GestioniVarie", "TastieraVirtuale", Err.Number, Err.Description: Resume Next
End Function
Public Function Solo_Numeri(Key_Ascii As Integer) As Integer
    Solo_Numeri = Key_Ascii
    If Solo_Numeri = 8 Or Solo_Numeri = 9 Or Solo_Numeri = 13 Or Solo_Numeri = 44 Or Solo_Numeri = 45 Then Exit Function
    If Solo_Numeri = 46 Then
       Solo_Numeri = 44
       Exit Function
    End If
    If Solo_Numeri < 48 Or Solo_Numeri > 57 Then Solo_Numeri = 0
End Function
Sub settaTastiera(nome As String)
On Error GoTo Errore
Dim dbtemp As Connection
Dim tbrecord As Recordset

Set dbtemp = New Connection
Set tbrecord = New Recordset
dbtemp.ConnectionString = nome
dbtemp.Open
If dbtemp.State = adStateOpen Then
  tbrecord.Open "Select azitastiera,azivalidate from anaazienda ", dbtemp, adOpenKeyset
  If tbrecord.RecordCount <> 0 Then
     Tastiera = tbrecord!aziTastiera
     Validate = tbrecord!azivalidate
 Else
     Tastiera = 1
     Validate = 1
 End If
 tbrecord.Close
 Set tbrecord = Nothing
Else
  Tastiera = 1
   Validate = 1
End If
dbtemp.Close
Set dbtemp = Nothing
Exit Sub
Errore:
  GestioneErrori "settaTastiera", "GestioneRecord", Err.Number, Err.Description: Resume Next
End Sub
Public Function DAzi(nome As String) As String
On Error GoTo Errore
Dim dbtemp As Connection
Dim tbrecord As Recordset
DAzi = ""
Set dbtemp = New Connection
Set tbrecord = New Recordset
dbtemp.ConnectionString = nome
dbtemp.Open
If dbtemp.State = adStateOpen Then
  tbrecord.Open "Select * from anaazienda ", dbtemp, adOpenKeyset
  If tbrecord.RecordCount <> 0 Then
     DAzi = tbrecord!aziRagSociale & vbCrLf
     If tbrecord!aziRagSocialeEstesa <> "" Then DAzi = DAzi & tbrecord!aziRagSocialeEstesa & vbCrLf
     DAzi = DAzi & tbrecord!aziIndirizzo & ", " & tbrecord!aziCitta & " " & tbrecord!aziCap & " " & UCase(tbrecord!aziProvincia)
     
 Else
     Tastiera = 1
       Validate = 1
 End If
 tbrecord.Close
 Set tbrecord = Nothing
Else
  Tastiera = 1
    Validate = 1
End If
dbtemp.Close
Set dbtemp = Nothing
Exit Function
Errore:
  GestioneErrori "settaTastiera", "GestioneRecord", Err.Number, Err.Description: Resume Next
End Function
Public Function IsRunning(ByVal proc As String) As Boolean
    Dim oProcs As Object
    Set oProcs = GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & proc & "'")
    IsRunning = (oProcs.Count > 0)
    Set oProcs = Nothing
End Function
Public Function Delcliente(Cod As Long) As Boolean
On Error GoTo Errore
Dim tbrecord As Recordset
Set tbrecord = New Recordset
Delcliente = False
Tabella tbrecord, "Select *  from documentotestata where dctCodiceCliente=" & Cod
If tbrecord.RecordCount <> 0 Then
  If MsgBox("Attenzione il cliente è movimentato, Continuare?", vbYesNo, "Atelier") = vbNo Then
    Delcliente = False
    Exit Function
  End If
End If
Delcliente = True


Exit Function
Errore:
  GestioneErrori "Delcliente", "GestioneRecord", Err.Number, Err.Description: Resume Next

End Function

Public Sub Delay(Intervallo As Long)
'PAGANO ANDREA
  Dim TEMPO1
  Dim TEMPO2
  TEMPO1 = Timer
  While Not TEMPO2 - TEMPO1 > Intervallo
    TEMPO2 = Timer
  Wend
End Sub

Public Function SetDefaultPrinter(objPrn As Printer) As Boolean
On Error GoTo Errore
    Dim X As Long, sztemp As String
    sztemp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    X = WriteProfileString("windows", "device", sztemp)
    'X = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
Exit Function
Errore:
 GestioneErrori "GestioniVarie", "SetDefaultPrinter", Err.Number, Err.Description: Resume Next
End Function

Public Function Leggistampante() As String
On Error GoTo Errore
Dim f As Integer
f = FreeFile
Leggistampante = ""
Open App.path & "\stampa.ini" For Input As #f
If Not EOF(f) Then
 Line Input #f, Leggistampante
End If
Close f
Exit Function
Errore:
 GestioneErrori "GestioniVarie", "Leggistampante", Err.Number, Err.Description: Resume Next
End Function


Public Sub InsertPremiStorico(ByVal Acfcod As Long, ByVal IdPremio As Long, ByVal dctid As Long, ByVal User As String)
On Error GoTo Errore
Dim cmd As Command
Set cmd = New Command
cmd.CommandTimeout = 3600000
cmd.ActiveConnection = Dbname
cmd.CommandType = adCmdStoredProc
cmd.Parameters.Append cmd.CreateParameter("AcfCod", adNumeric, adParamInput, , Acfcod)
cmd.Parameters("AcfCod").Precision = 18
cmd.Parameters("AcfCod").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("IdPremio", adNumeric, adParamInput, , IdPremio)
cmd.Parameters("IdPremio").Precision = 18
cmd.Parameters("IdPremio").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("Dctid", adNumeric, adParamInput, , dctid)
cmd.Parameters("Dctid").Precision = 18
cmd.Parameters("Dctid").NumericScale = 0
cmd.Parameters.Append cmd.CreateParameter("User", adVarChar, adParamInput, 10, User)
cmd.CommandText = "InsertPremiStorico"
cmd.Execute
Set cmd = Nothing
Exit Sub
Errore:
 GestioneErrori "GestioniVarie", "InsertPremiStorico", Err.Number, Err.Description: Resume Next
End Sub

