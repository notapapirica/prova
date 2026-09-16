VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{759B32CF-6E3A-4181-A4B4-AB5BFB008A77}#1.1#0"; "AMSTextBox.ocx"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form FrmClientiTel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ClientiTel"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAggiorna 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7320
      Picture         =   "FrmClientiTel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin AMSTB.AMST AMSRagSociale 
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   50
      SelLenght       =   0
   End
   Begin AMSTB.AMST AMST1 
      Height          =   450
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   50
      SelLenght       =   0
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   7290
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   12765
      _cx             =   22516
      _cy             =   12859
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483625
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmClientiTel.frx":0375
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin VB.TextBox LblPosition2 
         Height          =   345
         Left            =   540
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   990
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rag. Sociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   -840
         Visible         =   0   'False
         Width           =   1830
      End
   End
   Begin ARButtonCtrl.ARButton ARFiltri 
      Height          =   1035
      Left            =   8520
      TabIndex        =   8
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1826
      Caption         =   "4"
      ForeColor       =   -2147483640
      BackColorOnMouse=   16777215
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BorderStyle     =   1
      Picture         =   "FrmClientiTel.frx":0431
   End
   Begin ARButtonCtrl.ARButton ARCancella 
      Height          =   1035
      Left            =   9960
      TabIndex        =   11
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1826
      Caption         =   "4"
      ForeColor       =   -2147483640
      BackColorOnMouse=   16777215
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BorderStyle     =   1
      Picture         =   "FrmClientiTel.frx":109D
   End
   Begin ARButtonCtrl.ARButton ARSalva 
      Height          =   975
      Left            =   11760
      TabIndex        =   13
      Top             =   0
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1720
      Caption         =   "4"
      ForeColor       =   -2147483640
      BackColorOnMouse=   16777215
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VerticalAlignment=   1
      Style           =   1
      BorderStyle     =   3
      Picture         =   "FrmClientiTel.frx":2513
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Salva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   12000
      TabIndex        =   14
      Top             =   1080
      Width           =   1830
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancella"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   10200
      TabIndex        =   12
      Top             =   1080
      Width           =   1830
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pulisci Filtri"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   10
      Top             =   1080
      Width           =   1830
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7800
      TabIndex        =   9
      Top             =   720
      Width           =   1830
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   1830
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rag. Sociale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1830
   End
End
Attribute VB_Name = "FrmClientiTel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Roberta Sciarrone 04/09/2026 inizio
Option Explicit
Dim sql As String
Dim m_DS As FlexADO
Public IDForm As Long
Dim m_CodiceClienteInModifica As String

Private Sub Form_Load()
On Error GoTo Errore

Me.Width = Screen.Width * 0.7
Me.Height = Screen.Height * 0.7

Connessione = "Provider=SQLNCLI11;Server=(local);Database=AtelierSRVAnsal;Uid=amsatelier; Pwd=amsatelier;"
ApRiDB Connessione

Set m_DS = New FlexADO
m_DS.Dbname = "AtelierSRVAnsal"
'Form_Resize
InizializzaGrigliaMovimenti
Selezione
Exit Sub
Errore:
    GestioneErrori Me.Name, "Form_Load", Err.Number, Err.Description: Resume Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
fg.Width = Me.ScaleWidth - (fg.Left * 2)
fg.Height = Me.ScaleHeight - fg.Top - 400
End Sub

Private Sub Selezione()
On Error GoTo Errore
Dim sqlLocal As String
Dim rsClienti As New ADODB.Recordset

sqlLocal = "SELECT ACFCod,ACFRagSoc,ACFindi,ACFCap,ACFCitta,ACFProv, " & _
" CAST(ACFTel1 AS VARCHAR(20)) AS ACFTel1, " & _
" CAST(ACFTel2 AS VARCHAR(20)) AS ACFTel2 " & _
" FROM Clienti " & _
" WHERE (ISNULL(ACFTel1, '') <> 0)"

If Trim(AMSRagSociale.Text) <> "" Then
    sqlLocal = sqlLocal & " AND ACFRagSoc LIKE '%" & Replace(AMSRagSociale.Text, "'", "''") & "%' "
End If

If Trim(AMST1.Text) <> "" Then
    sqlLocal = sqlLocal & "AND (CAST(ACFTel1 AS VARCHAR(20))LIKE '%" & Replace(AMST1.Text, "'", "''") & "%' " & _
    " OR CAST(ACFTel2 AS VARCHAR(20)) LIKE '%" & Replace(AMST1.Text, "'", "''") & "%') "
End If

sqlLocal = sqlLocal & " ORDER BY ACFRagSoc"

   rsClienti.Open sqlLocal, DB, adOpenStatic, adLockReadOnly
    
    fg.Redraw = False
    fg.Rows = 1
    
    Do While Not rsClienti.EOF
        fg.AddItem rsClienti("ACFRagSoc") & vbTab & _
                   rsClienti("ACFIndi") & vbTab & _
                   rsClienti("ACFCod") & vbTab & _
                   rsClienti("ACFCap") & vbTab & _
                   UCase(rsClienti("ACFCitta") & "") & vbTab & _
                   rsClienti("ACFProv") & vbTab & _
                   rsClienti("ACFTel1")
        rsClienti.MoveNext
    Loop
    
    rsClienti.Close
    
    If fg.Rows > 1 Then
        fg.Cell(flexcpBackColor, 1, 0, fg.Rows - 1, 0) = &HC0FFFF
        fg.Cell(flexcpBackColor, 1, 1, fg.Rows - 1, 1) = &HF8F2E9
        fg.Cell(flexcpBackColor, 1, 2, fg.Rows - 1, 2) = &HE7F8EC
        fg.Cell(flexcpBackColor, 1, 3, fg.Rows - 1, 3) = &HC0E2FC
        fg.Cell(flexcpBackColor, 1, 4, fg.Rows - 1, 4) = &HFFF2FF
        fg.Cell(flexcpBackColor, 1, 5, fg.Rows - 1, 5) = &HE7F1EC
        fg.Cell(flexcpBackColor, 1, 6, fg.Rows - 1, 6) = &HFFC0C0
        fg.MergeCells = flexMergeFree
        fg.MergeCol(0) = True
        fg.MergeCol(3) = True
        fg.MergeCol(4) = True
        fg.Cell(flexcpFontBold, 1, 4, fg.Rows - 1, 4) = True
    End If
    
    fg.Redraw = True
    Exit Sub

Errore:
    fg.Redraw = True
    GestioneErrori Me.Name, "Selezione", Err.Number, Err.Description: Resume Next
End Sub

Private Sub CmdAggiorna_Click()
Selezione
End Sub

Private Sub InizializzaGrigliaMovimenti()
On Error GoTo Errore
fg.AutoResize = True
fg.Rows = 1

fg.ColWidth(0) = 3500
fg.ColAlignment(0) = flexAlignCenterCenter
fg.FixedAlignment(0) = flexAlignCenterCenter
fg.TextMatrix(0, 0) = "Rag. Sociale"
fg.ColWidth(1) = 4000
fg.ColAlignment(1) = flexAlignLeftCenter
fg.FixedAlignment(1) = flexAlignLeftCenter
fg.TextMatrix(0, 1) = "Indirizzo"
fg.ColWidth(2) = 0
fg.ColAlignment(2) = flexAlignCenterCenter
fg.FixedAlignment(2) = flexAlignLeftCenter
fg.TextMatrix(0, 2) = "ACFCod"
fg.ColHidden(2) = True
fg.ColWidth(3) = 2500
fg.ColAlignment(3) = flexAlignCenterCenter
fg.FixedAlignment(3) = flexAlignLeftCenter
fg.TextMatrix(0, 3) = "CAP"
fg.ColWidth(4) = 3000
fg.ColAlignment(4) = flexAlignCenterCenter
fg.FixedAlignment(4) = flexAlignLeftCenter
fg.TextMatrix(0, 4) = "Città"
fg.ColWidth(5) = 3300
fg.ColAlignment(5) = flexAlignCenterCenter
fg.FixedAlignment(5) = flexAlignLeftCenter
fg.TextMatrix(0, 5) = "Provincia"
fg.ColWidth(6) = 2000
fg.ColAlignment(6) = flexAlignCenterCenter
fg.FixedAlignment(6) = flexAlignCenterCenter
fg.TextMatrix(0, 6) = "Telefono"

Exit Sub
Errore:
  GestioneErrori Me.Name, "InizializzaGrigliaMovimenti", Err.Number, Err.Description: Resume Next
End Sub

Private Sub fg_DblClick()
On Error GoTo Errore
If fg.Row <> 0 Then
    m_CodiceClienteInModifica = fg.TextMatrix(fg.Row, 2)
    AMSRagSociale.Text = fg.TextMatrix(fg.Row, 0)
    AMST1.Text = fg.TextMatrix(fg.Row, 6)
End If
Exit Sub
Errore:
    GestioneErrori Me.Name, "fg_DblClick", Err.Number, Err.Description: Resume Next
End Sub

Private Sub ARFiltri_Click()
    AMSRagSociale.Text = ""
    AMST1.Text = ""
    Selezione
End Sub

Private Sub ARSalva_Click()
On Error GoTo Errore
Dim sqlUpdate As String
Dim telNumerico As String

If Trim(m_CodiceClienteInModifica) = "" Then
    MsgBox "Seleziona prima un cliente", vbInformation, "Attenzione!"
    Exit Sub
End If

telNumerico = Trim(AMST1.Text)
If telNumerico = "" Or Not IsNumeric(telNumerico) Then
    telNumerico = "0"
End If

sqlUpdate = "UPDATE Clienti SET " & _
            " ACFRagSoc = '" & Replace(UCase(AMSRagSociale.Text), "'", "''") & "', " & _
            " ACFTel1 = " & telNumerico & _
            " WHERE ACFCod = '" & Replace(m_CodiceClienteInModifica, "'", "''") & "'"
DB.Execute sqlUpdate

m_CodiceClienteInModifica = ""
Selezione
Exit Sub

Errore:
    GestioneErrori Me.Name, "ARSalva_Click", Err.Number, Err.Description: Resume Next
End Sub

Private Sub ARCancella_Click()
On Error GoTo Errore
Dim rsVerifica As New ADODB.Recordset
Dim sqlCheck As String
Dim sqlDelete As String

If Trim(m_CodiceClienteInModifica) = "" Then
    MsgBox "Seleziona prima un cliente", vbInformation, "Attenzione!"
    Exit Sub
End If

sqlCheck = "SELECT COUNT(*) AS totale FROM DocumentoTestata " & _
            "Where dctCodiceCliente = '" & Replace(m_CodiceClienteInModifica, "'", "''") & "'"
            
rsVerifica.Open sqlCheck, DB, adOpenStatic, adLockReadOnly
If Not rsVerifica.EOF Then
    If rsVerifica("Totale") > 0 Then
    MsgBox "Impossibile eliminare il cliente: sono presenti " & rsVerifica("Totale") & " vendite associate", vbCritical, "Operazione Fallita"
    rsVerifica.Close
    Exit Sub
    End If
End If
rsVerifica.Close

sqlDelete = "DELETE FROM Clienti " & _
            "WHERE ACFCod= '" & Replace(m_CodiceClienteInModifica, "'", "''") & "'"
DB.Execute sqlDelete
m_CodiceClienteInModifica = ""
ARFiltri_Click
Exit Sub
Errore:
    GestioneErrori Me.Name, "ARCancella_Click", Err.Number, Err.Description: Resume Next
End Sub
'Roberta Sciarrone 04/09/2026 fine
