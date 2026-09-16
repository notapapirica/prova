Attribute VB_Name = "Dichiarazioni"
Option Explicit


'STRUTTURA TICKET LOG
Public Type TickLog
    TipoOperazione As String
    descrizione As String
    ResoDaScCortesia As Boolean
    TotControllato As Boolean
    BarcodeVendita As String
    riferimentivendita As String
    TotaleManuale As Double
    Utente As String
    IdVendita As Long
    DataVendita As Date
    NumVendita As Long
End Type







Public Azikey As String                   'Codice Azienda per eventuali personalizzazioni
Global strGruppo As String                'AMMINISTRATORE,OPERATORE,UTENTE
Global strUtente As String                'UTENTE NEL SISTEMA
Public Tastiera As Integer
Public DepositoKids As String
Public Validate As Integer
Public SqlStr As String                   'Stringa Per Query Sql
Public TipoDocCliente As Integer
Global strEcommerce As String
Public Esci As Boolean
Public TipoDocDettaglio As Integer
Public TipodocInterno As Integer
Public TipodocOrdine As Integer
Public TipodocCarico As Integer
Public CaptionMnu  As String
Public FrmClienteVisible As Boolean
Public FrmCliente2Visible As Boolean
Public FrmFornitoreVisible As Boolean
Public FrmFornitore2Visible As Boolean
Public frmAnagraficaAziendaVisible As Boolean
Public frmVenditadettagliovisible As Boolean
Public frmdocumentointernovisible As Boolean
Public frmAnaDestVisible As Boolean
Public Intesta As String                   'Stringa Per Intestazione
Public lastqtamov As Double
Public Dati(9) As String
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000&
Public rsOldTaglieColori As Recordset
Global BaseDati As String

'GESTIONE STAMPA                        ' --- DA NON TOGLIERE ASSOLUTAMENTE ---
Global Menu As String                   'VARIABILE UTILIZZATA DAI FORM DEI REPORT
Global com As String                    'VARIABILE UTILIZZATA DAI FORM DEI REPORT flag di controllo
Global provenienza As String            'VARIABILE UTILIZZATA DAI FORM DEI REPORT provenienza del form frmricanaprod
Global contrcmd As String               'VARIABILE UTILIZZATA DAI FORM DEI REPORT controllo per sapere quale pulsante ho cliccato nei form

'GESTIONE CONNESSIONE DATABASE DEI REPORT
Public CrLogonID As Long             'VARIABILE UTILIZZATA PER CREARE LA CONNESSIONE DINAMICA DEL DATABASE A CUI PUNTANO I REPORT

'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------

'NEW DICHIARAZIONI BY PAGANO ANDREA 03/11/2003
'PERSONALIZAZZIONE  PER TIPO DOC "BUS" 03/02/2005 'SICILBUS
'VARIABILI VALORIZZATE DAL REGISTRO DI SISTEMA
'------------------------------------------------------------------------------------------------------
  Public regDbUtente As String 'ciccio001
  Public regDbPassword As String 'ciccio001
  'PREFISSO PROTOCOLLO
'  Public regPrefissoProtocollo As String
  'ULTIMO ACCESSO
  Public regUtente As String
  'LICENZA
  Public regNumeroSeriale As String
  Public regNumeroSerialeHD As String
  'PROFILI
  Public regNomeDatabase As String
  Public regLogInData As String
  Public regLogInOra As String
  Public regLogOutData As String
  Public regLogOutOra As String
  Public regUltimoMenu As String
  Public regMessaggio As String
  'CODICE A BARRE
  Public regPortaConnessione As String
  Public regAltezzaEtichetta As String
  Public regLarghezzaEtichetta As String
  Public regCodiceEan8 As String
  Public regCodiceEan13 As String
  Public regCodice128 As String
  Public regCodiceArticolo As String
  Public regPrezzoArticolo As String
  Public regDescrizioneArticolo As String
  Public regStampante As String
  Public regTaglia As String
  Public regColore As String
  Public regFornAbit As String
  Public regRifForn As String
  Public regVarie1 As String
  Public regVarie2 As String
  Public regVarie3 As String
  Public regVarie4 As String
  Public regVarie5 As String
  Public regVarie6 As String
  Public regVarie7 As String
  Public WebSvrAtelier As New WSATelier
  Public WebSvrAtelierClient As New WSATelier

'FUNZIONE NDecimali
'------------------------------------------------------------------------------------------------------
  Public Const Importo   As Long = 1 'NDecimali
  Public Const Vendita   As Long = 2 'NDecimali
  Public Const Acquisto  As Long = 3 'NDecimali

'GESTIONE DATABASE PER UTENTI E AZIENDE
'------------------------------------------------------------------------------------------------------
  Global ConnessioneStart As String       'UTENTI - STRINGA CONNESSIONE AL DATABASE
  Global DBStart As Connection            'UTENTI - DATABASE
  Global dbNamelogin As String            'UTENTI - NOME DEL DATABASE O DSN
  Global Connessione As String            'AZIENDA - STRINGA CONNESSIONE AL DATABASE
  Global DB As Connection                 'AZIENDA - DATABASE
  Global Dbname As String                 'AZIENDA - NOME DEL DATABASE O DSN
  Global DbnameProvv As String                 'AZIENDA - NOME DEL DATABASE O DSN
  Global ShortDbname As String
  Global ShortDbnamelogin As String
  Global TipoAzienda As String
  Global Numazienda As Long
  Global Server1 As Long
  Global Server2 As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type



'FUNZIONI LETTURA E SCRITTURA REGISTRO
'------------------------------------------------------------------------------------------------------
  Global Const KEY_QUERY_VALUE = &H1
  Global Const ERROR_SUCCESS = 0&
  Global Const REG_SZ = 1
  Global Const HKEY_LOCAL_MACHINE = &H80000002
  Global Const REG_DWORD = 4
  Global Const WM_CLOSE = &H10
  Global Const WM_ACTIVATE = &H6
  Public Const GWL_EXSTYLE = (-20)
  Public Const WS_EX_APPWINDOW = &H40000
  Public Const GWL_STYLE = -16&

 Public Enum DrawCombo
    FC_DRAWNORMAL = 0
    FC_DRAWRAISED = 1
    FC_DRAWPRESSED = 2
    FC_DRAWDISABLED = 3
End Enum

' the style of the pen used to create the mask over the combobox
Public Const SM_CXHTHUMB = 10
Public Const PS_SOLID = 0
Public Const PS_DOT = 2

'MOBILE DEVICE
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
 
Type RAPIINIT
    cbSize As Long
    heRapiInit As Long
    hrRapiInit As Long
End Type
 
Declare Function CeRapiUninit Lib "rapi.dll" () As Long

Declare Function CeRapiInitEx Lib "rapi.dll" ( _
    pRapiInit As RAPIINIT) As Long
    
Declare Function CeRegOpenKeyEx Lib "rapi.dll" ( _
    ByVal hkey As Long, _
    ByVal lpSubKey As Long, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
    
Declare Function CeRegQueryValueEx Lib "rapi.dll" ( _
    ByVal hkey As Long, _
    ByVal lpValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpdata As Long, _
    lpcbData As Long) As Long
    
Declare Function CeRegQueryValueExLong Lib "rapi.dll" _
    Alias "CeRegQueryValueEx" ( _
    ByVal hkey As Long, _
    ByVal lpValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpdata As Long, _
    lpcbData As Long) As Long
    
Declare Function CeRegQueryValueExString Lib "rapi.dll" _
    Alias "CeRegQueryValueEx" ( _
    ByVal hkey As Long, _
    ByVal lpValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpdata As Long, _
    lpcbData As Long) As Long

Declare Function CeRegCloseKey Lib "rapi.dll" ( _
    ByVal hkey As Long) As Long
 
 '***
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpdata As Any, lpcbData As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpvalue As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpvalue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
  'FUNCTION Declarations
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ShowWindow Lib "user32" (ByVal hdc As Long, ByVal bEnable As Boolean) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function AllowSetForegroundWindow Lib "user32.dll" (ByVal dwProcessId As Long) As Long
Declare Function SetFocusA Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
' Declare Function for Open Communication Serial
Declare Function VbCEFOpen Lib "CeFdll.dll" (ByVal intCom As Long, ByVal dwBaudRate As Long, ByVal byParity As Byte, ByVal byDataBit As Byte, ByVal byStopBit As Byte, ByVal byFlowControl As Byte, lpdwSysError As Long) As Long
' Declare Function for Open Communication Ethernet
Declare Function VbCEFOpenEth Lib "CeFdll.dll" (textip As String, ByVal dwPort As Long, lpdwSysError As Long) As Long
' Declare Function for Write Command on Communication Port
Declare Function VbCEFWrite Lib "CeFdll.dll" (textcmd As String, lpdwSysError As Long) As Long
' Declare Function for Read from Communication Port
Declare Function VbCEFRead Lib "CeFdll.dll" (RetData() As Byte, pdwByteRead As Long, lpdwSysError As Long) As Long
' Declare Function for Read DLL Version
Declare Function VbCEFGetVersion Lib "CeFdll.dll" (RetData() As Byte, lpdwSysError As Long) As Long
' Declare Function for Close Communication Port
Declare Function VbCEFClose Lib "CeFdll.dll" (lpdwSysError As Long) As Long


Const SYNCHRONIZE = &H100000
   '
   ' Wait forever
   '
   Const INFINITE = &HFFFF
   '
   ' The state of the specified object is signaled
   '
   Const WAIT_OBJECT_0 = 0
   '
   ' The time-out interval elapsed & the object’s state is not signaled
   '
   Const WAIT_TIMEOUT = &H102

   Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

   Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
   Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Declare Function OleTranslateColor Lib "oleaut32" (ByVal clr As OLE_COLOR, ByVal hPal As Long, dwRGB As Long) As Long

'Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, pdwResult As Long) As Long
Declare Function TerminateProcess Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetLayeredWindowAttributes _
Lib "user32" _
(ByVal hWnd As Long, _
ByVal crKey As Long, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long



'Private Declare Function GetWindowLong _
'Lib "user32" Alias "GetWindowLongA" _
'(ByVal hwnd As Long, _
'ByVal nIndex As Long) As Long

'Private Declare Function SetWindowLong _
'Lib "user32" Alias "SetWindowLongA" _
'(ByVal hwnd As Long, _
'ByVal nIndex As Long, _
'ByVal dwNewLong As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" _
                            (ByVal hWnd As Long, _
                            ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" _
                            (ByVal hMenu As Long, _
                            ByVal nPosition As Long, _
                            ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" _
                            (ByVal hWnd As Long) As Long
Public Const MF_BYCOMMAND = &H0
Public Const SC_MAXIMIZE As Long = &HF030&
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const HWND_TOP As Long = 0
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE  As Long = &H1

Public Porta As String
Public BoudRate As String
Public DataBits As String
Public Cassetto As String

Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long


Public Sub FormFade(ByRef frmForm As Form, ByVal Opacity As Long)
    Dim Msg As Long
    frmForm.Show vbModeless
    Msg = GetWindowLong(frmForm.hWnd, GWL_EXSTYLE)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong frmForm.hWnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes frmForm.hWnd, 0, Opacity, LWA_ALPHA
    frmForm.Refresh
End Sub





