VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WSATelier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private clnt As SoapClient30
Private m_dReturnErrorNumber As Double
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()

Public Sub Class_Initialize()
  InitCommonControls
End Sub

Public Function TestWebService(ByVal UrlConnection As String, Client As String, Deposito As String) As Boolean
On Error GoTo TestWebservice_Error
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 20000
  TestWebService = clnt.testConnessioneAmsAtelierClient(Client, Deposito)
'  ReturnErrorNumber = Err.Number
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
  Set clnt = Nothing
End Function

Public Function TestWebServiceClientTrasf(ByVal UrlConnection As String, Deposito As String, Deposito2 As String) As Boolean
On Error GoTo TestWebservice_Error
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 20000
  TestWebServiceClientTrasf = clnt.testConnessioneAtelierClientTrasf(Deposito, Deposito2)
'  ReturnErrorNumber = Err.Number
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
'  MsgBox "Errore Soap:" & UrlConnection & "?WSDL"
  Set clnt = Nothing
End Function

'Public Function TestWebServiceClientTrasf(ByVal UrlConnection As String, Deposito As String, Deposito2 As String) As Boolean
'On Error GoTo Errore
'  If clnt Is Nothing Then Set clnt = New SoapClient30
''  clnt.MSSoapInit UrlConnection & "?WSDL"
'  clnt.MSSoapInit UrlConnection & "?WSDL"
'  clnt.ConnectorProperty("Timeout") = 20000
'  TestWebServiceClientTrasf = clnt.testConnessioneAtelierClientTrasf(Deposito, Deposito2)
''  ReturnErrorNumber = Err.Number
'  Set clnt = Nothing
'Exit Function
'Errore:
'  MsgBox Err.Description
'  GestioneErrori "WebService", "TestWebServiceClientTrasf", Err.Number, Err.Description: Resume Next
'End Function

Public Function Recovery(ByVal UrlConnection As String, Client As String, Deposito As String, azienda As String) As Boolean
On Error GoTo TestWebservice_Error
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 1800000
  Recovery = clnt.Recovery(Client, Deposito, azienda)
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
  Set clnt = Nothing
End Function

Public Function TestWebServiceClient(ByVal UrlConnection As String, Deposito As String) As Boolean
On Error GoTo Errore
'On Error GoTo TestWebservice_Error
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 20000
  TestWebServiceClient = clnt.testConnessioneAmsAtelierClient(Deposito)
'  ReturnErrorNumber = Err.Number
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
  Set clnt = Nothing
Errore:
    GestioneErrori "Wsatelier", "Ws", Err.Number, Err.Description: Resume Next
End Function

Public Function TestWebServiceClientProdottoGenerico(ByVal UrlConnection As String, Deposito As String) As Boolean
On Error GoTo TestWebservice_Error
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 20000
  TestWebServiceClientProdottoGenerico = clnt.testConnessioneAmsAtelierClient("", Deposito)
'  ReturnErrorNumber = Err.Number
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
  Set clnt = Nothing
End Function




'Public Function TestWebServiceClient(ByVal UrlConnection As String, Deposito As String) As Boolean
'On Error GoTo Errore
'  If clnt Is Nothing Then Set clnt = New SoapClient30
'  clnt.MSSoapInit UrlConnection & "?WSDL"
'  clnt.ConnectorProperty("Timeout") = 20000
''  MsgBox UrlConnection
''  MsgBox clnt.testConnessioneAmsAtelierClient("", Deposito)
'  TestWebServiceClient = clnt.testConnessioneAmsAtelierClient(Deposito)
'  Set clnt = Nothing
'Exit Function
'Errore:
'  MsgBox Err.Description
'  GestioneErrori "WebService", "TestWebServiceClient", Err.Number, Err.Description: Resume Next
'End Function

Public Function RecoveryClient(ByVal UrlConnection As String, Deposito As String, azienda As String) As Boolean
'On Error GoTo TestWebservice_Error
On Error GoTo Errore
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 200000
  RecoveryClient = clnt.Recovery(Deposito, azienda)
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
  Set clnt = Nothing
Errore:
    GestioneErrori "Wsatelier", "Ws", Err.Number, Err.Description: Resume Next
End Function

Public Function Transfer(ByVal UrlConnection As String, Deposito As Long, Deposito2 As Long, IdTestata As Long) As Boolean
On Error GoTo TestWebservice_Error
  If clnt Is Nothing Then Set clnt = New SoapClient30
  clnt.MSSoapInit UrlConnection & "?WSDL"
  clnt.ConnectorProperty("Timeout") = 1800000
  Transfer = clnt.Transfer(Deposito, Deposito2, IdTestata)
  Set clnt = Nothing
Exit Function
TestWebservice_Error:
  Set clnt = Nothing
End Function

