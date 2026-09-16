VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' custom FlexGrid data sources implement IVSFlexDataSource
Implements IVSFlexDataSource

' ADO objects and record pointer
Public Dbname As String
Public sql As String
'Public m_Conn As New Connection
Public m_RS As Recordset
Dim m_Filter() As String
Dim m_lRecord As Long

' initialize class: create recordset
Public Sub Class_open()
   ' If m_Conn.State = adStateOpen Then m_Conn.Close
   ' m_Conn.Open Dbname
    Set m_RS = New Recordset
    'm_RS.CursorLocation = adUseClient
    m_RS.Open sql, DB, adOpenKeyset
    ReDim m_Filter(m_RS.Fields.Count)
    m_lRecord = -1
End Sub

Public Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String
'On Error Resume Next
    ' first record is custom (search record)
    If Record = 0 Then
        IVSFlexDataSource_GetData = m_Filter(Field)
        Exit Function
    End If
    If m_RS.RecordCount = 0 Then Exit Function
    If m_RS.RecordCount < Record Then GoTo sALTA
    ' move cursor if we have to
    If m_lRecord <> Record Then
        m_lRecord = Record
        m_RS.AbsolutePosition = Record
    End If
sALTA:
    If m_RS.EOF = True Then Exit Function
     If Not IsNull(m_RS.Fields(Field).Value) Then
        IVSFlexDataSource_GetData = m_RS.Fields(Field).Value
   ' retrieve value
    End If
    
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = m_RS.Fields.Count
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    IVSFlexDataSource_GetFieldName = m_RS.Fields(Field).Name
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = m_RS.RecordCount + 1
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
Dim Filtro As String
Dim FiltroNum As String
    ' we are only interested on the filter line
    If Record <> 0 Then Exit Sub
    
    ' save new item
    m_Filter(Field) = Trim(newData)
'    If InStr("><=", Left(m_Filter(Field), 1)) = 0 Then
'      Filtro = Trim(Replace(Replace(Replace(m_Filter(Field), "%", ""), "'", ""), "LIKE", ""))
'      If Not IsNumeric(Filtro) Then
''        If Mid(Filtro, 1, 1) <> "'" Then Filtro = "LIKE '" & Filtro &  "%" & "'"
'         Filtro = "LIKE '%" & Filtro & "%'"
'         FiltroNum = ""
''        If Mid(Filtro, Len(Filtro) - 1, 1) <> "'" Then Filtro = Filtro & "'"
'      Else
'       FiltroNum = "=" & Filtro & ""
'       Filtro = "LIKE '%" & Filtro & "%'"
'      End If
'    '  If m_RS.Fields(Field).Type <> 200 And m_RS.Fields(Field).Type <> 202 Then m_Filter(Field) = FiltroNum Else m_Filter(Field) = Filtro
'
'    End If
        
    ' build filter string
    Dim i%, sFilter$
    For i = 0 To UBound(m_Filter)
        If Len(m_Filter(i)) Then
            If Len(sFilter) Then sFilter = sFilter & " AND "
            sFilter = sFilter & m_RS.Fields(i).Name & " "
            If IsNumeric(m_Filter(i)) Then
              sFilter = sFilter & "=" & m_Filter(i)
            Else
              sFilter = sFilter & "LIKE '%" & m_Filter(i) & "%'"
            End If
        End If
    Next
        
    ' no change? no work (note that ADO converts "" into 0)
    Dim sOldFilter$
    sOldFilter = m_RS.Filter
    If sOldFilter = "0" Then sOldFilter = ""
    If sFilter = sOldFilter Then Exit Sub
    
    ' assign filter string
    On Error Resume Next
    m_RS.Filter = sFilter
    If Err > 0 Then
      '  MsgBox "Attenzione nei tipi dati stringa , bisogna racchiudere la stringa tra apici singoli (es. 'testo01').", vbInformation
        m_RS.Filter = ""
    End If
    m_RS.Requery
    On Error GoTo 0
End Sub


