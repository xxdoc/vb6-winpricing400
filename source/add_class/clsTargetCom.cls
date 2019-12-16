VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTargetCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsTarget"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub
'==
Public Function QueryTarget(Ua As CTarget, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim Tgdt As CTargetDetail

   RName = "QueryTarget"
   QueryTarget = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
    'CollPartUse ++++++++++++++++++++++++++
      
      Set Tgdt = New CTargetDetail
      Tgdt.TARGET_ID = Ua.TARGET_ID
      Call Tgdt.QueryData(1, m_Rs2, iCount)
      Set Tgdt = Nothing
      
      Set Ua.CollTargerDetail = Nothing
      Set Ua.CollTargerDetail = New Collection
      While Not m_Rs2.EOF
         Set Tgdt = New CTargetDetail
         Call Tgdt.PopulateFromRS(1, m_Rs2)
         Tgdt.Flag = "I"
         Call Ua.CollTargerDetail.add(Tgdt)
         Set Tgdt = Nothing
         m_Rs2.MoveNext
      Wend
      Set Tgdt = Nothing

   End If

   QueryTarget = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTarget = False
End Function
Public Function AddEditTarget(Ug As CTarget, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditTarget"
   AddEditTarget = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim Tgdt As CTargetDetail
   For Each Tgdt In Ug.CollTargerDetail
      If Tgdt.Flag = "A" Then
         Tgdt.TARGET_ID = Ug.TARGET_ID
         Tgdt.AddEditMode = SHOW_ADD
         Tgdt.AddEditData
      ElseIf Tgdt.Flag = "E" Then
         Tgdt.AddEditMode = SHOW_EDIT
         Tgdt.AddEditData
      ElseIf Tgdt.Flag = "D" Then
         Call Tgdt.DeleteData
      End If
   Next Tgdt
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditTarget = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditTarget = False
End Function
Public Function DeleteTarget(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CTarget

   RName = "DeleteTarget"
   DeleteTarget = False
   HasBegin = False

   Set B = New CTarget

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.TARGET_ID = UID
   
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteTarget = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteTarget = True
End Function
Public Function CopyTarget(Ug As CTarget, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyTarget"
   CopyTarget = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryTarget(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.TARGET_ID = -1
   Ug.YEAR_NO = Ug.YEAR_NO + 1
   Ug.TARGET_DESC = Ug.TARGET_DESC & "******"
   
   Dim Tgdt As CTargetDetail
   For Each Tgdt In Ug.CollTargerDetail
      Tgdt.Flag = "A"
      Tgdt.AddEditMode = SHOW_ADD
   Next Tgdt
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditTarget(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyTarget = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyTarget = False
End Function

