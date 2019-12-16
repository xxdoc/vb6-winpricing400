VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlanning"
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
   m_ModuleName = "clsPlanning"
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
Public Function QueryPlanning(Ua As CPlanning, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim Pni As CPlanningItem

   RName = "QueryPlanning"
   QueryPlanning = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
    'CollPartUse ++++++++++++++++++++++++++
      
      Set Pni = New CPlanningItem
      Pni.PLANNING_ID = Ua.PLANNING_ID
      Pni.PLANNING_SUB_TYPE = 1
      Pni.OrderBy = Ua.OrderBy
      Pni.OrderType = Ua.OrderType
      Call Pni.QueryData(1, m_Rs2, iCount)
      Set Pni = Nothing
      
      Set Ua.CollPartUse = Nothing
      Set Ua.CollPartUse = New Collection
      While Not m_Rs2.EOF
         Set Pni = New CPlanningItem
         Call Pni.PopulateFromRS(1, m_Rs2)
         Pni.Flag = "I"
         Call Ua.CollPartUse.add(Pni)
         Set Pni = Nothing
         m_Rs2.MoveNext
      Wend
      Set Pni = Nothing
'    CollPartUse ++++++++++++++++++++++++++

    'CollProductGet ++++++++++++++++++++++++++
      Set Pni = New CPlanningItem
      Pni.PLANNING_ID = Ua.PLANNING_ID
      Pni.PLANNING_SUB_TYPE = 2
      Call Pni.QueryData(1, m_Rs2, iCount)
      Set Pni = Nothing

      Set Ua.CollProductGet = Nothing
      Set Ua.CollProductGet = New Collection
      While Not m_Rs2.EOF
         Set Pni = New CPlanningItem
         Call Pni.PopulateFromRS(1, m_Rs2)
         Pni.Flag = "I"
         Call Ua.CollProductGet.add(Pni)
         Set Pni = Nothing
         m_Rs2.MoveNext
      Wend
      Set Pni = Nothing
   'CollProductGet ++++++++++++++++++++++++++


    'CollPartSup ++++++++++++++++++++++++++
      Set Pni = New CPlanningItem
      Pni.PLANNING_ID = Ua.PLANNING_ID
      Pni.PLANNING_SUB_TYPE = 3
      Call Pni.QueryData(1, m_Rs2, iCount)
      Set Pni = Nothing

      Set Ua.CollPartSup = Nothing
      Set Ua.CollPartSup = New Collection
      While Not m_Rs2.EOF
         Set Pni = New CPlanningItem
         Call Pni.PopulateFromRS(1, m_Rs2)
         Pni.Flag = "I"
         Call Ua.CollPartSup.add(Pni)
         Set Pni = Nothing
         m_Rs2.MoveNext
      Wend
      Set Pni = Nothing
   'CollPartSup ++++++++++++++++++++++++++
'   ElseIf Ua.QueryFlag = 2 Then
'    'CollPartUse ++++++++++++++++++++++++++
'      Set Pni = New CPlanningItem
'      Pni.PLANNING_ID = Ua.PLANNING_ID
'      Pni.PLANNING_SUB_TYPE = 1
'      Pni.OrderBy = Ua.OrderBy
'      Pni.OrderType = Ua.OrderType
'      Call Pni.QueryData(1, m_Rs2, iCount)
'      Set Pni = Nothing
'
'      Set Ua.CollPartUse = Nothing
'      Set Ua.CollPartUse = New Collection
'      While Not m_Rs2.EOF
'         Set Pni = New CPlanningItem
'         Call Pni.PopulateFromRS(1, m_Rs2)
'         Pni.Flag = "I"
'         Call Ua.CollPartUse.add(Pni)
'         Set Pni = Nothing
'         m_Rs2.MoveNext
'      Wend
'      Set Pni = Nothing
''    CollPartUse ++++++++++++++++++++++++++
   End If

   QueryPlanning = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPlanning = False
End Function
Public Function AddEditPlanning(Ug As CPlanning, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPlanning"
   AddEditPlanning = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim Pni As CPlanningItem
   For Each Pni In Ug.CollPartUse
      If Pni.Flag = "A" Then
         Pni.PLANNING_ID = Ug.PLANNING_ID
         Pni.PLANNING_SUB_TYPE = 1
         Pni.AddEditMode = SHOW_ADD
         Pni.AddEditData
      ElseIf Pni.Flag = "E" Then
         Pni.AddEditMode = SHOW_EDIT
         Pni.AddEditData
      ElseIf Pni.Flag = "D" Then
         Call Pni.DeleteData
      End If
   Next Pni
   
  For Each Pni In Ug.CollProductGet
      If Pni.Flag = "A" Then
         Pni.PLANNING_ID = Ug.PLANNING_ID
         Pni.PLANNING_SUB_TYPE = 2
         Pni.AddEditMode = SHOW_ADD
         Pni.AddEditData
      ElseIf Pni.Flag = "E" Then
         Pni.AddEditMode = SHOW_EDIT
         Pni.AddEditData
      ElseIf Pni.Flag = "D" Then
         Call Pni.DeleteData
      End If
   Next Pni
   
   For Each Pni In Ug.CollPartSup
      If Pni.Flag = "A" Then
         Pni.PLANNING_ID = Ug.PLANNING_ID
         Pni.PLANNING_SUB_TYPE = 3
         Pni.AddEditMode = SHOW_ADD
         Pni.AddEditData
      ElseIf Pni.Flag = "E" Then
         Pni.AddEditMode = SHOW_EDIT
         Pni.AddEditData
      ElseIf Pni.Flag = "D" Then
         Call Pni.DeleteData
      End If
   Next Pni
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPlanning = True
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
   AddEditPlanning = False
End Function

Public Function DeletePlanning(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CPlanning

   RName = "DeletePlanning"
   DeletePlanning = False
   HasBegin = False

   Set B = New CPlanning

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PLANNING_ID = UID
   
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePlanning = True
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

   DeletePlanning = True
End Function
