VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInventoryAct"
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
Public Function QueryInventoryAct(Ua As cInventoryAct, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim Iai As CInventoryActItem

   RName = "QueryInventoryAct"
   QueryInventoryAct = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
    'CollRawMaterials ++++++++++++++++++++++++++
      Set Iai = New CInventoryActItem
      Iai.INVENTORY_ACT_ID = Ua.INVENTORY_ACT_ID
      Iai.INVENTORY_ACT_ITEM_TYPE = 1
      Iai.FROM_DATE = -1
      Iai.TO_DATE = -1
      Iai.OrderBy = 4 'Ua.OrderBy
      Iai.OrderType = Ua.OrderType
      
      Call Iai.QueryData(1, m_Rs2, iCount)
      Set Iai = Nothing
      
      Set Ua.CollRawMaterials = Nothing
      Set Ua.CollRawMaterials = New Collection
      While Not m_Rs2.EOF
         Set Iai = New CInventoryActItem
         Call Iai.PopulateFromRS(1, m_Rs2)
         Iai.Flag = "I"
         Call Ua.CollRawMaterials.add(Iai)
         Set Iai = Nothing
         m_Rs2.MoveNext
      Wend
      Set Iai = Nothing
'    CollRawMaterials ++++++++++++++++++++++++++
'    CollPhamacyRoom ++++++++++++++++++++++++++
   Set Iai = New CInventoryActItem
   Iai.INVENTORY_ACT_ID = Ua.INVENTORY_ACT_ID
   Iai.INVENTORY_ACT_ITEM_TYPE = 2
   Iai.FROM_DATE = -1
   Iai.TO_DATE = -1
   Iai.OrderBy = 4 'Ua.OrderBy
   Iai.OrderType = Ua.OrderType
   Call Iai.QueryData(1, m_Rs2, iCount)
   Set Iai = Nothing
      
   Set Ua.CollPhamacyRoom = Nothing
   Set Ua.CollPhamacyRoom = New Collection
   While Not m_Rs2.EOF
      Set Iai = New CInventoryActItem
      Call Iai.PopulateFromRS(1, m_Rs2)
      Iai.Flag = "I"
      Call Ua.CollPhamacyRoom.add(Iai)
      Set Iai = Nothing
      m_Rs2.MoveNext
   Wend
   Set Iai = Nothing
'    CollPhamacyRoom ++++++++++++++++++++++++++
'    CollSolo ++++++++++++++++++++++++++
   Set Iai = New CInventoryActItem
   Iai.INVENTORY_ACT_ID = Ua.INVENTORY_ACT_ID
   Iai.INVENTORY_ACT_ITEM_TYPE = 3
   Iai.FROM_DATE = -1
   Iai.TO_DATE = -1
   Iai.OrderBy = 4
   Iai.OrderType = Ua.OrderType
   Call Iai.QueryData(1, m_Rs2, iCount)
   Set Iai = Nothing
      
   Set Ua.CollSilo = Nothing
   Set Ua.CollSilo = New Collection
   While Not m_Rs2.EOF
      Set Iai = New CInventoryActItem
      Call Iai.PopulateFromRS(1, m_Rs2)
      Iai.Flag = "I"
      Call Ua.CollSilo.add(Iai)
      Set Iai = Nothing
      m_Rs2.MoveNext
   Wend
   Set Iai = Nothing
'    CollSilo ++++++++++++++++++++++++++


   End If

   QueryInventoryAct = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryAct = False
End Function
Public Function AddEditInventoryAct(Ug As cInventoryAct, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditInventoryAct"
   AddEditInventoryAct = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim Iai As CInventoryActItem
   For Each Iai In Ug.CollRawMaterials
      If Iai.Flag = "A" Then
         Iai.INVENTORY_ACT_ID = Ug.INVENTORY_ACT_ID
         Iai.INVENTORY_ACT_ITEM_TYPE = 1
         Iai.AddEditMode = SHOW_ADD
         Iai.AddEditData
      ElseIf Iai.Flag = "E" Then
         Iai.AddEditMode = SHOW_EDIT
         Iai.AddEditData
      ElseIf Iai.Flag = "D" Then
         Call Iai.DeleteData
      End If
   Next Iai
   
   For Each Iai In Ug.CollPhamacyRoom
      If Iai.Flag = "A" Then
         Iai.INVENTORY_ACT_ID = Ug.INVENTORY_ACT_ID
         Iai.INVENTORY_ACT_ITEM_TYPE = 2
         Iai.AddEditMode = SHOW_ADD
         Iai.AddEditData
      ElseIf Iai.Flag = "E" Then
         Iai.AddEditMode = SHOW_EDIT
         Iai.AddEditData
      ElseIf Iai.Flag = "D" Then
         Call Iai.DeleteData
      End If
   Next Iai
   
   For Each Iai In Ug.CollSilo
      If Iai.Flag = "A" Then
         Iai.INVENTORY_ACT_ID = Ug.INVENTORY_ACT_ID
         Iai.INVENTORY_ACT_ITEM_TYPE = 3
         Iai.AddEditMode = SHOW_ADD
         Iai.AddEditData
      ElseIf Iai.Flag = "E" Then
         Iai.AddEditMode = SHOW_EDIT
         Iai.AddEditData
      ElseIf Iai.Flag = "D" Then
         Call Iai.DeleteData
      End If
   Next Iai
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditInventoryAct = True
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
   AddEditInventoryAct = False
End Function

Public Function DeleteInventoryAct(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As cInventoryAct

   RName = "DeleteInventoryAct"
   DeleteInventoryAct = False
   HasBegin = False

   Set B = New cInventoryAct

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.INVENTORY_ACT_ID = UID
   
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteInventoryAct = True
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

   DeleteInventoryAct = True
End Function
