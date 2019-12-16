VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsMaster"
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

'==
Public Function AddEditPartType(At As CPartType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartType"
   AddEditPartType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPartType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPartType = False
End Function

Public Function AddEditPartGroup(At As CPartGroup, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartGroup"
   AddEditPartGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Dim CstName As CHGroupItem
   For Each CstName In At.PartGroupItems
      If CstName.Flag = "A" Then
         CstName.PART_GROUP_ID = At.PART_GROUP_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPartGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPartGroup = False
End Function

Public Function AddEditHouseGroup(At As CHouseGroup, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditHouseGroup"
   AddEditHouseGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
'   Dim CstName As CHGroupItem
'   For Each CstName In At.HGroupItems
'      If CstName.Flag = "A" Then
'         CstName.PART_GROUP_ID = At.p
'         CstName.AddEditMode = SHOW_ADD
'         CstName.AddEditData
'      ElseIf CstName.Flag = "E" Then
'         CstName.AddEditMode = SHOW_EDIT
'         CstName.AddEditData
'      ElseIf CstName.Flag = "D" Then
'         CstName.DeleteData
'      End If
'   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditHouseGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditHouseGroup = False
End Function

Public Function AddEditStatusGroup(At As CStatusGroup, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditStatusGroup"
   AddEditStatusGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Dim CstName As CSGroupItem
   For Each CstName In At.HGroupItems
      If CstName.Flag = "A" Then
         CstName.STATUS_GROUP_ID = At.STATUS_GROUP_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditStatusGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditStatusGroup = False
End Function

Public Function DeletePartType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CPartType

   RName = "DeletePartType"
   DeletePartType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CPartType
   cData.PART_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePartType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePartType = False
End Function

Public Function DeletePartGroup(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CPartGroup

   RName = "DeletePartGroup"
   DeletePartGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CPartGroup
   cData.PART_GROUP_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePartGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePartGroup = False
End Function

Public Function DeleteHouseGroup(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CHouseGroup

   RName = "DeleteHouseGroup"
   DeleteHouseGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CHouseGroup
   cData.HOUSE_GROUP_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteHouseGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteHouseGroup = False
End Function

Public Function DeleteStatusGroup(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CStatusGroup

   RName = "DeleteStatusGroup"
   DeleteStatusGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CStatusGroup
   cData.STATUS_GROUP_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteStatusGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteStatusGroup = False
End Function

'==
Public Function AddEditLocation(At As CLocation, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLocation"
   AddEditLocation = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditLocation = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLocation = False
End Function

Public Function DeleteLocation(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CLocation

   RName = "DeleteLocation"
   DeleteLocation = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CLocation
   cData.LOCATION_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteLocation = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteLocation = False
End Function

'==
Public Function AddEditProductType(At As CProductType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditProductType"
   AddEditProductType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditProductType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditProductType = False
End Function

Public Function DeleteProductType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CProductType

   RName = "DeleteProductType"
   DeleteProductType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CProductType
   cData.PRODUCT_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteProductType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteProductType = False
End Function

'==
Public Function AddEditProductStatus(At As CProductStatus, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditProductStatus"
   AddEditProductStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditProductStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditProductStatus = False
End Function

Public Function DeleteProductStatus(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CProductStatus

   RName = "DeleteProductStatus"
   DeleteProductStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CProductStatus
   cData.PRODUCT_STATUS_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteProductStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteProductStatus = False
End Function
'==
Public Function AddEditHouse(At As CHouse, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditHouse"
   AddEditHouse = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditHouse = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditHouse = False
End Function

Public Function AddEditSellType(At As CSellType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSellType"
   AddEditSellType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSellType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSellType = False
End Function

Public Function DeleteHouse(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CHouse

   RName = "DeleteHouse"
   DeleteHouse = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CHouse
   cData.HOUSE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteHouse = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteHouse = False
End Function
'==

Public Function DeleteSellType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSellType

   RName = "DeleteSellType"
   DeleteSellType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSellType
   cData.SELL_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSellType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSellType = False
End Function

Public Function AddEditCountry(At As CCountry, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCountry"
   AddEditCountry = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCountry = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCountry = False
End Function

Public Function DeleteCountry(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCountry

   RName = "DeleteCountry"
   DeleteCountry = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCountry
   cData.COUNTRY_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCountry = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCountry = True
End Function

Public Function DeleteAgeRange(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CAgeRange

   RName = "DeleteAgeRange"
   DeleteAgeRange = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CAgeRange
   cData.AGE_RANGE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteAgeRange = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteAgeRange = True
End Function

Public Function DeleteFormulaType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CFormulaType

   RName = "DeleteFormulaType"
   DeleteFormulaType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CFormulaType
   cData.FORMULA_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteFormulaType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteFormulaType = True
End Function

'==
Public Function AddEditCustomerGrade(At As CCustomerGrade, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomerGrade"
   AddEditCustomerGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCustomerGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomerGrade = False
End Function

Public Function DeleteCustomerGrade(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCustomerGrade

   RName = "DeleteCustomerGrade"
   DeleteCustomerGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCustomerGrade
   cData.CSTGRADE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCustomerGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCustomerGrade = True
End Function

'==
Public Function AddEditCustomerType(At As CCustomerType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomerType"
   AddEditCustomerType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCustomerType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomerType = False
End Function

Public Function DeleteCustomerType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCustomerType

   RName = "DeleteCustomerType"
   DeleteCustomerType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCustomerType
   cData.CSTTYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCustomerType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCustomerType = True
End Function
'==
Public Function AddEditSupplierGrade(At As CSupplierGrade, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierGrade"
   AddEditSupplierGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSupplierGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierGrade = False
End Function

Public Function DeleteSupplierGrade(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSupplierGrade

   RName = "DeleteSupplierGrade"
   DeleteSupplierGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSupplierGrade
   cData.SUPPLIER_GRADE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSupplierGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSupplierGrade = True
End Function
'==
Public Function AddEditSupplierType(At As CSupplierType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierType"
   AddEditSupplierType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSupplierType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierType = False
End Function

Public Function DeleteSupplierType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSupplierType

   RName = "DeleteSupplierType"
   DeleteSupplierType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSupplierType
   cData.SUPPLIER_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSupplierType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSupplierType = True
End Function

'==
Public Function AddEditSupplierStatus(At As CSupplierStatus, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierStatus"
   AddEditSupplierStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSupplierStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierStatus = False
End Function

Public Function DeleteSupplierStatus(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSupplierStatus

   RName = "DeleteSupplierStatus"
   DeleteSupplierStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSupplierStatus
   cData.SUPPLIER_STATUS_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSupplierStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSupplierStatus = True
End Function
'==
Public Function AddEditPosition(Ug As CEmpPosition, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPosition"
   AddEditPosition = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPosition = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPosition = False
End Function

Public Function DeletePosition(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CEmpPosition

   RName = "DeletePosition"
   DeletePosition = False
   HasBegin = False

   Set B = New CEmpPosition

   m_ADOConn.BeginTrans
   HasBegin = True

   B.POSITION_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePosition = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePosition = True
End Function

'==
Public Function AddEditUnit(Ug As CUnit, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUnit"
   AddEditUnit = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditUnit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUnit = False
End Function

Public Function DeleteUnit(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CUnit

   RName = "DeleteUnit"
   DeleteUnit = False
   HasBegin = False

   Set B = New CUnit

   m_ADOConn.BeginTrans
   HasBegin = True

   B.UNIT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteUnit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteUnit = True
End Function

Public Function QueryPartGroup(Ua As CPartGroup, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryPartGroup"
   QueryPartGroup = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CHGroupItem
      
      Set Yw = New CHGroupItem
      Yw.HGROUP_ITEM_ID = -1
      Yw.PART_GROUP_ID = Ua.PART_GROUP_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing
      
      Set Ua.PartGroupItems = Nothing
      Set Ua.PartGroupItems = New Collection
      While Not m_Rs.EOF
         Set Yw = New CHGroupItem
         Call Yw.PopulateFromRS(1, m_Rs)
                  
         Yw.Flag = "I"
         Call Ua.PartGroupItems.add(Yw)
         
         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryPartGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPartGroup = False
End Function

Public Function QueryAgeRange(Ua As CAgeRange, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryAgeRange"
   QueryAgeRange = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
   End If
   
   QueryAgeRange = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAgeRange = False
End Function

Public Function QueryStatusGroup(Ua As CStatusGroup, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryStatusGroup"
   QueryStatusGroup = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CSGroupItem
      
      Set Yw = New CSGroupItem
      Yw.SGROUP_ITEM_ID = -1
      Yw.STATUS_GROUP_ID = Ua.STATUS_GROUP_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing
      
      Set Ua.HGroupItems = Nothing
      Set Ua.HGroupItems = New Collection
      While Not m_Rs.EOF
         Set Yw = New CSGroupItem
         Call Yw.PopulateFromRS(1, m_Rs)
                  
         Yw.Flag = "I"
         Call Ua.HGroupItems.add(Yw)
         
         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryStatusGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryStatusGroup = False
End Function

Public Function AddEditAgeRange(At As CAgeRange, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditAgeRange"
   AddEditAgeRange = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditAgeRange = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditAgeRange = False
End Function

Public Function AddEditFormulaType(At As CFormulaType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFormulaType"
   AddEditFormulaType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditFormulaType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditFormulaType = False
End Function

'==
Public Function AddEditReason(Ug As CReason, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReason"
   AddEditReason = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditReason = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReason = False
End Function

Public Function DeleteReason(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CReason

   RName = "DeleteReason"
   DeleteReason = False
   HasBegin = False

   Set B = New CReason

   m_ADOConn.BeginTrans
   HasBegin = True

   B.REASON_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteReason = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteReason = True
End Function

Public Function AddEditLayout(Ug As CLayout, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLayout"
   AddEditLayout = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditLayout = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLayout = False
End Function

Public Function AddEditResource(Ug As CResource, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditResource"
   AddEditResource = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditResource = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditResource = False
End Function

Public Function DeleteLayout(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CLayout

   RName = "DeleteLayout"
   DeleteLayout = False
   HasBegin = False

   Set B = New CLayout

   m_ADOConn.BeginTrans
   HasBegin = True

   B.LAY_OUT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteLayout = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteLayout = True
End Function

Public Function DeleteResource(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CResource

   RName = "DeleteResource"
   DeleteResource = False
   HasBegin = False

   Set B = New CResource

   m_ADOConn.BeginTrans
   HasBegin = True

   B.RESOURCE_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteResource = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteResource = True
End Function

Public Function AddEditDoType(At As CDoType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditDoType"
   AddEditDoType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditDoType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditDoType = False
End Function

Public Function DeleteDoType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CDoType

   RName = "DeleteDoType"
   DeleteDoType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CDoType
   cData.DO_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteDoType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteDoType = False
End Function

Public Function AddEditFeatureType(At As CFeatureType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFeatureType"
   AddEditFeatureType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditFeatureType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditFeatureType = False
End Function

Public Function DeleteFeatureType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CFeatureType

   RName = "DeleteFeatureType"
   DeleteFeatureType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CFeatureType
   cData.FEATURE_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteFeatureType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteFeatureType = False
End Function

Public Function AddEditWorkStatus(At As CWorkStatus, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditWorkStatus"
   AddEditWorkStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditWorkStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditWorkStatus = False
End Function
Public Function AddEditReligious(At As CReligious, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReligious"
   AddEditReligious = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditReligious = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReligious = False
End Function
Public Function AddEditResign(At As CResignReason, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditResign"
   AddEditResign = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditResign = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditResign = False
End Function
Public Function AddEditBankAccount(At As CBankAccount, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBankAccount"
   AddEditBankAccount = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditBankAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBankAccount = False
End Function

Public Function AddEditDocumentType(At As CDocumentType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditDocumentType"
   AddEditDocumentType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditDocumentType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditDocumentType = False
End Function
Public Function AddEditMonthlyAdd(At As CMonthlyAdd, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMonthlyAdd"
   AddEditMonthlyAdd = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditMonthlyAdd = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMonthlyAdd = False
End Function
Public Function AddEditMonthlySub(At As CMonthlySub, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMonthlySub"
   AddEditMonthlySub = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditMonthlySub = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMonthlySub = False
End Function

Public Function DeleteWorkStatus(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CWorkStatus

   RName = "DeleteWorkStatus"
   DeleteWorkStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CWorkStatus
   cData.WORK_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteWorkStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteWorkStatus = True
End Function
Public Function DeleteReligious(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CReligious

   RName = "DeleteReligious"
   DeleteReligious = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CReligious
   cData.RELIGIOUS_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteReligious = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteReligious = True
End Function
Public Function DeleteResign(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CResignReason

   RName = "DeleteResign"
   DeleteResign = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CResignReason
   cData.RSGRESON_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteResign = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteResign = True
End Function
Public Function DeleteBankAccount(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CBankAccount

   RName = "DeleteBankAccount"
   DeleteBankAccount = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CBankAccount
   cData.BANK_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteBankAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBankAccount = True
End Function

Public Function DeleteMonthlyAdd(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CMonthlyAdd

   RName = "DeleteMonthlyAdd"
   DeleteMonthlyAdd = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CMonthlyAdd
   cData.MONTHLY_ADD_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMonthlyAdd = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMonthlyAdd = True
End Function
Public Function DeleteMonthlySub(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CMonthlySub

   RName = "DeleteMonthlySub"
   DeleteMonthlySub = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CMonthlySub
   cData.MONTHLY_SUB_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMonthlySub = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMonthlySub = True
End Function

Public Function DeleteDocumentType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CDocumentType

   RName = "DeleteDocumentType"
   DeleteDocumentType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CDocumentType
   cData.DOCTYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteDocumentType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteDocumentType = True
End Function

Public Function AddEditProcess(Ug As CProcess, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditProcess"
   AddEditProcess = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CParameterItem
   For Each CstName In Ug.ParameterItems
      If CstName.Flag = "A" Then
         CstName.PROCESS_ID = Ug.PROCESS_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName


   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditProcess = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditProcess = False
End Function

Public Function DeleteProcess(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CProcess

   RName = "DeleteProcess"
   DeleteProcess = False
   HasBegin = False

   Set B = New CProcess

   m_ADOConn.BeginTrans
   HasBegin = True

   B.PROCESS_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteProcess = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteProcess = True
End Function

Public Function AddEditMachine(Ug As CMachine, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMachine"
   AddEditMachine = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditMachine = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMachine = False
End Function

Public Function DeleteMachine(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CMachine

   RName = "DeleteMachine"
   DeleteMachine = False
   HasBegin = False

   Set B = New CMachine

   m_ADOConn.BeginTrans
   HasBegin = True

   B.MACHINE_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMachine = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteMachine = True
End Function

Public Function AddEditParameterProcess(At As CParameterProcess, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditParameterProcess"
   AddEditParameterProcess = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   
   Call At.AddEditData
   
'   Dim CstName As CParameterItem
'   For Each CstName In At.ParameterItems
'      If CstName.Flag = "A" Then
'         CstName.PARAMETER_ID = At.PARAMETER_PROCESS_ID
'         CstName.AddEditMode = SHOW_ADD
'         CstName.AddEditData
'      ElseIf CstName.Flag = "E" Then
'         CstName.AddEditMode = SHOW_EDIT
'         CstName.AddEditData
'      ElseIf CstName.Flag = "D" Then
'         CstName.DeleteData
'      End If
'   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditParameterProcess = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditParameterProcess = False
End Function

Public Function QueryParameterProcess(Ua As CParameterProcess, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryParameterProcess"
   QueryParameterProcess = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

'   If Ua.QueryFlag = 1 Then
'      Dim Yw As CParameterItem
'
'      Set Yw = New CParameterItem
'      Yw.PARAMETER_ITEM_ID = -1
'      Yw.PARAMETER_ID = Ua.PARAMETER_PROCESS_ID
'      Yw.OrderBy = 1
'      Call Yw.QueryData(1, m_Rs, iCount)
'      Set Yw = Nothing
'
'      Set Ua.ParameterItems = Nothing
'      Set Ua.ParameterItems = New Collection
'      While Not m_Rs.EOF
'         Set Yw = New CParameterItem
'         Call Yw.PopulateFromRS(1, m_Rs)
'
'         Yw.Flag = "I"
'         Call Ua.ParameterItems.Add(Yw)
'
'         Set Yw = Nothing
'         m_Rs.MoveNext
'      Wend
'   End If
   
   QueryParameterProcess = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryParameterProcess = False
End Function

Public Function DeleteParameterProcess(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CParameterProcess

   RName = "DeleteParameterProcess"
   DeleteParameterProcess = False
   HasBegin = False

   Set B = New CParameterProcess

   m_ADOConn.BeginTrans
   HasBegin = True

   B.PARAMETER_PROCESS_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteParameterProcess = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteParameterProcess = True
End Function

Public Function QueryProcess(Ua As CProcess, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryProcess"
   QueryProcess = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CParameterItem

      Set Yw = New CParameterItem
      Yw.PARAMETER_ITEM_ID = -1
      Yw.PROCESS_ID = Ua.PROCESS_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing

      Set Ua.ParameterItems = Nothing
      Set Ua.ParameterItems = New Collection
      While Not m_Rs.EOF
         Set Yw = New CParameterItem
         Call Yw.PopulateFromRS(1, m_Rs)

         Yw.Flag = "I"
         Call Ua.ParameterItems.add(Yw)

         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryProcess = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryProcess = False
End Function

Public Function DeleteMoneyFamily(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CMoneyFamily

   RName = "DeleteMoneyFamily"
   DeleteMoneyFamily = False
   HasBegin = False

   Set B = New CMoneyFamily

   m_ADOConn.BeginTrans
   HasBegin = True

   B.MONEY_FAMILY_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMoneyFamily = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteMoneyFamily = True
End Function

Public Function AddEditMoneyFamily(Ug As CMoneyFamily, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMoneyFamily"
   AddEditMoneyFamily = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditMoneyFamily = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMoneyFamily = False
End Function

Public Function AddEditBank(At As CBank, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBank"
   AddEditBank = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditBank = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBank = False
End Function

Public Function DeleteBank(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CBank

   RName = "DeleteBank"
   DeleteBank = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CBank
   cData.BANK_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteBank = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBank = True
End Function

Public Function AddEditBankBranch(At As CBankBranch, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBankBranch"
   AddEditBankBranch = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditBankBranch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBankBranch = False
End Function

Public Function DeleteBankBranch(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CBankBranch

   RName = "DeleteBankBranch"
   DeleteBankBranch = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CBankBranch
   cData.BBRANCH_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteBankBranch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBankBranch = True
End Function

Public Function AddEditPackaging(Ug As CPackaging, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPackaging"
   AddEditPackaging = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPackaging = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPackaging = False
End Function

Public Function DeletePackaging(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CPackaging

   RName = "DeletePackaging"
   DeletePackaging = False
   HasBegin = False

   Set B = New CPackaging

   m_ADOConn.BeginTrans
   HasBegin = True

   B.PACKAGING_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePackaging = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePackaging = True
End Function

Public Function AddEditPurchaseExpense(Ug As CPurchaseExpense, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPurchaseExpense"
   AddEditPurchaseExpense = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPurchaseExpense = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPurchaseExpense = False
End Function

Public Function DeletePurchaseExpense(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CPurchaseExpense

   RName = "DeletePurchaseExpense"
   DeletePurchaseExpense = False
   HasBegin = False

   Set B = New CPurchaseExpense

   m_ADOConn.BeginTrans
   HasBegin = True

   B.PUREXP_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePurchaseExpense = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePurchaseExpense = True
End Function

Public Function QueryMasterRef(Ua As CMasterRef, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryMasterRef"
   QueryMasterRef = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CPrtItemSet
      
      Set Yw = New CPrtItemSet
      Call Yw.SetFieldValue("PRTITEM_SET_ID", -1)
      Call Yw.SetFieldValue("MASTER_REF_ID", Ua.KEY_ID)
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing
      
      Set Ua.PrtItemSets = Nothing
      Set Ua.PrtItemSets = New Collection
      While Not m_Rs.EOF
         Set Yw = New CPrtItemSet
         Call Yw.PopulateFromRS(1, m_Rs)
         
         Yw.Flag = "I"
         Call Ua.PrtItemSets.add(Yw)
         
         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryMasterRef = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryMasterRef = False
End Function
Public Function AddEditMasterRef(Ug As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterRef"
   AddEditMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CPrtItemSet
   For Each CstName In Ug.PrtItemSets
      If CstName.Flag = "A" Then
         Call CstName.SetFieldValue("MASTER_REF_ID", Ug.KEY_ID)
         CstName.ShowMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.ShowMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMasterRef = True
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
   AddEditMasterRef = False
End Function
Public Function DeleteMasterRef(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CMasterRef

   RName = "DeleteMasterRef"
   DeleteMasterRef = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CMasterRef
   cData.KEY_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterRef = True
End Function

