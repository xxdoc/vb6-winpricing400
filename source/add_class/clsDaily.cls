VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
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
Private m_Rs4 As ADODB.Recordset
Private m_Rs5 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
   Set m_Rs4 = New ADODB.Recordset
   Set m_Rs5 = New ADODB.Recordset
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
   
   If m_Rs4.State = adStateOpen Then
      m_Rs4.Close
   End If
   Set m_Rs4 = Nothing
   
   If m_Rs5.State = adStateOpen Then
      m_Rs5.Close
   End If
   Set m_Rs5 = Nothing
End Sub
'===
Public Function QueryYearSeq(Ua As CYearSeq, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim Exp As CXItem

   RName = "QueryYearSeq"
   QueryYearSeq = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim Yw As CYearWeek
      
      Set Yw = New CYearWeek
      Yw.YEAR_WEEK_ID = -1
      Yw.YEAR_SEQ_ID = Ua.YEAR_SEQ_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.YearWeeks = Nothing
      Set Ua.YearWeeks = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CYearWeek
         Call Yw.PopulateFromRS(1, m_Rs1)
         
         If Yw.PART_ITEM_ID1 > 0 Then
            Yw.PartItem1.PART_ITEM_ID = Yw.PART_ITEM_ID1
            Call Yw.PartItem1.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Yw.PartItem1.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         If Yw.PART_ITEM_ID2 > 0 Then
            Yw.PartItem2.PART_ITEM_ID = Yw.PART_ITEM_ID2
            Call Yw.PartItem2.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Yw.PartItem2.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         If Yw.PART_ITEM_ID3 > 0 Then
            Yw.PartItem3.PART_ITEM_ID = Yw.PART_ITEM_ID3
            Call Yw.PartItem3.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Yw.PartItem3.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         If Yw.PART_ITEM_ID4 > 0 Then
            Yw.PartItem4.PART_ITEM_ID = Yw.PART_ITEM_ID4
            Call Yw.PartItem4.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Yw.PartItem4.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         If Yw.PART_ITEM_ID5 > 0 Then
            Yw.PartItem5.PART_ITEM_ID = Yw.PART_ITEM_ID5
            Call Yw.PartItem5.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Yw.PartItem5.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         If Yw.PART_ITEM_ID6 > 0 Then
            Yw.PartItem6.PART_ITEM_ID = Yw.PART_ITEM_ID6
            Call Yw.PartItem6.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Yw.PartItem6.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Yw.Flag = "I"
         Call Ua.YearWeeks.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryYearSeq = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryYearSeq = False
End Function

Public Function AddEditYearSeq(Ug As CYearSeq, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditYearSeq"
   AddEditYearSeq = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Yw As CYearWeek
   For Each Yw In Ug.YearWeeks
      If Yw.Flag = "A" Then
         Yw.PartItem1.AddEditMode = SHOW_ADD
         Yw.PartItem1.AddEditData
         
         Yw.PartItem2.AddEditMode = SHOW_ADD
         Yw.PartItem2.AddEditData
         
         Yw.PartItem3.AddEditMode = SHOW_ADD
         Yw.PartItem3.AddEditData
         
         Yw.PartItem4.AddEditMode = SHOW_ADD
         Yw.PartItem4.AddEditData
         
         Yw.PartItem5.AddEditMode = SHOW_ADD
         Yw.PartItem5.AddEditData
         
         Yw.PartItem6.AddEditMode = SHOW_ADD
         Yw.PartItem6.AddEditData
         
         Yw.YEAR_SEQ_ID = Ug.YEAR_SEQ_ID
         Yw.PART_ITEM_ID1 = Yw.PartItem1.PART_ITEM_ID
         Yw.PART_ITEM_ID2 = Yw.PartItem2.PART_ITEM_ID
         Yw.PART_ITEM_ID3 = Yw.PartItem3.PART_ITEM_ID
         Yw.PART_ITEM_ID4 = Yw.PartItem4.PART_ITEM_ID
         Yw.PART_ITEM_ID5 = Yw.PartItem5.PART_ITEM_ID
         Yw.PART_ITEM_ID6 = Yw.PartItem6.PART_ITEM_ID
         Yw.AddEditMode = SHOW_ADD
         Yw.AddEditData
      ElseIf Yw.Flag = "E" Then
         Yw.PartItem1.AddEditMode = SHOW_EDIT
         Yw.PartItem1.AddEditData
         
         Yw.PartItem2.AddEditMode = SHOW_EDIT
         Yw.PartItem2.AddEditData
         
         Yw.PartItem3.AddEditMode = SHOW_EDIT
         Yw.PartItem3.AddEditData
         
         Yw.PartItem4.AddEditMode = SHOW_EDIT
         Yw.PartItem4.AddEditData
         
         Yw.PartItem5.AddEditMode = SHOW_EDIT
         Yw.PartItem5.AddEditData
         
         Yw.PartItem6.AddEditMode = SHOW_EDIT
         Yw.PartItem6.AddEditData
         
         Yw.AddEditMode = SHOW_EDIT
         Yw.AddEditData
      ElseIf Yw.Flag = "D" Then
         Call Yw.DeleteData
         
         Call Yw.PartItem1.DeleteData
         Call Yw.PartItem2.DeleteData
         Call Yw.PartItem3.DeleteData
         Call Yw.PartItem4.DeleteData
         Call Yw.PartItem5.DeleteData
         Call Yw.PartItem6.DeleteData
      End If
   Next Yw

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditYearSeq = True
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
   AddEditYearSeq = False
End Function

Public Function DeleteYearSeq(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CYearSeq

   RName = "DeleteYearSeq"
   DeleteYearSeq = False
   HasBegin = False

   Set B = New CYearSeq

   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.YEAR_SEQ_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteYearSeq = True
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

   DeleteYearSeq = True
End Function

'===
Public Function QueryPartItem(Ua As CPartItem, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim TempSuppliers As Collection

   RName = "QueryPartItem"
   QueryPartItem = False

   Set TempSuppliers = New Collection
   Call LoadSupplier(Nothing, TempSuppliers)
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CPartLocation
      
      Set Yw = New CPartLocation
      Yw.PART_LOCATION_ID = -1
      Yw.PART_ITEM_ID = Ua.PART_ITEM_ID
      Yw.LOCATION_ID = -1
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.PartLocations = Nothing
      Set Ua.PartLocations = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CPartLocation
         Call Yw.PopulateFromRS(m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.PartLocations.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   
   
      Dim Li As CLotItem
      Dim Sp As CSupplier
      
      Set Li = New CLotItem
      Li.LOT_ITEM_ID = -1
      Li.PART_ITEM_ID = Ua.PART_ITEM_ID
      Li.TX_TYPE = "I"
      Li.LOCATION_ID = -1
      Li.OrderBy = 1
      Call Li.QueryData(7, m_Rs1, iCount)
      Set Li = Nothing
      
      Set Ua.Suppliers = Nothing
      Set Ua.Suppliers = New Collection
      While Not m_Rs1.EOF
         Set Li = New CLotItem
         Call Li.PopulateFromRS(7, m_Rs1)
         
         If Li.SUPPLIER_ID > 0 Then
            Set Sp = TempSuppliers(Trim(str(Li.SUPPLIER_ID)))
            Sp.TOTAL_INCLUDE_PRICE = Li.TOTAL_INCLUDE_PRICE
            Sp.TX_AMOUNT = Li.TX_AMOUNT
            Call Ua.Suppliers.add(Sp)
         End If
         
         Set Li = Nothing
         m_Rs1.MoveNext
      Wend
   
      Dim Ps As CPartItemSpec
      
      Set Ps = New CPartItemSpec
      Ps.PARTITEM_SPEC_ID = -1
      Ps.PART_ITEM_ID = Ua.PART_ITEM_ID
      Ps.OrderBy = 1
      Call Ps.QueryData(1, m_Rs1, iCount)
      Set Ps = Nothing

      Set Ua.HumidRates = Nothing
      Set Ua.HumidRates = New Collection
      While Not m_Rs1.EOF
         Set Ps = New CPartItemSpec
         Call Ps.PopulateFromRS(1, m_Rs1)

         Ps.Flag = "I"
         Call Ua.HumidRates.add(Ps)
         
         Set Ps = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   Dim CPP As CPartItemPicture
   Set CPP = New CPartItemPicture
   Call CPP.SetFieldValue("PART_ITEM_ID", Ua.PART_ITEM_ID)
   Call CPP.QueryData(1, m_Rs1, iCount)
   Set CPP = Nothing

   Set Ua.Pictures = Nothing
   Set Ua.Pictures = New Collection

   While Not m_Rs1.EOF
      Set CPP = New CPartItemPicture
      Call CPP.PopulateFromRS(1, m_Rs1)
      
      Call Ua.Pictures.add(CPP)
      Set CPP = Nothing

      m_Rs1.MoveNext
   Wend
   
   Set TempSuppliers = Nothing
   QueryPartItem = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPartItem = False
End Function
'===
Public Function QueryPartMaster(Ua As CPartMaster, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryPartMaster"
   QueryPartMaster = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Pi As CPartItem
      Set Pi = New CPartItem
      Pi.PART_ITEM_ID = -1
      Pi.PART_MASTER_ID = Ua.PART_MASTER_ID
      Pi.OrderBy = 1
      Call Pi.QueryData(1, m_Rs1, iCount)
      Set Pi = Nothing
      
      Set Ua.PartItems = Nothing
      Set Ua.PartItems = New Collection
      While Not m_Rs1.EOF
         Set Pi = New CPartItem
         Call Pi.PopulateFromRS(1, m_Rs1)
         Pi.Flag = "I"
         Call Ua.PartItems.add(Pi)
         Set Pi = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim Cg As CCusGroup
      Set Cg = New CCusGroup
      Cg.CUS_GROUPS_ID = -1
      Cg.PART_MASTER_ID = Ua.PART_MASTER_ID
      Cg.OrderBy = 1
      Call Cg.QueryData(1, m_Rs1, iCount)
      Set Cg = Nothing
      Set Ua.CusGroups = Nothing
      Set Ua.CusGroups = New Collection
      While Not m_Rs1.EOF
         Set Cg = New CCusGroup
         Call Cg.PopulateFromRS(1, m_Rs1)
         Cg.Flag = "I"
         Call Ua.CusGroups.add(Cg)
         Set Cg = Nothing
         m_Rs1.MoveNext
      Wend
   End If

   QueryPartMaster = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPartMaster = False
End Function
Public Function QueryLotItemWh(Ua As CLotItemWH, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim TempSuppliers As Collection

   RName = "QueryLotItemWH"
   QueryLotItemWh = False
   
   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)
   QueryLotItemWh = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLotItemWh = False
End Function
Public Function QueryLotItemWhByID(Ua As CLotItemWH, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim TempSuppliers As Collection

   RName = "QueryLotItemWhByID"
   QueryLotItemWhByID = False
   
   IsOK = True
   Call Ua.QueryData(2, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim LTD As CLotDoc
      Set LTD = New CLotDoc
      LTD.LOT_DOC_ID = -1
      LTD.LOT_ITEM_WH_ID = Ua.LOT_ITEM_WH_ID
      LTD.OrderBy = -1
      LTD.OrderType = 1
      Call LTD.QueryData(7, m_Rs1, iCount)
      Set LTD = Nothing
      
      Set Ua.C_LotDoc = Nothing
      Set Ua.C_LotDoc = New Collection
      While Not m_Rs1.EOF
         Set LTD = New CLotDoc
         Call LTD.PopulateFromRS(7, m_Rs1)
         
'         'get pallet
'         Dim PD As CPalletDoc
'         Set PD = New CPalletDoc
'         PD.PALLET_DOC_ID = -1
'         PD.LOT_DOC_ID = LTD.LOT_DOC_ID
'         PD.OrderBy = -1
'         PD.OrderType = 1
'         Call PD.QueryData(1, m_Rs2, iCount)
'         Set PD = Nothing
'
'         Set LTD.C_PalletDoc = Nothing
'         Set LTD.C_PalletDoc = New Collection
'         While Not m_Rs2.EOF
'            Set PD = New CPalletDoc
'            Call PD.PopulateFromRS(1, m_Rs2)
'
'            PD.Flag = "I"
'            Call LTD.C_PalletDoc.add(PD)
'
'            Set PD = Nothing
'            m_Rs2.MoveNext
'         Wend
'         '
                  
         LTD.Flag = "I"
         Call Ua.C_LotDoc.add(LTD)
         
         Set LTD = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryLotItemWhByID = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLotItemWhByID = False
End Function
Public Function QueryLotItemWhAge(Ua As CLotItemWH, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim TempSuppliers As Collection

   RName = "QueryLotItemWhAge"
   QueryLotItemWhAge = False
   
   IsOK = True
   Call Ua.QueryData(13, Rs, ItemCount)
   QueryLotItemWhAge = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLotItemWhAge = False
End Function
Public Function QueryLotItemWhPart(Ua As CLotItemWH, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog, Optional ByRef TotalPart As Double) As Boolean
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
Dim TempSuppliers As Collection
Dim LWH As CLotItemWH
Dim m_CollPallet As Collection
Dim Key As String
Dim CLotI As Collection
Dim CLotE As Collection
Dim LTD As CLotDoc

   RName = "QueryLotItemWhPart"
   QueryLotItemWhPart = False
   
   Set CLotI = New Collection
   Set CLotE = New Collection


   
   IsOK = True
   Call Ua.QueryData(7, Rs, ItemCount)
   If Ua.QueryFlag = 1 Then
   
      Call LoadLotFromLotDoc(Nothing, CLotI, , , , 1, Ua.PART_ITEM_ID, 3, , , "I", , , Ua.LOCATION_ID)
      Call LoadLotFromLotDoc(Nothing, CLotE, , , , 1, Ua.PART_ITEM_ID, 10, , , "E", , , Ua.LOCATION_ID)
   
      'LotDoc ++++++++++++++++++++++++++
      Dim TempData As CLotDoc
      Set TempData = New CLotDoc
      TempData.PART_ITEM_ID = Ua.PART_ITEM_ID
      TempData.LOCATION_ID = Ua.LOCATION_ID
      TempData.TX_TYPE = "I"
      TempData.BALANCE_FLAG = "N" 'เลือกเฉพาะที่ยังไม่ปรับยอด
      TempData.OUT_STOCK_FLAG = "N" ' เลือกเฉพาะตัวที่ยังมีเหลือใน stock
      Call TempData.QueryData(6, m_Rs2, iCount)
      Set TempData = Nothing
      
      If Ua.C_LotDoc Is Nothing Then
         Set Ua.C_LotDoc = Nothing
         Set Ua.C_LotDoc = New Collection
      End If
      While Not m_Rs2.EOF
         Set TempData = New CLotDoc
         Call TempData.PopulateFromRS(6, m_Rs2)
         
         If TempData.LOT_NO = "LG611021232" Then
         Debug.Print
         End If
      Set LTD = GetObject("CLotDoc", CLotI, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "I"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = LTD.LOT_AMOUNT
      End If
      Set LTD = GetObject("CLotDoc", CLotE, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "E"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = TempData.LOT_AMOUNT - LTD.LOT_AMOUNT
      End If
         
         
         TempData.Flag = "I"
         
         TotalPart = TotalPart + TempData.LOT_AMOUNT
'         Set m_CollPallet = New Collection
'         Call LoadPalletDocAmount(Nothing, m_CollPallet, LTD.LOT_ID, 2, , 2, "I", , , LTD.LOT_DOC_ID, , LTD.LOT_ITEM_WH_ID, Ua.DOCUMENT_TYPE, Ua.PART_ITEM_ID, LTD.LOT_DOC_ID)
         If Val(TempData.LOT_AMOUNT) > 0 Then 'ถ้าไม่มีข้อมูลใน pallet แล้วก็ไม่ต้องให้แสดง Lot
         
         Dim PD As CPalletDoc
         Set PD = New CPalletDoc
         PD.LOT_DOC_ID = TempData.LOT_DOC_ID
         PD.OrderBy = 1
         PD.BALANCE_FLAG = "N" 'กรณีดูยอดคงเหลือ ให้เอาเฉพาะตัวที่เป็น N เท่านั้น
         Call PD.QueryData(2, m_Rs3, iCount)
         Set PD = Nothing
            While Not m_Rs3.EOF
               Set PD = New CPalletDoc
               Call PD.PopulateFromRS(2, m_Rs3)
               PD.Flag = "I"
                 Call TempData.C_PalletDoc.add(PD)
               Set PD = Nothing
               m_Rs3.MoveNext
            Wend
         Call Ua.C_LotDoc.add(TempData)
         
         Set TempData = Nothing
          End If
         m_Rs2.MoveNext
      Wend
      'LotDoc ++++++++++++++++++++++++++
   End If
   QueryLotItemWhPart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLotItemWhPart = False
End Function
Public Function QueryLotItemWhDistinctPart(Ua As CLotItemWH, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim TempSuppliers As Collection
Dim LWH As CLotItemWH

   RName = "QueryLotItemWhDistinctPart"
   QueryLotItemWhDistinctPart = False
   
   IsOK = True
   Call Ua.QueryData(14, Rs, ItemCount)
   QueryLotItemWhDistinctPart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLotItemWhDistinctPart = False
End Function
Public Function QueryLotItemWhPart2(Ua As CLotItemWH, Cl As Collection, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim TempSuppliers As Collection
Dim LWH As CLotItemWH
   RName = "QueryLotItemWhPart2"
   QueryLotItemWhPart2 = False
   
   IsOK = True
   Call Ua.QueryData(9, m_Rs1, ItemCount)
   
 While Not m_Rs1.EOF
 Set LWH = New CLotItemWH
 Set Cl = New Collection
 Call LWH.PopulateFromRS(9, m_Rs1)
 LWH.Flag = "I"
 
   If Ua.QueryFlag = 1 Then
      'LotDoc ++++++++++++++++++++++++++
      Dim LTD As CLotDoc
      Set LTD = New CLotDoc
      LTD.PART_ITEM_ID = Ua.PART_ITEM_ID
      LTD.TX_TYPE = "I"
      Call LTD.QueryData(6, m_Rs2, iCount)
      Set LTD = Nothing
      
      If Ua.C_LotDoc Is Nothing Then
         Set Ua.C_LotDoc = Nothing
         Set Ua.C_LotDoc = New Collection
      End If
      While Not m_Rs2.EOF
         Set LTD = New CLotDoc
         Call LTD.PopulateFromRS(6, m_Rs2)
         LTD.Flag = "I"
        
         Dim PD As CPalletDoc
         Set PD = New CPalletDoc
         PD.LOT_DOC_ID = LTD.LOT_DOC_ID
         PD.OrderBy = 1
         Call PD.QueryData(2, m_Rs3, iCount)
         Set PD = Nothing
            While Not m_Rs3.EOF
               Set PD = New CPalletDoc
               Call PD.PopulateFromRS(2, m_Rs3)
               PD.Flag = "I"
               Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO))
               Set PD = Nothing
               m_Rs3.MoveNext
            Wend
         Call LWH.C_LotDoc.add(LTD)
'         Call LWH.C_LotDoc.add(LTD, Trim(str(LTD.LOT_ID)))
         Set LTD = Nothing
         m_Rs2.MoveNext
      Wend
      'LotDoc ++++++++++++++++++++++++++
   End If
   Call Cl.add(LWH, Trim(str(LWH.PART_ITEM_ID)))
   Set LWH = Nothing
   m_Rs1.MoveNext
Wend
   QueryLotItemWhPart2 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLotItemWhPart2 = False
End Function
Public Function DeletePartItem(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CPartItem

   RName = "DeletePartItem"
   DeletePartItem = False
   HasBegin = False

   Set B = New CPartItem

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PART_ITEM_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePartItem = True
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

   DeletePartItem = True
End Function
Public Function DeletePartMaster(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CPartMaster

   RName = "DeletePartMaster"
   DeletePartMaster = False
   HasBegin = False

   Set B = New CPartMaster

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PART_MASTER_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePartMaster = True
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

   DeletePartMaster = True
End Function
Public Function AddEditPartItem(Ug As CPartItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartItem"
   AddEditPartItem = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CPartItemSpec
   For Each CstName In Ug.HumidRates
      If CstName.Flag = "A" Then
         CstName.PART_ITEM_ID = Ug.PART_ITEM_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Dim PP As CPartItemPicture
   For Each PP In Ug.Pictures
      If PP.Flag = "A" Then
         Call PP.SetFieldValue("PART_ITEM_ID", Ug.PART_ITEM_ID)
         PP.ShowMode = SHOW_ADD
         
         Call PP.SetFieldValue("PART_ITEM_PICTURE_PATH", SavePictureToServer(PP.GetFieldValue("PART_ITEM_PICTURE_PATH")))
         PP.AddEditData
      ElseIf PP.Flag = "E" Then
         PP.ShowMode = SHOW_EDIT
         PP.AddEditData
      ElseIf PP.Flag = "D" Then
         PP.DeleteData
         Call DeletePictureFromDisk(PP.GetFieldValue("PART_ITEM_PICTURE_PATH"))
      End If
   Next PP

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPartItem = True
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
   AddEditPartItem = False
End Function
Public Function AddEditPartMaster(Ug As CPartMaster, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartMaster"
   AddEditPartMaster = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CCusGroup As CCusGroup
   For Each CCusGroup In Ug.CusGroups
      If CCusGroup.Flag = "A" Then
         CCusGroup.PART_MASTER_ID = Ug.PART_MASTER_ID
         CCusGroup.AddEditMode = SHOW_ADD
         CCusGroup.AddEditData
      ElseIf CCusGroup.Flag = "E" Then
         CCusGroup.AddEditMode = SHOW_EDIT
         CCusGroup.AddEditData
      ElseIf CCusGroup.Flag = "D" Then
         CCusGroup.DeleteData
      End If
   Next CCusGroup
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPartMaster = True
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
   AddEditPartMaster = False
End Function
Public Function AddEditPallet(Ug As CPalletDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPallet"
   AddEditPallet = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

'   Dim CPallet As CPalletDoc
'   For Each CPallet In Ug.HumidRates
'      If CstName.Flag = "A" Then
'         CstName.PART_ITEM_ID = Ug.PART_ITEM_ID
'         CstName.AddEditMode = SHOW_ADD
'         CstName.AddEditData
'      ElseIf CstName.Flag = "E" Then
'         CstName.AddEditMode = SHOW_EDIT
'         CstName.AddEditData
'      ElseIf CstName.Flag = "D" Then
'         CstName.DeleteData
'      End If
'   Next CstName

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPallet = True
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
   AddEditPallet = False
End Function
'==
Public Function QueryEnterprise(Ua As CEnterprise, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryEnterprise"
   QueryEnterprise = False
   
   IsOK = True
   
   Call Ua.QueryData(Rs, ItemCount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Contact ++++++++++++++++++++++++++
      Dim cCstContact As CEnterprisePerson
      Set cCstContact = New CEnterprisePerson
      cCstContact.ENTERPRISE_PERSON_ID = -1
      cCstContact.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cCstContact.OrderBy = 1
      Call cCstContact.QueryData(m_Rs1, iCount)
      Set cCstContact = Nothing

      Set Ua.EnpPersons = Nothing
      Set Ua.EnpPersons = New Collection
      While Not m_Rs1.EOF
         Set cCstContact = New CEnterprisePerson
         Call cCstContact.PopulateFromRS(m_Rs1)

         Dim N As CName
         Set N = New CName
         N.NAME_ID = cCstContact.NAME_ID
         Call N.QueryData(m_Rs2, iCount)
         Set N = Nothing
         While Not m_Rs2.EOF
            Set N = New CName
            Call N.PopulateFromRS(m_Rs2)
            N.Flag = "I"
            Set cCstContact.NAME = N
            Set N = Nothing

            m_Rs2.MoveNext
         Wend

         If cCstContact.ADDRESS_ID > 0 Then
            Dim A As CAddress
            Set A = New CAddress
            A.ADDRESS_ID = cCstContact.ADDRESS_ID
            Call A.QueryData(m_Rs2, iCount)
            Set A = Nothing
            If Not m_Rs2.EOF Then
               Set A = New CAddress
               Call A.PopulateFromRS(m_Rs2)
               A.Flag = "I"
               Set cCstContact.Address = A
               Set A = Nothing
            End If
         End If

         cCstContact.Flag = "I"
         Call Ua.EnpPersons.add(cCstContact)
         Set cCstContact = Nothing
         m_Rs1.MoveNext
      Wend
'      Coontact ++++++++++++++++++++++++++
         
      'Name ++++++++++++++++++++++++++
      Dim cEnpname As CEnterpriseName
      Set cEnpname = New CEnterpriseName
      cEnpname.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cEnpname.OrderBy = 1
      Call cEnpname.QueryData(m_Rs1, iCount)
      Set cEnpname = Nothing
      
      Set Ua.EnpNames = Nothing
      Set Ua.EnpNames = New Collection
      While Not m_Rs1.EOF
         Set cEnpname = New CEnterpriseName
         Call cEnpname.PopulateFromRS(m_Rs1)

         Dim CName As CName
         Dim iCount2 As Long
         Set CName = New CName
         CName.NAME_ID = cEnpname.NAME_ID
         Call CName.QueryData(m_Rs2, iCount2)
         Set CName = Nothing
         While Not m_Rs2.EOF
            Set CName = New CName
            Call CName.PopulateFromRS(m_Rs2)
            CName.Flag = "I"
            Call cEnpname.Names.add(CName)
            Set CName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpname.Flag = "I"
         Call Ua.EnpNames.add(cEnpname)
         Set cEnpname = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
      
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CEnterpriseAddress
      Set cEnpAddr = New CEnterpriseAddress
      cEnpAddr.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.EnpAddresses = Nothing
      Set Ua.EnpAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CEnterpriseAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            Call cEnpAddr.Addresses.add(cAddr)
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.EnpAddresses.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If

   QueryEnterprise = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryEnterprise = False
End Function

Public Function AddEditEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim NAME As CName
Dim Addr1 As CAddress

   RName = "AddEditEnterprise"
   AddEditEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
         
   Dim EnpName As CEnterpriseName
   Dim EmpAddr As CEnterpriseAddress
   Dim Addr As CAddress
   
   For Each EmpAddr In Ug.EnpAddresses
      If EmpAddr.Flag = "A" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData
                  
         EmpAddr.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EmpAddr.ADDRESS_ID = Addr.ADDRESS_ID
         EmpAddr.AddEditMode = SHOW_ADD
         EmpAddr.AddEditData
      ElseIf EmpAddr.Flag = "E" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData
         
      ElseIf EmpAddr.Flag = "D" Then
         Call EmpAddr.DeleteData
         
         Set Addr = EmpAddr.Addresses(1)
         Call Addr.DeleteData
      End If
   Next EmpAddr
   
   For Each EnpName In Ug.EnpNames
      If EnpName.Flag = "A" Then
         Set NAME = EnpName.Names(1)
         NAME.AddEditMode = SHOW_ADD
         Call NAME.AddEditData
                  
         EnpName.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EnpName.NAME_ID = NAME.NAME_ID
         EnpName.AddEditMode = SHOW_ADD
         EnpName.AddEditData
      ElseIf EnpName.Flag = "E" Then
         Set NAME = EnpName.Names(1)
         NAME.AddEditMode = SHOW_EDIT
         Call NAME.AddEditData
      End If
   Next EnpName
   
   Dim EnpPerson As CEnterprisePerson
   For Each EnpPerson In Ug.EnpPersons
      If EnpPerson.Flag = "A" Then
         Set NAME = EnpPerson.NAME
         NAME.AddEditMode = SHOW_ADD
         Call NAME.AddEditData
                  
         EnpPerson.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EnpPerson.NAME_ID = NAME.NAME_ID
         EnpPerson.AddEditMode = SHOW_ADD
         EnpPerson.AddEditData
      ElseIf EnpPerson.Flag = "E" Then
         Set NAME = EnpPerson.NAME
         NAME.AddEditMode = SHOW_EDIT
         Call NAME.AddEditData
      
         EnpPerson.AddEditMode = SHOW_EDIT
         EnpPerson.AddEditData
      ElseIf EnpPerson.Flag = "D" Then
         Call EnpPerson.DeleteData

         Set NAME = EnpPerson.NAME
         Call NAME.DeleteData
      End If
   Next EnpPerson
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditEnterprise = True
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
   AddEditEnterprise = False
End Function

Public Function QueryCustomerInfo(Ua As CCustomer, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCustomerInfo"
   QueryCustomerInfo = False

   IsOK = True
   Call Ua.QueryData1(Rs, ItemCount)
         
   QueryCustomerInfo = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomerInfo = False
End Function

'==
Public Function QueryCustomer(Ua As CCustomer, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCustomer"
   QueryCustomer = False
   
   IsOK = True
   Ua.OrderType = 1
   Ua.OrderBy = 1
   Call Ua.QueryData2(Rs, ItemCount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CCustomerName
      Set cCstName = New CCustomerName
      cCstName.CUSTOMER_NAME_ID = -1
      cCstName.CUSTOMER_ID = Ua.CUSTOMER_ID
      cCstName.OrderBy = 1
      Call cCstName.QueryData(m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CCustomerName
         Call cCstName.PopulateFromRS(m_Rs1)

         Dim CName As CName
         Dim iCount2 As Long
         Set CName = New CName
         CName.NAME_ID = cCstName.NAME_ID
         Call CName.QueryData(m_Rs2, iCount2)
         Set CName = Nothing
         While Not m_Rs2.EOF
            Set CName = New CName
            Call CName.PopulateFromRS(m_Rs2)
            CName.Flag = "I"
            Set cCstName.NAME = CName
            Set CName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CCustomerAddress
      Set cEnpAddr = New CCustomerAddress
      cEnpAddr.CUSTOMER_ID = Ua.CUSTOMER_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddr = Nothing
      Set Ua.CstAddr = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CCustomerAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            cAddr.ADDRESS_TYPE = cEnpAddr.ADDRESS_TYPE
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddr.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   
      'Account ++++++++++++++++++++++++++
      Dim cAcc As CAccount
      Set cAcc = New CAccount
      cAcc.ACCOUNT_ID = -1
      cAcc.CUSTOMER_ID = Ua.CUSTOMER_ID
      cAcc.ENABLE_FLAG = "Y"
      cAcc.OrderBy = 1
      Call cAcc.QueryData(2, m_Rs1, iCount)
      Set cAcc = Nothing
      
      Dim cSub As CSubscriber
      Dim cAgr As CAgreement
      
      Set Ua.CstAccounts = Nothing
      Set Ua.CstAccounts = New Collection
      While Not m_Rs1.EOF
         Set cAcc = New CAccount
         Call cAcc.PopulateFromRS(2, m_Rs1)
      
         cAcc.Flag = "I"
         Call Ua.CstAccounts.add(cAcc)
         
         Set cSub = New CSubscriber
         cSub.ACCOUNT_ID = cAcc.ACCOUNT_ID
         cSub.SUBSCRIBER_ID = -1
         Call cSub.QueryData(m_Rs2, iCount)
         Set cSub = Nothing
         Set cAcc.ActSubs = Nothing
         Set cAcc.ActSubs = New Collection
         While Not m_Rs2.EOF
            Set cSub = New CSubscriber
            Call cSub.PopulateFromRS(m_Rs2)
            cSub.Flag = "I"
            
            Call cAcc.ActSubs.add(cSub)
            Set cSub = Nothing
            m_Rs2.MoveNext
         Wend
         
         Set cAgr = New CAgreement
         cAgr.ACCOUNT_ID = cAcc.ACCOUNT_ID
         cAgr.SUBSCRIBER_ID = -1
         cAgr.AGREEMENT_ID = -1
         Call cAgr.QueryData(m_Rs2, iCount)
         Set cAgr = Nothing
         Set cAcc.ActAgrmnts = Nothing
         Set cAcc.ActAgrmnts = New Collection
         While Not m_Rs2.EOF
            Set cAgr = New CAgreement
            Call cAgr.PopulateFromRS(m_Rs2)
            cAgr.Flag = "I"
            
            Call cAcc.ActAgrmnts.add(cAgr)
            Set cAgr = Nothing
            m_Rs2.MoveNext
         Wend
         
         Set cAcc = Nothing
         m_Rs1.MoveNext
      Wend
      'Account ++++++++++++++++++++++++++
      
      Dim CDP As CCustomerPicture
      Set CDP = New CCustomerPicture
      Call CDP.SetFieldValue("CUSTOMER_ID", Ua.CUSTOMER_ID)
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing

      Set Ua.CstPicture = Nothing
      Set Ua.CstPicture = New Collection

      While Not m_Rs1.EOF
         Set CDP = New CCustomerPicture
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         Call Ua.CstPicture.add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend

   
      Dim Acl As CCustomerAccountList
      Set Acl = New CCustomerAccountList
      Call Acl.SetFieldValue("CUSTOMER_ID", Ua.CUSTOMER_ID)
      Call Acl.QueryData(1, m_Rs1, iCount)
      Set Acl = Nothing

      Set Ua.CstAccountList1s = Nothing
      Set Ua.CstAccountList2s = Nothing
      Set Ua.CstAccountList3s = Nothing
      Set Ua.CstAccountList1s = New Collection
      Set Ua.CstAccountList2s = New Collection
      Set Ua.CstAccountList3s = New Collection
      
      While Not m_Rs1.EOF
         Set Acl = New CCustomerAccountList
         Call Acl.PopulateFromRS(1, m_Rs1)
         If Acl.GetFieldValue("ACCOUNT_LIST_TYPE") = 1 Then
            Call Ua.CstAccountList1s.add(Acl)
         ElseIf Acl.GetFieldValue("ACCOUNT_LIST_TYPE") = 2 Then
            Call Ua.CstAccountList2s.add(Acl)
         ElseIf Acl.GetFieldValue("ACCOUNT_LIST_TYPE") = 3 Then
            Call Ua.CstAccountList3s.add(Acl)
         End If
         Set Acl = Nothing

         m_Rs1.MoveNext
      Wend
      
'      Dim AccFol As CCustomerAccFol
'      Set AccFol = New CCustomerAccFol
'      Call AccFol.SetFieldValue("CUSTOMER_ID", Ua.CUSTOMER_ID)
'      Call AccFol.QueryData(1, m_Rs1, iCount)
'      Set AccFol = Nothing
'
'      Set Ua.CstAccFol = Nothing
'      Set Ua.CstAccFol = New Collection
'
'      While Not m_Rs1.EOF
'         Set AccFol = New CCustomerAccFol
'         Call AccFol.PopulateFromRS(1, m_Rs1)
'
'         Call Ua.CstAccFol.add(AccFol)
'         Set AccFol = Nothing
'         m_Rs1.MoveNext
'      Wend
      
      
     Dim AccFol As CAccFol
      Set AccFol = New CAccFol
      AccFol.CUSTOMER_ID = Ua.CUSTOMER_ID
      AccFol.OrderBy = 2
      Call AccFol.QueryData(m_Rs1, iCount)
      Set AccFol = Nothing
      
      Set Ua.CstAccFol = Nothing
      Set Ua.CstAccFol = New Collection
      While Not m_Rs1.EOF
         Set AccFol = New CAccFol
         Call AccFol.PopulateFromRS(m_Rs1)
      
         AccFol.Flag = "I"
         Call Ua.CstAccFol.add(AccFol)
         Set AccFol = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim MKTFol As CMKTFol
      Set MKTFol = New CMKTFol
      MKTFol.CUSTOMER_ID = Ua.CUSTOMER_ID
 '     EE.OrderBy = 1
      Call MKTFol.QueryData(m_Rs1, iCount)
      Set MKTFol = Nothing
      
      Set Ua.CstMKTFol = Nothing
      Set Ua.CstMKTFol = New Collection
      While Not m_Rs1.EOF
         Set MKTFol = New CMKTFol
         Call MKTFol.PopulateFromRS(m_Rs1)
      
         MKTFol.Flag = "I"
         Call Ua.CstMKTFol.add(MKTFol)
         Set MKTFol = Nothing
         m_Rs1.MoveNext
      Wend
 
      Dim MFreelanceItem As CFreelanceItem
      Set MFreelanceItem = New CFreelanceItem
      Dim MFreelance As CFreelance
      Set MFreelance = New CFreelance
      'Call CFreelanceItem.SetFieldValue("CUSTOMER_ID", Ua.CUSTOMER_ID)
      MFreelanceItem.CUSTOMER_ID = Ua.CUSTOMER_ID
      Call MFreelanceItem.QueryData(m_Rs1, iCount)
      Set MFreelanceItem = Nothing
      
      
      Set Ua.CstFreelance = Nothing
      Set Ua.CstFreelance = New Collection
     
      While Not m_Rs1.EOF
         Set MFreelanceItem = New CFreelanceItem
         Call MFreelanceItem.PopulateFromRS(1, m_Rs1)
         
         Call Ua.CstFreelance.add(MFreelanceItem)
         Set MFreelanceItem = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim MDeliveryCusItem As CDeliveryCus
      Set MDeliveryCusItem = New CDeliveryCus
      MDeliveryCusItem.CUSTOMER_ID = Ua.CUSTOMER_ID
      Call MDeliveryCusItem.QueryData(m_Rs1, iCount)
      Set MDeliveryCusItem = Nothing
      
      
      Set Ua.CstdDeliveryCus = Nothing
      Set Ua.CstdDeliveryCus = New Collection
     
      While Not m_Rs1.EOF
         Set MDeliveryCusItem = New CDeliveryCus
         Call MDeliveryCusItem.PopulateFromRS(1, m_Rs1)
         
         Call Ua.CstdDeliveryCus.add(MDeliveryCusItem)
         Set MDeliveryCusItem = Nothing

         m_Rs1.MoveNext
      Wend
      
'      Dim MPromotional As CPromotional
'      Set MPromotional = New CPromotional
'      MPromotional.CUSTOMER_ID = Ua.CUSTOMER_ID
'      Call MPromotional.QueryData(m_Rs1, iCount)
'      Set MPromotional = Nothing
'
'
'      Set Ua.CstdPromotion = Nothing
'      Set Ua.CstdPromotion = New Collection
'
'      While Not m_Rs1.EOF
'         Set MPromotional = New CPromotional
'         Call MPromotional.PopulateFromRS(1, m_Rs1)
'
'         Call Ua.CstdPromotion.add(MPromotional)
'         Set MPromotional = Nothing
'
'         m_Rs1.MoveNext
'      Wend


   End If
      
   QueryCustomer = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomer = False
End Function

Public Function DeleteCustomer(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CCustomer

   RName = "DeleteCustomer"
   DeleteCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CCustomer
   
   cCust.CUSTOMER_ID = UID
   Call cCust.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomer = True
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

   DeleteCustomer = True
End Function
Public Function AddEditBillTransport(Ug As CBillTransport, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBillTransport"
   AddEditBillTransport = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
      
   Dim BTI As CBillTransportItem
   For Each BTI In Ug.C_BillTransportItem
      If BTI.Flag = "A" Or BTI.Flag = "N" Then
         BTI.BILL_TRANSPORT_ID = Ug.BILL_TRANSPORT_ID
         BTI.AddEditMode = SHOW_ADD
         BTI.AddEditData
      ElseIf BTI.Flag = "E" Then
         BTI.BILL_TRANSPORT_ID = Ug.BILL_TRANSPORT_ID
         BTI.AddEditMode = SHOW_EDIT
         BTI.AddEditData
      ElseIf BTI.Flag = "D" Then
         BTI.DeleteData
      End If
   Next BTI

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditBillTransport = True
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
   AddEditBillTransport = False
End Function
Public Function AddEditCustomer(Ug As CCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomer"
   AddEditCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
      
   Dim CstName As CCustomerName
   Dim NAME As CName
   Dim CstAddr As CCustomerAddress
   Dim Addr As CAddress
   
   For Each CstAddr In Ug.CstAddr
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set NAME = CstName.NAME
         NAME.AddEditMode = SHOW_ADD
         Call NAME.AddEditData
                  
         CstName.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstName.NAME_ID = NAME.NAME_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set NAME = CstName.NAME
         NAME.AddEditMode = SHOW_EDIT
         Call NAME.AddEditData
      End If
   Next CstName
   
   Dim cAcc As CAccount
   Dim cSubc As CSubscriber
   Dim cAgr As CAgreement
   For Each cAcc In Ug.CstAccounts
      If cAcc.Flag = "A" Then
         cAcc.AddEditMode = SHOW_ADD
         cAcc.CUSTOMER_ID = Ug.CUSTOMER_ID
         Call cAcc.AddEditData
         
         Set cSubc = cAcc.ActSubs(1)
         cSubc.ACCOUNT_ID = cAcc.ACCOUNT_ID
         cSubc.AddEditMode = SHOW_ADD
         Call cSubc.AddEditData
      
         Set cAgr = cAcc.ActAgrmnts(1)
         cAgr.SUBSCRIBER_ID = cSubc.SUBSCRIBER_ID
         cAgr.AddEditMode = SHOW_ADD
         Call cAgr.AddEditData
      ElseIf cAcc.Flag = "E" Then
         cAcc.AddEditMode = SHOW_EDIT
         Call cAcc.AddEditData
         
         Set cSubc = cAcc.ActSubs(1)
         cSubc.AddEditMode = SHOW_EDIT
         Call cSubc.AddEditData
      
         Set cAgr = cAcc.ActAgrmnts(1)
         cAgr.AddEditMode = SHOW_EDIT
         Call cAgr.AddEditData
      ElseIf cAcc.Flag = "D" Then
         cAcc.AddEditMode = SHOW_EDIT
         cAcc.ENABLE_FLAG = "N"
         Call cAcc.AddEditData
         
         Set cSubc = cAcc.ActSubs(1)
         cSubc.AddEditMode = SHOW_EDIT
         Call cSubc.AddEditData
      End If
   Next cAcc
         
   Dim Bh As CCustomerPicture
   For Each Bh In Ug.CstPicture
      If Bh.Flag = "A" Then
         Call Bh.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
         Bh.ShowMode = SHOW_ADD
         
         Call Bh.SetFieldValue("CUSTOMER_PICTURE_PATH", SavePictureToServer(Bh.GetFieldValue("CUSTOMER_PICTURE_PATH")))
         Bh.AddEditData
      ElseIf Bh.Flag = "E" Then
         Bh.ShowMode = SHOW_EDIT
         Bh.AddEditData
      ElseIf Bh.Flag = "D" Then
         Bh.DeleteData
         Call DeletePictureFromDisk(Bh.GetFieldValue("CUSTOMER_PICTURE_PATH"))
      End If
   Next Bh
   
   Dim Acl1 As CCustomerAccountList
   For Each Acl1 In Ug.CstAccountList1s
      If Acl1.Flag = "A" Then
         Call Acl1.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
         Acl1.ShowMode = SHOW_ADD
         Acl1.AddEditData
      ElseIf Acl1.Flag = "E" Then
         Acl1.ShowMode = SHOW_EDIT
         Acl1.AddEditData
      ElseIf Acl1.Flag = "D" Then
         Acl1.DeleteData
      End If
   Next Acl1
   
   Dim Acl2 As CCustomerAccountList
   For Each Acl2 In Ug.CstAccountList2s
      If Acl2.Flag = "A" Then
         Call Acl2.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
         Acl2.ShowMode = SHOW_ADD
         Acl2.AddEditData
      ElseIf Acl2.Flag = "E" Then
         Acl2.ShowMode = SHOW_EDIT
         Acl2.AddEditData
      ElseIf Acl2.Flag = "D" Then
         Acl2.DeleteData
      End If
   Next Acl2
   
   Dim Acl3 As CCustomerAccountList
   For Each Acl3 In Ug.CstAccountList3s
      If Acl3.Flag = "A" Then
         Call Acl3.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
         Acl3.ShowMode = SHOW_ADD
         Acl3.AddEditData
      ElseIf Acl3.Flag = "E" Then
         Acl3.ShowMode = SHOW_EDIT
         Acl3.AddEditData
      ElseIf Acl3.Flag = "D" Then
         Acl3.DeleteData
      End If
   Next Acl3
   
   
'   Dim AccFol As CAccFol
'   For Each AccFol In Ug.CstAccFol
'      If AccFol.Flag = "A" Then
'         Call AccFol.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
'        AccFol.ShowMode = SHOW_ADD
'        AccFol.AddEditData
'      ElseIf AccFol.Flag = "E" Then
'         AccFol.ShowMode = SHOW_EDIT
'         AccFol.AddEditData
'      ElseIf AccFol.Flag = "D" Then
'         AccFol.DeleteData
'      End If
'   Next AccFol


   Dim AccFol As CAccFol
   For Each AccFol In Ug.CstAccFol
      If AccFol.Flag = "A" Then
         AccFol.CUSTOMER_ID = Ug.CUSTOMER_ID
         AccFol.AddEditMode = SHOW_ADD
         AccFol.AddEditData
      ElseIf AccFol.Flag = "E" Then
         AccFol.AddEditMode = SHOW_EDIT
         AccFol.AddEditData
      ElseIf AccFol.Flag = "D" Then
         Call AccFol.DeleteData
      End If
   Next AccFol
   
   Dim MKTFol As CMKTFol
   For Each MKTFol In Ug.CstMKTFol
      If MKTFol.Flag = "A" Then
         MKTFol.CUSTOMER_ID = Ug.CUSTOMER_ID
         MKTFol.AddEditMode = SHOW_ADD
         MKTFol.AddEditData
      ElseIf MKTFol.Flag = "E" Then
         MKTFol.AddEditMode = SHOW_EDIT
         MKTFol.AddEditData
      ElseIf MKTFol.Flag = "D" Then
         Call MKTFol.DeleteData
      End If
   Next MKTFol
   
   Dim Freelance As CFreelance
   Dim FreelanceItem As CFreelanceItem
   For Each FreelanceItem In Ug.CstFreelance
     
      If FreelanceItem.Flag = "A" Then
         FreelanceItem.AddEditMode = SHOW_ADD
         FreelanceItem.CUSTOMER_ID = Ug.CUSTOMER_ID
         Call FreelanceItem.AddEditData
      ElseIf FreelanceItem.Flag = "E" Then
         FreelanceItem.AddEditMode = SHOW_EDIT
         Call FreelanceItem.AddEditData
      ElseIf FreelanceItem.Flag = "D" Then
         FreelanceItem.AddEditMode = SHOW_EDIT
         Call FreelanceItem.DeleteData
      End If
   Next FreelanceItem
   
   Dim DC As CDeliveryCus
   For Each DC In Ug.CstdDeliveryCus
     
      If DC.Flag = "A" Then
         DC.AddEditMode = SHOW_ADD
         DC.CUSTOMER_ID = Ug.CUSTOMER_ID
         Call DC.AddEditData
      ElseIf DC.Flag = "E" Then
         DC.AddEditMode = SHOW_EDIT
         Call DC.AddEditData
      ElseIf DC.Flag = "D" Then
         DC.AddEditMode = SHOW_EDIT
         Call DC.DeleteData
      End If
   Next DC
   
'   Dim Pt As CPromotional
'   For Each Pt In Ug.CstdPromotion
'
'   If Pt.Flag = "A" Then
'         Pt.AddEditMode = SHOW_ADD
'         Pt.CUSTOMER_ID = Ug.CUSTOMER_ID
'         Call Pt.AddEditData
'      ElseIf Pt.Flag = "E" Then
'         Pt.AddEditMode = SHOW_EDIT
'         Call Pt.AddEditData
'      ElseIf Pt.Flag = "D" Then
'         Pt.AddEditMode = SHOW_EDIT
'         Call Pt.DeleteData
'      End If
'   Next Pt
    
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCustomer = True
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
   AddEditCustomer = False
End Function

'==
Public Function QuerySupplierInfo(Ua As CSupplier, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QuerySupplierInfo"
   QuerySupplierInfo = False

   IsOK = True
   Call Ua.QueryData1(Rs, ItemCount)
         
   QuerySupplierInfo = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QuerySupplierInfo = False
End Function

Public Function QuerySupplier(Ua As CSupplier, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QuerySupplier"
   QuerySupplier = False
   
   IsOK = True
   Call Ua.QueryData2(Rs, ItemCount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      Dim cCstContact As CSupplierContact
      Set cCstContact = New CSupplierContact
      cCstContact.SUPPLIER_CONTACT_ID = -1
      cCstContact.SUPPLIER_ID = Ua.SUPPLIER_ID
      cCstContact.OrderBy = 1
      Call cCstContact.QueryData(m_Rs1, iCount)
      Set cCstContact = Nothing
      
      Set Ua.CstContacts = Nothing
      Set Ua.CstContacts = New Collection
      While Not m_Rs1.EOF
         Set cCstContact = New CSupplierContact
         Call cCstContact.PopulateFromRS(m_Rs1)

         Dim N As CName
         Set N = New CName
         N.NAME_ID = cCstContact.NAME_ID
         Call N.QueryData(m_Rs2, iCount)
         Set N = Nothing
         While Not m_Rs2.EOF
            Set N = New CName
            Call N.PopulateFromRS(m_Rs2)
            N.Flag = "I"
            Set cCstContact.NAME = N
            Set N = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstContact.Flag = "I"
         Call Ua.CstContacts.add(cCstContact)
         Set cCstContact = Nothing
         m_Rs1.MoveNext
      Wend
      'Coontact ++++++++++++++++++++++++++
   
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CSupplierName
      Set cCstName = New CSupplierName
      cCstName.SUPPLIER_NAME_ID = -1
      cCstName.SUPPLIER_ID = Ua.SUPPLIER_ID
      cCstName.OrderBy = 1
      Call cCstName.QueryData(m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CSupplierName
         Call cCstName.PopulateFromRS(m_Rs1)

         Dim CName As CName
         Dim iCount2 As Long
         Set CName = New CName
         CName.NAME_ID = cCstName.NAME_ID
         Call CName.QueryData(m_Rs2, iCount2)
         Set CName = Nothing
         While Not m_Rs2.EOF
            Set CName = New CName
            Call CName.PopulateFromRS(m_Rs2)
            CName.Flag = "I"
            Set cCstName.NAME = CName
            Set CName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CSupplierAddress
      Set cEnpAddr = New CSupplierAddress
      cEnpAddr.SUPPLIER_ID = Ua.SUPPLIER_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddr = Nothing
      Set Ua.CstAddr = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CSupplierAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            cAddr.ADDRESS_TYPE = cEnpAddr.ADDRESS_TYPE
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddr.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
      
'      Dim Li As CLotItem
'      Set Li = New CLotItem
'      Li.SUPPLIER_ID = Ua.SUPPLIER_ID
'      Li.OrderBy = -1
'      Li.OrderType = 1
'      Call Li.QueryData(10, m_Rs1, iCount)
'      Set Li = Nothing
'
'      Set Ua.PartItems = Nothing
'      Set Ua.PartItems = New Collection
'      While Not m_Rs1.EOF
'         Set Li = New CLotItem
'         Call Li.PopulateFromRS(10, m_Rs1)
'
'         Li.Flag = "I"
'         Call Ua.PartItems.add(Li)
'
'         Set Li = Nothing
'         m_Rs1.MoveNext
'      Wend
   
      Dim Sp As CSupplierSpec
      Set Sp = New CSupplierSpec
      Sp.SUPPLIER_ID = Ua.SUPPLIER_ID
      Sp.OrderBy = -1
      Sp.OrderType = 1
      Call Sp.QueryData(1, m_Rs1, iCount)
      Set Sp = Nothing
      
      Set Ua.SupplierSpecs = Nothing
      Set Ua.SupplierSpecs = New Collection
      While Not m_Rs1.EOF
         Set Sp = New CSupplierSpec
         Call Sp.PopulateFromRS(1, m_Rs1)
                  
         Sp.Flag = "I"
         Call Ua.SupplierSpecs.add(Sp)
         
         Set Sp = Nothing
         m_Rs1.MoveNext
      Wend
   
      Dim Su As CSupplierUsed
      Set Su = New CSupplierUsed
      Su.SUPPLIER_ID = Ua.SUPPLIER_ID
      Su.OrderBy = -1
      Su.OrderType = 1
      Call Su.QueryData(1, m_Rs1, iCount)
      Set Su = Nothing
      
      Set Ua.SupplierUseds = Nothing
      Set Ua.SupplierUseds = New Collection
      While Not m_Rs1.EOF
         Set Su = New CSupplierUsed
         Call Su.PopulateFromRS(1, m_Rs1)
                  
         Su.Flag = "I"
         Call Ua.SupplierUseds.add(Su)
         
         Set Su = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim St As CSupplierTranSport
      Set St = New CSupplierTranSport
      St.SUPPLIER_ID = Ua.SUPPLIER_ID
      St.OrderBy = -1
      St.OrderType = 1
      Call St.QueryData(1, m_Rs1, iCount)
      Set St = Nothing
      
      Set Ua.SupplierTranSport = Nothing
      Set Ua.SupplierTranSport = New Collection
      While Not m_Rs1.EOF
         Set St = New CSupplierTranSport
         Call St.PopulateFromRS(1, m_Rs1)
                  
         St.Flag = "I"
         Call Ua.SupplierTranSport.add(St)
         
         Set St = Nothing
         m_Rs1.MoveNext
      Wend
      
     Dim SA As CSupplierAccount
      Set SA = New CSupplierAccount
      SA.SUPPLIER_ID = Ua.SUPPLIER_ID
      SA.OrderBy = -1
      SA.OrderType = 1
      Call SA.QueryData(1, m_Rs1, iCount)
      Set SA = Nothing
      
      Set Ua.SupplierAccount = Nothing
      Set Ua.SupplierAccount = New Collection
      While Not m_Rs1.EOF
         Set SA = New CSupplierAccount
         Call SA.PopulateFromRS(1, m_Rs1)
                  
         SA.Flag = "I"
         Call Ua.SupplierAccount.add(SA)
         
         Set SA = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QuerySupplier = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QuerySupplier = False
End Function

Public Function AddEditSupplier(Ug As CSupplier, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplier"
   AddEditSupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
   
   Dim CstName As CSupplierName
   Dim NAME As CName
   Dim CstAddr As CSupplierAddress
   Dim Addr As CAddress
      
   For Each CstAddr In Ug.CstAddr
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set NAME = CstName.NAME
         NAME.AddEditMode = SHOW_ADD
         Call NAME.AddEditData
                  
         CstName.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstName.NAME_ID = NAME.NAME_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set NAME = CstName.NAME
         NAME.AddEditMode = SHOW_EDIT
         Call NAME.AddEditData
      End If
   Next CstName
   
   Dim CstContact As CSupplierContact
   For Each CstContact In Ug.CstContacts
      If CstContact.Flag = "A" Then
         Set NAME = CstContact.NAME
         NAME.AddEditMode = SHOW_ADD
         Call NAME.AddEditData
                  
         CstContact.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstContact.NAME_ID = NAME.NAME_ID
         CstContact.AddEditMode = SHOW_ADD
         CstContact.AddEditData
      ElseIf CstContact.Flag = "E" Then
         Set NAME = CstContact.NAME
         NAME.AddEditMode = SHOW_EDIT
         Call NAME.AddEditData
      
         CstContact.AddEditMode = SHOW_EDIT
         CstContact.AddEditData
      ElseIf CstContact.Flag = "D" Then
         Call CstContact.DeleteData

         Set NAME = CstContact.NAME
         Call NAME.DeleteData
      End If
   Next CstContact
            
   Dim Sp As CSupplierSpec
   For Each Sp In Ug.SupplierSpecs
      If Sp.Flag = "A" Then
         Sp.SUPPLIER_ID = Ug.SUPPLIER_ID
         Sp.AddEditMode = SHOW_ADD
         Sp.AddEditData
      ElseIf Sp.Flag = "E" Then
         Sp.AddEditMode = SHOW_EDIT
         Sp.AddEditData
      ElseIf Sp.Flag = "D" Then
         Sp.DeleteData
      End If
   Next Sp

   Dim Su As CSupplierUsed
   For Each Su In Ug.SupplierUseds
      If Su.Flag = "A" Then
         Su.SUPPLIER_ID = Ug.SUPPLIER_ID
         Su.AddEditMode = SHOW_ADD
         Su.AddEditData
      ElseIf Su.Flag = "E" Then
         Su.AddEditMode = SHOW_EDIT
         Su.AddEditData
      ElseIf Su.Flag = "D" Then
         Su.DeleteData
      End If
   Next Su
   
   Dim St As CSupplierTranSport
   For Each St In Ug.SupplierTranSport
      If St.Flag = "A" Then
         St.SUPPLIER_ID = Ug.SUPPLIER_ID
         St.AddEditMode = SHOW_ADD
         St.AddEditData
      ElseIf St.Flag = "E" Then
         St.AddEditMode = SHOW_EDIT
         St.AddEditData
      ElseIf St.Flag = "D" Then
         St.DeleteData
      End If
   Next St
   
   Dim SA As CSupplierAccount
   For Each SA In Ug.SupplierAccount
      If SA.Flag = "A" Then
         SA.SUPPLIER_ID = Ug.SUPPLIER_ID
         SA.AddEditMode = SHOW_ADD
         SA.AddEditData
      ElseIf SA.Flag = "E" Then
         SA.AddEditMode = SHOW_EDIT
         SA.AddEditData
      ElseIf SA.Flag = "D" Then
         SA.DeleteData
      End If
   Next SA


   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditSupplier = True
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
   AddEditSupplier = False
End Function

Public Function DeleteSupplier(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CSupplier

   RName = "DeleteSupplier"
   DeleteSupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CSupplier
   
   cCust.SUPPLIER_ID = UID
   Call cCust.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSupplier = True
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

   DeleteSupplier = True
End Function

'==
Public Function QueryEmployee(Ua As CEmployee, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

Dim Qp As CEmployee

   RName = "QueryEmployee"
   QueryEmployee = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Set Ua.EmpName = Nothing
      Set Ua.EmpName = New CEmployeeName
      Ua.EmpName.EMPLOYEE_NAME_ID = -1
      Ua.EmpName.EMP_ID = Ua.EMP_ID
      Call Ua.EmpName.QueryData(m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EmpName.PopulateFromRS(m_Rs1)
      End If
      
      Set Ua.EName = Nothing
      Set Ua.EName = New CName
      Ua.EName.NAME_ID = Ua.EmpName.NAME_ID
      Call Ua.EName.QueryData(m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EName.PopulateFromRS(m_Rs1)
      End If
    'Address ++++++++++++++++++++++++++
      Dim cEmpAddr As CEmpAddress
      Set cEmpAddr = New CEmpAddress
      cEmpAddr.EMP_ID = Ua.EMP_ID
      cEmpAddr.OrderBy = 1
      Call cEmpAddr.QueryData(m_Rs2, iCount)
      Set cEmpAddr = Nothing
      
      Set Ua.Contacts = Nothing
      Set Ua.Contacts = New Collection
      While Not m_Rs2.EOF
         Set cEmpAddr = New CEmpAddress
         Call cEmpAddr.PopulateFromRS(m_Rs2)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEmpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs3, iCount3)
         Set cAddr = Nothing
         While Not m_Rs3.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs3)
            cAddr.Flag = "I"
            Call cEmpAddr.Addresses.add(cAddr)
            Set cAddr = Nothing
            
            m_Rs3.MoveNext
         Wend
      
         cEmpAddr.Flag = "I"
         Call Ua.Contacts.add(cEmpAddr)
         Set cEmpAddr = Nothing
         m_Rs2.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   
      'Document ++++++++++++++++++++++++++
      Dim cCstProof As CEmployeeProof
      Set cCstProof = New CEmployeeProof
      cCstProof.AddEditMode = -1
      cCstProof.EMP_ID = Ua.EMP_ID
      cCstProof.OrderBy = 1
      Call cCstProof.QueryData(m_Rs1, iCount)
      Set cCstProof = Nothing
      
      Set Ua.Cards = Nothing
      Set Ua.Cards = New Collection
      While Not m_Rs1.EOF
         Set cCstProof = New CEmployeeProof
         Call cCstProof.PopulateFromRS(m_Rs1)

         Dim Doc As CDocument
         Set Doc = New CDocument
         Doc.DOCUMENT_ID = cCstProof.DOCUMENT_ID
         Call Doc.QueryData(m_Rs2, iCount3)
         Call Doc.PopulateFromRS(m_Rs2)
         
         Dim Addr As CAddress
         Set Addr = New CAddress
         
         Addr.ADDRESS_ID = Doc.ISSUE_PLACE_ID
         Call Addr.QueryData(m_Rs3, iCount3)
         Call Addr.PopulateFromRS(m_Rs3)
         Addr.Flag = "I"
         
         Set Doc.Address = Addr
         Set Addr = Nothing
      
         Doc.Flag = "I"
         cCstProof.Flag = "I"
         Set cCstProof.Doc = Doc
         Call Ua.Cards.add(cCstProof)
         Set cCstProof = Nothing
         m_Rs1.MoveNext
      Wend
      'Document ++++++++++++++++++++++++++
   'EmpWorked ++++++++++++++++++++++++++
      Dim EW As CEmpWorked
      Set EW = New CEmpWorked
      EW.EMP_ID = Ua.EMP_ID
      EW.OrderBy = 1
      Call EW.QueryData(m_Rs1, iCount)
      Set EW = Nothing
      
      Set Ua.EmpWorked = Nothing
      Set Ua.EmpWorked = New Collection
      While Not m_Rs1.EOF
         Set EW = New CEmpWorked
         Call EW.PopulateFromRS(m_Rs1)
      
         EW.Flag = "I"
         Call Ua.EmpWorked.add(EW)
         Set EW = Nothing
         m_Rs1.MoveNext
      Wend
      'EmpWorked ++++++++++++++++++++++++++

      'EmpEducation ++++++++++++++++++++++++++
      Dim EE As CEmpEducation
      Set EE = New CEmpEducation
      EE.EMP_ID = Ua.EMP_ID
      EE.OrderBy = 1
      Call EE.QueryData(m_Rs1, iCount)
      Set EE = Nothing
      
      Set Ua.EmpEducation = Nothing
      Set Ua.EmpEducation = New Collection
      While Not m_Rs1.EOF
         Set EE = New CEmpEducation
         Call EE.PopulateFromRS(m_Rs1)
      
         EE.Flag = "I"
         Call Ua.EmpEducation.add(EE)
         Set EE = Nothing
         m_Rs1.MoveNext
      Wend
      'EmpEducation ++++++++++++++++++++++++++
     
      'EmpDependency ++++++++++++++++++++++++++
      Dim Ed As CEmpDependency
      Set Ed = New CEmpDependency
      Ed.EMP_ID = Ua.EMP_ID
      Ed.OrderBy = 1
      Call Ed.QueryData(m_Rs1, iCount)
      Set Ed = Nothing
      
      Set Ua.EmpDependency = Nothing
      Set Ua.EmpDependency = New Collection
      While Not m_Rs1.EOF
         Set Ed = New CEmpDependency
         Call Ed.PopulateFromRS(m_Rs1)

         Dim ND As CName
         Set ND = New CName
         ND.NAME_ID = Ed.NAME_ID
         Call ND.QueryData(m_Rs2, iCount)
         Set ND = Nothing
         If Not m_Rs2.EOF Then
            Set ND = New CName
            Call ND.PopulateFromRS(m_Rs2)
            ND.Flag = "I"
            Set Ed.NAME = ND
            Set ND = Nothing
         End If
         Ed.Flag = "I"
         Call Ua.EmpDependency.add(Ed)
         Set Ed = Nothing
         m_Rs1.MoveNext
      Wend
      'EmpDependency ++++++++++++++++++++++++++
      
      'EmpChild ++++++++++++++++++++++++++
      Dim EC As CEmpChild
      Set EC = New CEmpChild
      EC.EMP_ID = Ua.EMP_ID
      EC.OrderBy = 1
      Call EC.QueryData(m_Rs1, iCount)
      Set EC = Nothing
      
      Set Ua.EmpChild = Nothing
      Set Ua.EmpChild = New Collection
      While Not m_Rs1.EOF
         Set EC = New CEmpChild
         Call EC.PopulateFromRS(m_Rs1)

         Dim NC As CName
         Set NC = New CName
         NC.NAME_ID = EC.NAME_ID
         Call NC.QueryData(m_Rs2, iCount)
         Set NC = Nothing
         If Not m_Rs2.EOF Then
            Set NC = New CName
            Call NC.PopulateFromRS(m_Rs2)
            NC.Flag = "I"
            Set EC.NAME = NC
            Set NC = Nothing
         End If
         EC.Flag = "I"
         Call Ua.EmpChild.add(EC)
         Set EC = Nothing
         m_Rs1.MoveNext
      Wend
      'EmpChild ++++++++++++++++++++++++++
'EmpHealty ++++++++++++++++++++++++++
      Dim EH As CEmpHealty
      Set EH = New CEmpHealty
      EH.EMP_ID = Ua.EMP_ID
      EH.OrderBy = 1
      Call EH.QueryData(m_Rs1, iCount)
      Set EH = Nothing
      
      Set Ua.EmpHealty = Nothing
      Set Ua.EmpHealty = New Collection
      While Not m_Rs1.EOF
         Set EH = New CEmpHealty
         Call EH.PopulateFromRS(m_Rs1)
      
         EH.Flag = "I"
         Call Ua.EmpHealty.add(EH)
         Set EH = Nothing
         m_Rs1.MoveNext
      Wend
      'EmpHealty ++++++++++++++++++++++++++
      
   End If
   
   QueryEmployee = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmployee = False
End Function

Public Function QueryFreeLance(Ua As CFreelance, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryFreelance"
   QueryFreeLance = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryFreeLance = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryFreeLance = False
End Function
Public Function QueryDeliveryCus(Ua As CDeliveryCus, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryDeliveryCus"
   QueryDeliveryCus = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryDeliveryCus = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryDeliveryCus = False
End Function
Public Function QueryFreeLanceByCustomer(Ua As CFreelance, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryFreeLanceByCustomer"
   QueryFreeLanceByCustomer = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount, 2)
   
   QueryFreeLanceByCustomer = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryFreeLanceByCustomer = False
End Function

Public Function DeleteEmployee(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CEmployee

   RName = "DeleteEmployee"
   DeleteEmployee = False
   HasBegin = False

   Set B = New CEmployee

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.EMP_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEmployee = True
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

   DeleteEmployee = True
End Function

Public Function DeleteFreelance(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CFreelance

   RName = "DeleteFreelance"
   DeleteFreelance = False
   HasBegin = False

   Set B = New CFreelance

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.FREELANCE_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteFreelance = True
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
   DeleteFreelance = True
End Function

Public Function AddEditFreelance(At As CFreelance, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFreelance"
   AddEditFreelance = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditFreelance = True
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
   AddEditFreelance = False
End Function
Public Function AddEditDeliveryCus(At As CDeliveryCus, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditDeliveryCus"
   AddEditDeliveryCus = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call At.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditDeliveryCus = True
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
   AddEditDeliveryCus = False
End Function

Public Function AddEditEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEmployee"
   AddEditEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Call Ug.EName.AddEditData
   
   Ug.EmpName.EMP_ID = Ug.EMP_ID
   Ug.EmpName.NAME_ID = Ug.EName.NAME_ID
   Call Ug.EmpName.AddEditData
      
   '   Dim CstName As CCustomerName
   'Dim Name As cName
   Dim CstAddr As CEmpAddress
   Dim Addr As CAddress
   
   For Each CstAddr In Ug.Contacts
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses(1)
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.EMP_ID = Ug.EMP_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses(1)
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.EMP_ID = Ug.EMP_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses(1)
         Call Addr.DeleteData
      End If
   Next CstAddr

   
Dim CstProof As CEmployeeProof
Dim Doc  As CDocument
 For Each CstProof In Ug.Cards
      If CstProof.Flag = "A" Then
         Set Doc = CstProof.Doc
         Set Addr = Doc.Address
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData
         
         Doc.ISSUE_PLACE_ID = Addr.ADDRESS_ID
         Doc.AddEditMode = SHOW_ADD
         Call Doc.AddEditData
         
         CstProof.EMP_ID = Ug.EMP_ID
         CstProof.DOCUMENT_ID = Doc.DOCUMENT_ID
         CstProof.AddEditMode = SHOW_ADD
         Call CstProof.AddEditData
      ElseIf CstProof.Flag = "E" Then
         Set Doc = CstProof.Doc
         Set Addr = Doc.Address
         
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         Doc.AddEditMode = SHOW_EDIT
         Call Doc.AddEditData
         
         CstProof.EMP_ID = Ug.EMP_ID
         CstProof.DOCUMENT_ID = Doc.DOCUMENT_ID
         CstProof.AddEditMode = SHOW_EDIT
         CstProof.AddEditData
      ElseIf CstProof.Flag = "D" Then
         Call CstProof.DeleteData

         Set Doc = CstProof.Doc
         Call Doc.DeleteData
         
         Set Addr = Doc.Address
         Call Addr.DeleteData
      End If
   Next CstProof
  

   Dim EW As CEmpWorked
   For Each EW In Ug.EmpWorked
      If EW.Flag = "A" Then
         EW.EMP_ID = Ug.EMP_ID
         EW.AddEditMode = SHOW_ADD
         EW.AddEditData
      ElseIf EW.Flag = "E" Then
         EW.AddEditMode = SHOW_EDIT
         EW.AddEditData
      ElseIf EW.Flag = "D" Then
         Call EW.DeleteData
      End If
   Next EW
   
      Dim EE As CEmpEducation
   For Each EE In Ug.EmpEducation
      If EE.Flag = "A" Then
         EE.EMP_ID = Ug.EMP_ID
         EE.AddEditMode = SHOW_ADD
         EE.AddEditData
      ElseIf EE.Flag = "E" Then
         EE.AddEditMode = SHOW_EDIT
         EE.AddEditData
      ElseIf EE.Flag = "D" Then
         Call EE.DeleteData
      End If
   Next EE


   Dim Ed As CEmpDependency
   Dim ND As CName
   For Each Ed In Ug.EmpDependency
      If Ed.Flag = "A" Then
         Set ND = Ed.NAME
         ND.AddEditMode = SHOW_ADD
         Call ND.AddEditData

         Ed.EMP_ID = Ug.EMP_ID
         Ed.NAME_ID = ND.NAME_ID
         Ed.AddEditMode = SHOW_ADD
         Call Ed.AddEditData
      ElseIf Ed.Flag = "E" Then
         Set ND = Ed.NAME
         ND.AddEditMode = SHOW_EDIT
         Call ND.AddEditData

         Ed.EMP_ID = Ug.EMP_ID
         Ed.NAME_ID = ND.NAME_ID
         Ed.AddEditMode = SHOW_EDIT
         Ed.AddEditData
      ElseIf Ed.Flag = "D" Then
         Call Ed.DeleteData

         Set ND = Ed.NAME
         Call ND.DeleteData
      End If
   Next Ed


   Dim EC As CEmpChild
   Dim NC As CName
   For Each EC In Ug.EmpChild
      If EC.Flag = "A" Then
         Set NC = EC.NAME
         NC.AddEditMode = SHOW_ADD
         Call NC.AddEditData

         EC.EMP_ID = Ug.EMP_ID
         EC.NAME_ID = NC.NAME_ID
         EC.AddEditMode = SHOW_ADD
         Call EC.AddEditData
      ElseIf EC.Flag = "E" Then
         Set NC = EC.NAME
         NC.AddEditMode = SHOW_EDIT
         Call NC.AddEditData

         EC.EMP_ID = Ug.EMP_ID
         EC.NAME_ID = NC.NAME_ID
         EC.AddEditMode = SHOW_EDIT
         EC.AddEditData
      ElseIf EC.Flag = "D" Then
         Call EC.DeleteData

         Set NC = EC.NAME
         Call NC.DeleteData
      End If
   Next EC
   
   
   Dim EH As CEmpHealty
   For Each EH In Ug.EmpHealty
      If EH.Flag = "A" Then
         EH.EMP_ID = Ug.EMP_ID
         EH.AddEditMode = SHOW_ADD
         EH.AddEditData
      ElseIf EH.Flag = "E" Then
         EH.AddEditMode = SHOW_EDIT
         EH.AddEditData
      ElseIf EH.Flag = "D" Then
         Call EH.DeleteData
      End If
   Next EH
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEmployee = True
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
   AddEditEmployee = False
End Function

Private Sub QuerySubLotItem(Li As CLotItem)
Dim C As CSubLotItem
Dim iCount As Long

   Set C = New CSubLotItem
   C.LOT_ITEM_ID = Li.LOT_ITEM_ID
   Call C.QueryData(1, m_Rs2, iCount)
   Set C = Nothing
   
   Set Li.SubLotItems = Nothing
   Set Li.SubLotItems = New Collection
   While Not m_Rs2.EOF
      
      Set C = New CSubLotItem
      C.Flag = "I"
      Call C.PopulateFromRS(1, m_Rs2)
      Call Li.SubLotItems.add(C)
      Set C = Nothing
      
      m_Rs2.MoveNext
   Wend
   
   Set C = Nothing
End Sub
'==
Public Function QueryInventoryDoc(Ua As CInventoryDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryInventoryDoc"
   QueryInventoryDoc = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CLotItem
      Set Yw = New CLotItem
      Yw.LOT_ITEM_ID = -1
      Yw.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.ImportExports = Nothing
      Set Ua.ImportExports = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CLotItem
         Call Yw.PopulateFromRS(1, m_Rs1)

         Call QuerySubLotItem(Yw)

         Yw.Flag = "I"
         Call Ua.ImportExports.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
      
'      Dim IWDExport As CInventoryWHDoc
'      Set IWDExport = New CInventoryWHDoc
'      IWDExport.INVENTORY_WH_DOC_ID = -1
'      IWDExport.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
'      IWDExport.OrderBy = 1
'      Call IWDExport.QueryData(1, m_Rs5, iCount)
'      Set IWDExport = Nothing
'
'      Set Ua.WhExportItems = Nothing
'      Set Ua.WhExportItems = New Collection
'      While Not m_Rs5.EOF
'         Set IWDExport = New CInventoryWHDoc
'         Call IWDExport.PopulateFromRS(1, m_Rs5)
'           IWDExport.QueryFlag = 1
'           Call glbDaily.QueryInventoryWhDoc(IWDExport, m_Rs2, iCount, IsOK, glbErrorLog)
'
'         IWDExport.Flag = "I"
'         Call Ua.WhExportItems.add(IWDExport)
'
'         Set IWDExport = Nothing
'         m_Rs5.MoveNext
'      Wend
      
   End If
   
   QueryInventoryDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryDoc = False
End Function
'==
Public Function QueryInventoryDocLinkSaleOrder(Ua As CInventoryDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryInventoryDocLinkSaleOrder"
   QueryInventoryDocLinkSaleOrder = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount, 2)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CLotItem
      Set Yw = New CLotItem
      Yw.LOT_ITEM_ID = -1
      Yw.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(35, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.ImportExports = Nothing
      Set Ua.ImportExports = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CLotItem
         Call Yw.PopulateFromRS(35, m_Rs1)

         Call QuerySubLotItem(Yw)

         Yw.Flag = "I"
         Call Ua.ImportExports.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryInventoryDocLinkSaleOrder = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryDocLinkSaleOrder = False
End Function
'==
Public Function QueryInventoryWhDoc(Ua As CInventoryWHDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim LWH As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim I As Long
   RName = "QueryInventoryWhDoc"
   QueryInventoryWhDoc = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   Ua.Flag = "I"
   If Ua.QueryFlag = 1 Then
      Set Ua.C_LotItemsWH = Nothing
      Set Ua.C_LotItemsWH = New Collection
      Set LWH = New CLotItemWH
      LWH.LOT_ITEM_WH_ID = -1
      LWH.INVENTORY_WH_DOC_ID = Ua.INVENTORY_WH_DOC_ID
      LWH.OrderBy = 1
      Call LWH.QueryData(1, m_Rs1, iCount)
      Set LWH = Nothing
      While Not m_Rs1.EOF
         Set LWH = New CLotItemWH
         Call LWH.PopulateFromRS(1, m_Rs1)
               LWH.Flag = "I"
               Set LTD = New CLotDoc
               LTD.LOT_ITEM_WH_ID = LWH.LOT_ITEM_WH_ID
               LTD.OrderBy = 1
               If LWH.TX_TYPE = "E" Then
               Call LTD.QueryData(4, m_Rs2, iCount)
               ElseIf LWH.TX_TYPE = "I" Then
               Call LTD.QueryData(2, m_Rs2, iCount)
               End If
               Set LTD = Nothing
                  While Not m_Rs2.EOF
                   Set LTD = New CLotDoc
                  Call LTD.PopulateFromRS(4, m_Rs2)
                  LTD.Flag = "I"
                     Set PD = New CPalletDoc
                     PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                     PD.OrderBy = 2
                     PD.OrderType = 1
                     Call PD.QueryData(2, m_Rs3, iCount)
                     Set PD = Nothing
                        While Not m_Rs3.EOF
                           Set PD = New CPalletDoc
                           Call PD.PopulateFromRS(2, m_Rs3)
                           PD.Flag = "I"
                           Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(LTD.LOT_ID)))
                           Set PD = Nothing
                           m_Rs3.MoveNext
                        Wend
                  Call LWH.C_LotDoc.add(LTD)
                  Set LTD = Nothing
                  m_Rs2.MoveNext
                  Wend
           Call Ua.C_LotItemsWH.add(LWH)
           Set LWH = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryInventoryWhDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryWhDoc = False
End Function

Public Function QueryInventoryWhDocForImBulk(Ua As CInventoryWHDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim LWH As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim I As Long
   RName = "QueryInventoryWhDocForImBulk"
   QueryInventoryWhDocForImBulk = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   Ua.Flag = "I"
   If Ua.QueryFlag = 1 Then
      Set Ua.C_LotItemsWH = Nothing
      Set Ua.C_LotItemsWH = New Collection
      Set LWH = New CLotItemWH
      LWH.LOT_ITEM_WH_ID = -1
      LWH.INVENTORY_WH_DOC_ID = Ua.INVENTORY_WH_DOC_ID
      LWH.OrderBy = 1
      Call LWH.QueryData(1, m_Rs1, iCount)
      Set LWH = Nothing
      While Not m_Rs1.EOF
         Set LWH = New CLotItemWH
         Call LWH.PopulateFromRS(1, m_Rs1)
               LWH.Flag = "I"
               Set LTD = New CLotDoc
               LTD.LOT_ITEM_WH_ID = LWH.LOT_ITEM_WH_ID
               LTD.OrderBy = 1
               If LWH.TX_TYPE = "E" Then
               Call LTD.QueryData(4, m_Rs2, iCount)
               ElseIf LWH.TX_TYPE = "I" Then
               Call LTD.QueryData(2, m_Rs2, iCount)
               End If
               Set LTD = Nothing
                  While Not m_Rs2.EOF
                   Set LTD = New CLotDoc
                  Call LTD.PopulateFromRS(4, m_Rs2)
                     LTD.Flag = "I"
                     Set PD = New CPalletDoc
                     PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                     PD.OrderBy = 2
                     PD.OrderType = 1
                     Call PD.QueryData(2, m_Rs3, iCount)
                     Set PD = Nothing
                        While Not m_Rs3.EOF
                           Set PD = New CPalletDoc
                           Call PD.PopulateFromRS(2, m_Rs3)
                           PD.Flag = "I"
                           Call LTD.C_PalletDoc.add(PD, Trim(str(LWH.PART_ITEM_ID)))
                           Set PD = Nothing
                           m_Rs3.MoveNext
                        Wend
                  Call LWH.C_LotDoc.add(LTD, Trim(str(LWH.PART_ITEM_ID)))
                  Set LTD = Nothing
                  m_Rs2.MoveNext
                  Wend
           Call Ua.C_LotItemsWH.add(LWH, Trim(str(LWH.PART_ITEM_ID)))
           Set LWH = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryInventoryWhDocForImBulk = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryWhDocForImBulk = False
End Function
'==
Public Function QueryInventoryWhDocInput(Ua As CInventoryWHDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim LWH As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim I As Long
   RName = "QueryInventoryWhDoc"
   QueryInventoryWhDocInput = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   Ua.Flag = "I"
   If Ua.QueryFlag = 1 Then
      Set Ua.C_LotItemsWH = Nothing
      Set Ua.C_LotItemsWH = New Collection
      Set LWH = New CLotItemWH
      LWH.LOT_ITEM_WH_ID = -1
      LWH.INVENTORY_WH_DOC_ID = Ua.INVENTORY_WH_DOC_ID
      LWH.OrderBy = 1
      Call LWH.QueryData(16, m_Rs1, iCount)
      Set LWH = Nothing
      While Not m_Rs1.EOF
         Set LWH = New CLotItemWH
         Call LWH.PopulateFromRS(16, m_Rs1)
               LWH.Flag = "I"

               Set LTD = New CLotDoc
               LTD.LOT_ITEM_WH_ID = LWH.LOT_ITEM_WH_ID
               LTD.OrderBy = 1
               If LWH.TX_TYPE = "E" Then
               Call LTD.QueryData(4, m_Rs2, iCount)
               ElseIf LWH.TX_TYPE = "I" Then
               Call LTD.QueryData(2, m_Rs2, iCount)
               End If
               Set LTD = Nothing
                  While Not m_Rs2.EOF
                   Set LTD = New CLotDoc
                  Call LTD.PopulateFromRS(4, m_Rs2)
                  LTD.Flag = "I"
                     Set PD = New CPalletDoc
                     PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                     PD.OrderBy = 2
                     PD.OrderType = 1
                     Call PD.QueryData(2, m_Rs3, iCount)
                     Set PD = Nothing
                        While Not m_Rs3.EOF
                           Set PD = New CPalletDoc
                           Call PD.PopulateFromRS(2, m_Rs3)
                           PD.Flag = "I"
                           Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(LTD.LOT_ID)))
                           Set PD = Nothing
                           m_Rs3.MoveNext
                        Wend
                  Call LWH.C_LotDoc.add(LTD)
                  Set LTD = Nothing
                  m_Rs2.MoveNext
                  Wend
           Call Ua.C_LotItemsWH.add(LWH)
           Set LWH = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryInventoryWhDocInput = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryWhDocInput = False
End Function
Public Function QueryInventoryWhDocForLG(Ua As CInventoryWHDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim LWH As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim I As Long
   RName = "QueryInventoryWhDocForLG"
   QueryInventoryWhDocForLG = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   Ua.Flag = "I"
   If Ua.QueryFlag = 1 Then
      Set Ua.C_LotItemsWH = Nothing
      Set Ua.C_LotItemsWH = New Collection
      Set LWH = New CLotItemWH
      LWH.LOT_ITEM_WH_ID = -1
      LWH.INVENTORY_WH_DOC_ID = Ua.INVENTORY_WH_DOC_ID
      LWH.OrderBy = 1
      
      Call LWH.QueryData(10, m_Rs1, iCount)
      Set LWH = Nothing
      While Not m_Rs1.EOF
         Set LWH = New CLotItemWH
         Call LWH.PopulateFromRS(10, m_Rs1)
               LWH.Flag = "I"
               Set LTD = New CLotDoc
               LTD.LOT_ITEM_WH_ID = LWH.LOT_ITEM_WH_ID
               LTD.OrderBy = 1
               If LWH.TX_TYPE = "E" Then
               Call LTD.QueryData(4, m_Rs2, iCount)
               ElseIf LWH.TX_TYPE = "I" Then
               Call LTD.QueryData(2, m_Rs2, iCount)
               End If
               Set LTD = Nothing
                  While Not m_Rs2.EOF
                   Set LTD = New CLotDoc
                  Call LTD.PopulateFromRS(4, m_Rs2)
                  LTD.Flag = "I"
                     Set PD = New CPalletDoc
                     PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                     PD.OrderBy = 2
                     PD.OrderType = 1
                     Call PD.QueryData(2, m_Rs3, iCount)
                     Set PD = Nothing
                        While Not m_Rs3.EOF
                           Set PD = New CPalletDoc
                           Call PD.PopulateFromRS(2, m_Rs3)
                           PD.Flag = "I"
                           Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(LTD.LOT_ID)))
'Call LTD.C_PalletDoc.add(PD)
                           Set PD = Nothing
                           m_Rs3.MoveNext
                        Wend
                  Call LWH.C_LotDoc.add(LTD)
                  Set LTD = Nothing
                  m_Rs2.MoveNext
                  Wend
           Call Ua.C_LotItemsWH.add(LWH)
           Set LWH = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryInventoryWhDocForLG = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryWhDocForLG = False
End Function
Public Function QueryInventoryWhDocForIG(Ua As CInventoryWHDoc, C_IWD As Collection, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim LWH As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim I As Long
   RName = "QueryInventoryWhDocForIG"
   QueryInventoryWhDocForIG = False
   
   IsOK = True
   Call Ua.QueryData(5, Rs, ItemCount)
   Ua.Flag = "I"
   If Ua.QueryFlag = 1 Then
   Set Ua = Nothing
    While Not Rs.EOF
         Set LWH = New CLotItemWH
         Set Ua = New CInventoryWHDoc
         Call Ua.PopulateFromRS(5, Rs)
         LWH.LOT_ITEM_WH_ID = -1
         LWH.INVENTORY_WH_DOC_ID = Ua.INVENTORY_WH_DOC_ID
         LWH.OrderBy = 1
         Call LWH.QueryData(1, m_Rs1, iCount)
         Set LWH = Nothing
         While Not m_Rs1.EOF
            Set LWH = New CLotItemWH
            Call LWH.PopulateFromRS(1, m_Rs1)
                  LWH.Flag = "I"
                  Set LTD = New CLotDoc
                  LTD.LOT_ITEM_WH_ID = LWH.LOT_ITEM_WH_ID
                  LTD.OrderBy = 1
                  
                  Call LTD.QueryData(2, m_Rs2, iCount)
                  Set LTD = Nothing
                     While Not m_Rs2.EOF
                      Set LTD = New CLotDoc
                     Call LTD.PopulateFromRS(4, m_Rs2)
                     LTD.Flag = "I"
                        Set PD = New CPalletDoc
                        PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                        PD.OrderBy = 2
                        PD.OrderType = 1
                        Call PD.QueryData(2, m_Rs3, iCount)
                        Set PD = Nothing
                           While Not m_Rs3.EOF
                              Set PD = New CPalletDoc
                              Call PD.PopulateFromRS(2, m_Rs3)
                              PD.Flag = "I"
                              Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(LTD.LOT_ID)))
                              Set PD = Nothing
                              m_Rs3.MoveNext
                           Wend
                     Call LWH.C_LotDoc.add(LTD)
                     Set LTD = Nothing
                     m_Rs2.MoveNext
                     Wend
              Call Ua.C_LotItemsWH.add(LWH)
              Set LWH = Nothing
            m_Rs1.MoveNext
         Wend
         Call C_IWD.add(Ua)
         Set Ua = Nothing
         Rs.MoveNext
      Wend
   End If
   
   QueryInventoryWhDocForIG = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryWhDocForIG = False
End Function
Public Function QueryInventoryWhDocForPG(Ua As CInventoryWHDoc, C_IWD As Collection, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim LWH As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim I As Long
Dim SumPallet As Double
Dim PartNo As String
Dim LotNo As String

   RName = "QueryInventoryWhDocForPG"
   QueryInventoryWhDocForPG = False

   PartNo = Ua.PART_NO
   LotNo = Ua.LOT_NO
   IsOK = True
   Call Ua.QueryData(6, Rs, ItemCount)
   Ua.Flag = "I"
   If Ua.QueryFlag = 1 Then
   Set Ua = Nothing
    While Not Rs.EOF
         Set LWH = New CLotItemWH
         Set Ua = New CInventoryWHDoc
         Call Ua.PopulateFromRS(6, Rs)
         LWH.LOT_ITEM_WH_ID = -1
         LWH.INVENTORY_WH_DOC_ID = Ua.INVENTORY_WH_DOC_ID
         LWH.PART_NO = PartNo
         LWH.LOT_NO = LotNo
         LWH.OrderBy = 1
         Call LWH.QueryData(12, m_Rs1, iCount)
         Set LWH = Nothing
         While Not m_Rs1.EOF
            Set LWH = New CLotItemWH
            Call LWH.PopulateFromRS(12, m_Rs1)
                  LWH.Flag = "I"
                  Set LTD = New CLotDoc
                  LTD.LOT_DOC_ID = -1
                  LTD.LOT_ITEM_WH_ID = LWH.LOT_ITEM_WH_ID
                  LTD.OrderBy = 1
                  Call LTD.QueryData(9, m_Rs2, iCount)
                  Set LTD = Nothing
                     While Not m_Rs2.EOF
                      Set LTD = New CLotDoc
                     Call LTD.PopulateFromRS(9, m_Rs2)
                     LTD.Flag = "I"
                        Set PD = New CPalletDoc
                        PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                        PD.OrderBy = 2
                        PD.OrderType = 1
                        Call PD.QueryData(5, m_Rs3, iCount)
                         SumPallet = 0
                        Set PD = Nothing
                           While Not m_Rs3.EOF
                              Set PD = New CPalletDoc
                              Call PD.PopulateFromRS(5, m_Rs3)
                              PD.Flag = "I"
                              SumPallet = SumPallet + PD.CAPACITY_AMOUNT
                              Call LTD.C_PalletDoc.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(LTD.LOT_ID)))
                              Set PD = Nothing
                              m_Rs3.MoveNext
                           Wend
                     LTD.CAPACITY_AMOUNT = SumPallet
                     Call LWH.C_LotDoc.add(LTD)
                     Set LTD = Nothing
                     m_Rs2.MoveNext
                     Wend
              Call Ua.C_LotItemsWH.add(LWH)
              Set LWH = Nothing
            m_Rs1.MoveNext
         Wend
         Call C_IWD.add(Ua)
         Set Ua = Nothing
         Rs.MoveNext
      Wend
   End If
   
   QueryInventoryWhDocForPG = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryWhDocForPG = False
End Function
Public Function DeleteInventoryDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CInventoryDoc

   RName = "DeleteInventoryDoc"
   DeleteInventoryDoc = False
   HasBegin = False

   Set B = New CInventoryDoc

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   B.INVENTORY_DOC_ID = UID
   Call B.DeleteData
   
   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteInventoryDoc = True
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

   DeleteInventoryDoc = True
End Function
Public Function DeleteInventoryWhDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CInventoryWHDoc

   RName = "DeleteInventoryWhDoc"
   DeleteInventoryWhDoc = False
   HasBegin = False

   Set B = New CInventoryWHDoc

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   B.INVENTORY_WH_DOC_ID = UID
   Call B.DeleteData
   
   
   
   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteInventoryWhDoc = True
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

   DeleteInventoryWhDoc = False
End Function
Public Function AddEditInventoryDoc1(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim GuiID As Long
   
   RName = "AddEditInventoryDoc1"
   AddEditInventoryDoc1 = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Yw As CLotItem
   For Each O In Ug.ImportExports
      
      If O.TX_TYPE = "I" Then 'Import
         Set Yw = O
         If Yw.Flag = "A" Then
            Set PL = New CPartLocation
            PL.PART_LOCATION_ID = -1
            PL.PART_ITEM_ID = Yw.PART_ITEM_ID
            PL.LOCATION_ID = Yw.LOCATION_ID
            Call PL.QueryData(1, m_Rs2, iCount)
            If m_Rs2.EOF Then
               PL.AddEditMode = SHOW_ADD
               Call PL.AddEditData
            Else
               Call PL.PopulateFromRS(m_Rs2)
            End If
                                       
            If (PL.CURRENT_AMOUNT + Yw.TX_AMOUNT) > 0 Then
               NewPrice = (PL.CURRENT_AMOUNT * PL.AVG_PRICE + Yw.INCLUDE_UNIT_PRICE * Yw.TX_AMOUNT) / (PL.CURRENT_AMOUNT + Yw.TX_AMOUNT)
            Else
               NewPrice = 0
            End If

            Yw.NEW_AMOUNT = PL.CURRENT_AMOUNT + Yw.TX_AMOUNT
            If Yw.CALCULATE_FLAG = "Y" Then
               Yw.NEW_PRICE = NewPrice
            Else
               Yw.NEW_PRICE = PL.AVG_PRICE
            End If
            Yw.PREVIOUS_AMOUNT = PL.CURRENT_AMOUNT
            Yw.PREVIOUS_PRICE = PL.AVG_PRICE
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.GUI_ID = GuiID
            Yw.AddEditMode = SHOW_ADD
            Call Yw.AddEditData
            
            PL.AddEditMode = SHOW_EDIT
            If Yw.CALCULATE_FLAG = "Y" Then
               PL.AVG_PRICE = NewPrice
               PL.LAST_PRICE = Yw.INCLUDE_UNIT_PRICE
            End If
            PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + Yw.TX_AMOUNT
            If Ug.COMMIT_FLAG = "Y" Then
               Call PL.AddEditData
            End If
            
            Set PL = Nothing
         ElseIf Yw.Flag = "E" Then
            Set PL = New CPartLocation
            PL.PART_LOCATION_ID = -1
            PL.PART_ITEM_ID = Yw.PART_ITEM_ID
            PL.LOCATION_ID = Yw.LOCATION_ID
            Call PL.QueryData(1, m_Rs2, iCount)
            If m_Rs2.EOF Then
               PL.AddEditMode = SHOW_ADD
               Call PL.AddEditData
            Else
               Call PL.PopulateFromRS(m_Rs2)
            End If
                                       
            If (PL.CURRENT_AMOUNT + Yw.TX_AMOUNT) > 0 Then
               NewPrice = (PL.CURRENT_AMOUNT * PL.AVG_PRICE + Yw.INCLUDE_UNIT_PRICE * Yw.TX_AMOUNT) / (PL.CURRENT_AMOUNT + Yw.TX_AMOUNT)
            Else
               NewPrice = 0
            End If

            Yw.NEW_AMOUNT = PL.CURRENT_AMOUNT + Yw.TX_AMOUNT
            If Yw.CALCULATE_FLAG = "Y" Then
               Yw.NEW_PRICE = NewPrice
            Else
               Yw.NEW_PRICE = PL.AVG_PRICE
            End If
            Yw.PREVIOUS_AMOUNT = PL.CURRENT_AMOUNT
            Yw.PREVIOUS_PRICE = PL.AVG_PRICE
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.AddEditMode = SHOW_EDIT
            Call Yw.AddEditData
            
            PL.AddEditMode = SHOW_EDIT
            If Yw.CALCULATE_FLAG = "Y" Then
               PL.AVG_PRICE = NewPrice
               PL.LAST_PRICE = Yw.INCLUDE_UNIT_PRICE
            End If
            PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + Yw.TX_AMOUNT
            If Ug.COMMIT_FLAG = "Y" Then
               Call PL.AddEditData
            End If
            
            Set PL = Nothing
         ElseIf Yw.Flag = "D" Then
            Call Yw.DeleteData
         End If
      ElseIf O.TX_TYPE = "E" Then 'Export
      End If
   Next O

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditInventoryDoc1 = True
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
   AddEditInventoryDoc1 = False
End Function

Private Sub AddSubLotItem(Li As CLotItem)
Dim D As CSubLotItem
Dim id As Long

   Set D = New CSubLotItem
   D.LOT_ITEM_ID = Li.LOT_ITEM_ID
   Call D.DeleteAllData
   Set D = Nothing
   
   For Each D In Li.SubLotItems
      If D.Flag <> "D" Then
         D.LOT_ITEM_ID = Li.LOT_ITEM_ID
         D.AddEditMode = SHOW_ADD
         Call D.AddEditData
      ElseIf D.Flag = "D" Then
         Call D.DeleteData
      End If
   Next D
End Sub

Private Function DecreaseLotAmount(Li As CLotItem)
Dim C As CSubLotItem
Dim TempLi As CLotItem

   Set TempLi = New CLotItem
   
   For Each C In Li.SubLotItems
      If C.Flag <> "D" Then
         TempLi.LOT_ITEM_ID = C.FROM_LOT_ID
         Call TempLi.DecreaseLeftAmount(C.SUBLOT_AMT)
      End If
   Next C
   
   Set TempLi = Nothing
End Function

Private Function ProcessIDToInventoryDoc(ProcessID As Long) As Long
Dim TempID As Long

   'Hard code here
   If ProcessID = 2 Then 'บรรจุ
      TempID = 14
   ElseIf ProcessID = 4 Then 'ผสม bulk
      TempID = 13
   ElseIf ProcessID = 5 Then 'ผสม premix
      TempID = 12
   ElseIf ProcessID = 6 Then 'รับ BAG จากการ RE-BAG to BAG
      TempID = 17
   ElseIf ProcessID = 7 Then 'รับ BAG จากการ RE-BAG to BULK
      TempID = 18
   ElseIf ProcessID = 8 Then 'รับ BAG จากการ RE-BAG to RM
      TempID = 19
   End If
   
   ProcessIDToInventoryDoc = TempID
End Function
Private Function ProcessIDToInventoryDocInput(ProcessID As Long) As Long
Dim TempID As Long

   'Hard code here
 If ProcessID = 2 Then 'จ่าย Bulk เพื่อ Pack bag
      TempID = 2003
 ElseIf ProcessID = 6 Or ProcessID = 7 Or ProcessID = 8 Then      'โอน BAG เข้าผลิต เพื่อ RE-BAG
      TempID = 2002
'ElseIf ProcessID = 7 Then  'โอน BAG เข้าผลิต เพื่อ RE-Bulk
'      TempID = 2004
'ElseIf ProcessID = 8 Then  'โอน BAG เข้าผลิต เพื่อ RE-Rm
'      TempID = 2005
End If
   
   ProcessIDToInventoryDocInput = TempID
End Function

Public Function AddEditInventoryDoc(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim GuiID As Long

   RName = "AddEditInventoryDoc"
   AddEditInventoryDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Yw As CLotItem
   For Each O In Ug.ImportExports
      Set Yw = O
      
      If O.TX_TYPE = "I" Then 'Import
         Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
         Yw.GUI_ID = GuiID
         Yw.LEFT_AMOUNT = Yw.TX_AMOUNT
         
         If Yw.Flag = "A" Then
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.AddEditMode = SHOW_ADD
            Call Yw.AddEditData
         ElseIf Yw.Flag = "E" Then
            Yw.AddEditMode = SHOW_EDIT
            Call Yw.AddEditData
         ElseIf Yw.Flag = "D" Then
            Call Yw.DeleteData
         End If
      ElseIf (O.TX_TYPE = "E") Or (O.TX_TYPE = "A") Then 'Export
         Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
         Call glbDatabaseMngr.GetSeqID("GUI_ID_SEQ", GuiID, ErrorObj)
         Yw.GUI_ID = GuiID
         Yw.LEFT_AMOUNT = -1
         If Yw.Flag = "A" Then
            Yw.AddEditMode = SHOW_ADD
            Call Yw.AddEditData
         ElseIf Yw.Flag = "E" Then
            Yw.AddEditMode = SHOW_EDIT
            Call Yw.AddEditData
         ElseIf Yw.Flag = "D" Then
            Call Yw.DeleteData
         End If
      End If
   Next O

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditInventoryDoc = True
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
   AddEditInventoryDoc = False
End Function
Public Function AddEditInventoryWhDoc(Ug As CInventoryWHDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim GuiID As Long
Dim m_CollPallet As Collection
Dim m_Pallets As CPalletDoc
Dim m_InventoryWHDoc As CInventoryWHDoc
Dim m_LotItemWh As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim SumGoodAmount As Long
   Set m_CollPallet = New Collection
   
   RName = "AddEditInventoryWhDoc"
   AddEditInventoryWhDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   If Ug.Flag = "D" Then
      Call Ug.DeleteData
   Else
      Call Ug.AddEditData
   End If

   Dim Cd As CConfigDoc
      If Ug.AddEditMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ug.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ug.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ug.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
      End If
      
   For Each m_LotItemWh In Ug.C_LotItemsWH
   SumGoodAmount = 0
      
    If m_LotItemWh.Flag = "A" Then
      m_LotItemWh.INVENTORY_WH_DOC_ID = Ug.INVENTORY_WH_DOC_ID
      m_LotItemWh.CALCULATE_FLAG = "N"
      m_LotItemWh.AddEditMode = SHOW_ADD
      Call m_LotItemWh.AddEditData
   ElseIf m_LotItemWh.Flag = "E" Then
      m_LotItemWh.AddEditMode = SHOW_EDIT
      Call m_LotItemWh.AddEditData
   ElseIf m_LotItemWh.Flag = "D" Then
      Call m_LotItemWh.DeleteData
   End If
   
   If Not m_LotItemWh.C_LotDoc Is Nothing Then
      For Each LTD In m_LotItemWh.C_LotDoc
           If LTD.Flag = "A" Then
             LTD.LOT_ITEM_WH_ID = m_LotItemWh.LOT_ITEM_WH_ID
             LTD.AddEditMode = SHOW_ADD
             Call LTD.AddEditData
             LTD.LOT_DOC_ID = LTD.LOT_DOC_ID
           ElseIf LTD.Flag = "E" Then
             LTD.AddEditMode = SHOW_EDIT
             Call LTD.AddEditData
             If LTD.IS_EDIT_LOT Then
               Call LTD.AddEditLot
             End If
             
             LTD.AddEditBinNo 'UPDATE ถังบรรจุที่จ่ายออกไปด้วย
           ElseIf LTD.Flag = "D" Then
            Call LTD.DeleteData
           End If
           
            For Each PD In LTD.C_PalletDoc
                   If PD.Flag = "A" Then
                      PD.AddEditMode = SHOW_ADD
                     PD.LOT_DOC_ID = LTD.LOT_DOC_ID
                     SumGoodAmount = SumGoodAmount + PD.CAPACITY_AMOUNT
                     Call PD.AddEditData
                  ElseIf PD.Flag = "E" Then
                     PD.AddEditMode = SHOW_EDIT
                     Call PD.AddEditData
                   ElseIf PD.Flag = "D" Then
                     Call PD.DeleteData
                  End If
            Next PD
      Next LTD
     End If
   Next m_LotItemWh

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditInventoryWhDoc = True
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
   AddEditInventoryWhDoc = False
End Function
Public Function AddBalancePart(Ug As CInventoryWHDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional id As Long = 0)
On Error GoTo ErrorHandler
Dim Rs As ADODB.Recordset
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim m_LotItemWh As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc

   Set Rs = New ADODB.Recordset
   
   RName = "AddBalancePart"
   AddBalancePart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
       For Each m_LotItemWh In Ug.C_LotItemsWH
         If m_LotItemWh.Flag = "A" Then
         Set LTD = New CLotDoc
         LTD.PART_ITEM_ID = m_LotItemWh.PART_ITEM_ID
         LTD.BALANCE_FLAG = "N"
         LTD.TX_TYPE = "I" 'เลือกเฉพาะที่รับเข้า
         LTD.IS_BALANCE = 1 'บอกว่าเป็นการปรับยอด
         LTD.LOT_DOC_ID_REF = m_LotItemWh.C_LotDoc(1).LOT_DOC_ID
         LTD.DOCUMENT_DATE = Ug.DOCUMENT_DATE
         LTD.DOCUMENT_TYPE = Ug.DOCUMENT_TYPE
         Call LTD.QueryData(8, Rs, iCount)
         Set LTD = Nothing
           While Not Rs.EOF
                  Set LTD = New CLotDoc
                  Call LTD.PopulateFromRS(8, Rs)
                  LTD.AddEditMode = SHOW_EDIT
                  LTD.LOT_DOC_ID_REF = m_LotItemWh.C_LotDoc(1).LOT_DOC_ID
                  LTD.BALANCE_FLAG = "Y"
                  Call LTD.AddEditBalance
                  Set LTD = Nothing
                  Rs.MoveNext
            Wend
         End If
      Next m_LotItemWh
            
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddBalancePart = True
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
   AddBalancePart = False
End Function
Public Function EditBalancePart(Ug As CInventoryWHDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional id As Long = 0)
On Error GoTo ErrorHandler
Dim Rs As ADODB.Recordset
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim m_LotItemWh As CLotItemWH
Dim Lt As cLot
Dim LTD As CLotDoc

   Set Rs = New ADODB.Recordset
   
   RName = "EditBalancePart"
   EditBalancePart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
       For Each m_LotItemWh In Ug.C_LotItemsWH
         If m_LotItemWh.Flag = "D" Then
            Set LTD = New CLotDoc
            LTD.AddEditMode = SHOW_EDIT
            LTD.LOT_DOC_ID = m_LotItemWh.C_LotDoc(1).LOT_DOC_ID
            LTD.BALANCE_FLAG = "N"
            Call LTD.AddEditBalanceReFlag
            Set LTD = Nothing
         End If
      Next m_LotItemWh
            
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   EditBalancePart = True
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
   EditBalancePart = False
End Function
'==
Public Function VerifyStockBalance(Col2 As Collection, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim O As Object
Dim iCount As Long
Dim PL As CPartLocation
Dim PartLocations As Collection
Dim FoundFlag As Boolean
Dim TempCol As Collection
Dim TempLi As CLotItem
Dim C As CSubLotItem

   RName = "VerifyStockBalance"
   VerifyStockBalance = False
   HasBegin = False

   Set PartLocations = New Collection
   Set TempCol = New Collection
   Set TempLi = New CLotItem
   
   TempLi.LOT_ITEM_ID = -1
   TempLi.COMMIT_FLAG = "Y"
   TempLi.TX_TYPE = "I"
   TempLi.OrderType = 1
   Call LoadSubLotItem(Nothing, TempCol, TempLi)
   
   Set PL = New CPartLocation
   PL.PART_LOCATION_ID = -1
   PL.PART_ITEM_ID = -1
   PL.LOCATION_ID = -1
   Call PL.QueryData(1, m_Rs1, iCount)
   Set PL = Nothing
   While Not m_Rs1.EOF
      Set PL = New CPartLocation
      Call PL.PopulateFromRS(m_Rs1)
      
      Call PartLocations.add(PL)
      
      Set PL = Nothing
      m_Rs1.MoveNext
   Wend
   Set PL = Nothing
   
   FoundFlag = False
   If Not (Col2 Is Nothing) Then
      For Each O In Col2
         If O.TX_TYPE = "E" Then
            For Each PL In PartLocations
               If (PL.PART_ITEM_ID = O.PART_ITEM_ID) And (PL.LOCATION_ID = O.LOCATION_ID) Then
                  If O.TX_TYPE = "E" Then
                     PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - O.TX_AMOUNT
                  Else
                     PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + O.TX_AMOUNT
                  End If
                  FoundFlag = True
                  
                  If PL.CURRENT_AMOUNT < 0 Then
                     FoundFlag = False
                  End If
                  Exit For
               End If
            Next PL
         Else
            FoundFlag = True
         End If
                  
         If Not FoundFlag Then
            glbErrorLog.LocalErrorMsg = MapText("มีปริมาณ ") & O.PART_DESC & " " & MapText(" ไม่เพียงพอ ที่คลัง ") & O.LOCATION_NAME
            glbErrorLog.ShowUserError
            
            VerifyStockBalance = False
            Exit Function
         End If
         
'         For Each TempLi In TempCol
'            For Each C In O.SubLotItems
'               If (C.Flag <> "D") And (C.FROM_LOT_ID = TempLi.LOT_ITEM_ID) Then
'                  If C.SUBLOT_AMT > TempLi.LEFT_AMOUNT Then
'                     glbErrorLog.LocalErrorMsg = MapText("มีปริมาณ ") & O.PART_DESC & " " & MapText(" ไม่เพียงพอ ที่ล็อต ") & C.DOCUMENT_NO
'                     glbErrorLog.ShowUserError
'
'                     VerifyStockBalance = False
'                     Exit Function
'                  Else
'                     TempLi.LEFT_AMOUNT = TempLi.LEFT_AMOUNT - C.SUBLOT_AMT
'                  End If
'               End If
'            Next C
'         Next TempLi
         
      Next O
   End If
   
   Set TempLi = Nothing
   Set TempCol = Nothing
   Set PartLocations = Nothing
   
   VerifyStockBalance = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   VerifyStockBalance = False
End Function

Private Function CompareKey(D1 As Object, D2 As Object) As Boolean
Dim TempResult As Boolean

   TempResult = D1.TRANSACTION_SEQ < D2.TRANSACTION_SEQ
         
   CompareKey = TempResult
End Function

Private Function CompareKeyEx(D1 As Object, D2 As Object) As Boolean
Dim TempResult As Boolean

   TempResult = D1.TRANSACTION_SEQ < D2.TRANSACTION_SEQ
         
   CompareKeyEx = TempResult
End Function

Private Sub Selectionsort(List As Collection, MIN As Long, MAX As Long)
Dim I As Long
Dim J As Long
Dim best_value As Object
Dim Temp As Object
Dim best_j As Integer

    For I = MIN To MAX - 1
        Set best_value = List(I)
        best_j = I
        For J = I + 1 To MAX
            If CompareKeyEx(List(J), best_value) Then
                Set best_value = List(J)
                best_j = J
            End If
        Next J
        
        Set Temp = List(I)
        List.Remove (best_j)
        If best_j > List.Count Then
         Call List.add(Temp, , , best_j - 1)
      Else
         Call List.add(Temp, , best_j)
      End If
    
        List.Remove (I)
        If I > List.Count Then
         Call List.add(best_value, , , I - 1)
      Else
         Call List.add(best_value, , I)
      End If
    Next I
    Set best_value = Nothing
End Sub


Public Function DateToPartItem(D1 As Date) As CPartItem
Dim Count As Long
Dim Pw As CYearWeek
Dim iCount As Long
Dim p As CPartItem

   Set Pw = New CYearWeek
   Pw.YEAR_WEEK_ID = -1
   Pw.YEAR_SEQ_ID = -1
   
   Call Pw.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      If (DateToStringInt(D1) >= NVLS(m_Rs1("FROM_DATE"), "")) And (DateToStringInt(D1) <= NVLS(m_Rs1("TO_DATE"), "")) Then
         Set p = New CPartItem
         p.PART_ITEM_ID = NVLI(m_Rs1("PART_ITEM_ID6"), -1)
         If p.PART_ITEM_ID > 0 Then
            Call p.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call p.PopulateFromRS(1, m_Rs2)
            End If
            Set DateToPartItem = p
         End If
         
         Exit Function
      End If
      
      m_Rs1.MoveNext
   Wend
   
   Set Pw = Nothing
   Set DateToPartItem = Nothing
End Function

'==
Public Function QueryBillingDoc(Ua As CBillingDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
   
   RName = "QueryBillingDoc"
   QueryBillingDoc = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Yw As CDoItem
      
      Set Yw = New CDoItem
      Yw.DO_ITEM_ID = -1
      Yw.DO_ID = Ua.BILLING_DOC_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.DoItems = Nothing
      Set Ua.DoItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CDoItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.DoItems.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim Sp As CSupItem
      Set Sp = New CSupItem
      Sp.SUP_ITEM_ID = -1
      Sp.DO_ID = Ua.BILLING_DOC_ID
      Sp.OrderBy = -1
      Sp.OrderType = 1
      Call Sp.QueryData(1, m_Rs1, iCount)
      Set Sp = Nothing
      
      Set Ua.SupItems = Nothing
      Set Ua.SupItems = New Collection
      While Not m_Rs1.EOF
         Set Sp = New CSupItem
         Call Sp.PopulateFromRS(1, m_Rs1)
                  
         Sp.Flag = "I"
         Call Ua.SupItems.add(Sp)
         
         Set Sp = Nothing
         m_Rs1.MoveNext
      Wend
      
      '*******************
      Dim SR As CSupItem
      
      Set SR = New CSupItem
      SR.SUP_ITEM_ID = -1
      SR.PO_ID = Ua.BILLING_DOC_ID
      SR.OrderBy = 1
      SR.OrderType = 1
      Call SR.QueryData(113, m_Rs4, iCount)
      Set SR = Nothing
      
      Set Ua.SupReceiptItems = Nothing
      Set Ua.SupReceiptItems = New Collection
      While Not m_Rs4.EOF
         Set SR = New CSupItem
         Call SR.PopulateFromRS(113, m_Rs4)
                  
         SR.Flag = "I"
         Call Ua.SupReceiptItems.add(SR)
         
         Set SR = Nothing
         m_Rs4.MoveNext
      Wend
      '*******************
      
      Dim SO As CSaleOrder
      
      Set SO = New CSaleOrder
      SO.SALE_ORDER_ID = -1
      SO.DO_ID = Ua.BILLING_DOC_ID
      If Ua.DOCUMENT_TYPE_WH = 2000 Or Ua.DOCUMENT_TYPE_WH = 2001 Or Ua.DOCUMENT_TYPE_WH = 2004 Then
         SO.PART_TYPE = Ua.PART_TYPE
         SO.PART_TYPE_OVER = Ua.PART_TYPE_OVER
      End If
      SO.OrderBy = -1
      SO.OrderType = 1
      Call SO.QueryData(1, m_Rs1, iCount)
      Set SO = Nothing
      
      Set Ua.SaleOrders = Nothing
      Set Ua.SaleOrders = New Collection
      While Not m_Rs1.EOF
         Set SO = New CSaleOrder
         Call SO.PopulateFromRS(1, m_Rs1)
         
         SO.Flag = "I"
         Call Ua.SaleOrders.add(SO)
         
         Set SO = Nothing
         m_Rs1.MoveNext
      Wend
      
      
      Dim RT As CReceiptItem
      
      Set RT = New CReceiptItem
      RT.RECEIPT_ITEM_ID = -1
      RT.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      RT.OrderBy = -1
      RT.OrderType = 1
      Call RT.QueryData(1, m_Rs1, iCount)
      Set RT = Nothing
      
      Set Ua.ReceiptItems = Nothing
      Set Ua.ReceiptItems = New Collection
      While Not m_Rs1.EOF
         Set RT = New CReceiptItem
         Call RT.PopulateFromRS(1, m_Rs1)
                  
         RT.Flag = "I"
         Call Ua.ReceiptItems.add(RT)
         
         Set RT = Nothing
         m_Rs1.MoveNext
      Wend
      
      
        Dim RCD As CReceiptChequeDoc
      
      Set RCD = New CReceiptChequeDoc
      RCD.RECEIPT_CHEQUE_DOC_ID = -1
      RCD.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      RCD.OrderBy = -1
      RCD.OrderType = 1
      Call RCD.QueryData(1, m_Rs1, iCount)
      Set RCD = Nothing
      
      Set Ua.ReceiptChequeDocItems = Nothing
      Set Ua.ReceiptChequeDocItems = New Collection
      While Not m_Rs1.EOF
         Set RCD = New CReceiptChequeDoc
         Call RCD.PopulateFromRS(1, m_Rs1)
                  
         RCD.Flag = "I"
         Call Ua.ReceiptChequeDocItems.add(RCD)
         
         Set RCD = Nothing
         m_Rs1.MoveNext
      Wend
      
      
   
     Dim Rtt As CPkglst
      
      Set Rtt = New CPkglst
      Rtt.PKGLST_ITEM_ID = -1
      Rtt.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Rtt.OrderBy = -1
      Rtt.OrderType = 1
      Call Rtt.QueryData(1, m_Rs1, iCount)
      Set Rtt = Nothing
      
      Set Ua.Pkglsts = Nothing
      Set Ua.Pkglsts = New Collection
      While Not m_Rs1.EOF
         Set Rtt = New CPkglst
         Call Rtt.PopulateFromRS(1, m_Rs1)
                  
         Rtt.Flag = "I"
         Call Ua.Pkglsts.add(Rtt)
         
         Set Rtt = Nothing
         m_Rs1.MoveNext
      Wend

    Dim Meb As CMemoBank
      
      Set Meb = New CMemoBank
      Meb.MEMO_BANK_ID = -1
      Meb.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Meb.OrderBy = -1
      Meb.OrderType = 1
      Call Meb.QueryData(1, m_Rs1, iCount)
      Set Meb = Nothing
      
      Set Ua.MemoBank = Nothing
      Set Ua.MemoBank = New Collection
      While Not m_Rs1.EOF
         Set Meb = New CMemoBank
         Call Meb.PopulateFromRS(1, m_Rs1)
                  
         Meb.Flag = "I"
         Call Ua.MemoBank.add(Meb)
         
         Set Meb = Nothing
         m_Rs1.MoveNext
      Wend

      Dim Bdc As CBillingDiscount
      Set Bdc = New CBillingDiscount
      Bdc.BILLING_DISCOUNT_ID = -1
      Bdc.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Bdc.OrderBy = -1
      Bdc.OrderType = 1
      Call Bdc.QueryData(1, m_Rs1, iCount)
      Set Bdc = Nothing
      
      Set Ua.BillingDiscounts = Nothing
      Set Ua.BillingDiscounts = New Collection
      While Not m_Rs1.EOF
         Set Bdc = New CBillingDiscount
         Call Bdc.PopulateFromRS(1, m_Rs1)
                  
         Bdc.Flag = "I"
         Call Ua.BillingDiscounts.add(Bdc)
         
         Set Bdc = Nothing
         m_Rs1.MoveNext
      Wend
   
      Dim Bh As CBulkHole
      Set Bh = New CBulkHole
      Bh.BULK_HOLE_ID = -1
      Bh.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Bh.OrderBy = -1
      Bh.OrderType = 1
      Call Bh.QueryData(1, m_Rs1, iCount)
      Set Bh = Nothing
      
      Set Ua.BulkHoles = Nothing
      Set Ua.BulkHoles = New Collection
      While Not m_Rs1.EOF
         Set Bh = New CBulkHole
         Call Bh.PopulateFromRS(1, m_Rs1)
                  
         Bh.Flag = "I"
         Call Ua.BulkHoles.add(Bh)
         
         Set Bh = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim BT As CBillTransport
      Set BT = New CBillTransport
      BT.BILL_TRANSPORT_ID = -1
      BT.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      BT.OrderBy = 1
      BT.OrderType = 1
      Call BT.QueryData(1, m_Rs1, iCount)
      Set BT = Nothing
      
      Set Ua.BillTransport = Nothing
      Set Ua.BillTransport = New Collection
      Dim BTI As CBillTransportItem
      While Not m_Rs1.EOF
         Set BT = New CBillTransport
         Call BT.PopulateFromRS(1, m_Rs1)
         
         Set BTI = New CBillTransportItem
         BTI.BILL_TRANSPORT_ITEM_ID = -1
         BTI.BILL_TRANSPORT_ID = BT.BILL_TRANSPORT_ID
         BTI.OrderBy = 1
         BTI.OrderType = 1
         Call BTI.QueryData(1, m_Rs2, iCount)
         Set BTI = Nothing
         
         Set BT.C_BillTransportItem = Nothing
         Set BT.C_BillTransportItem = New Collection
         While Not m_Rs2.EOF
            Set BTI = New CBillTransportItem
            Call BTI.PopulateFromRS(1, m_Rs2)
            BTI.Flag = "I"
            Call BT.C_BillTransportItem.add(BTI)
            Set BTI = Nothing
            m_Rs2.MoveNext
         Wend
         
         BT.Flag = "I"
         Call Ua.BillTransport.add(BT)
         
         Set BT = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.Payments = Nothing
      Set Ua.Payments = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHECK_ID") > 0 Then
            If Ct.GetFieldValue("SUPPLIER_ID") > 0 Then
               Call Ct.Cheque.SetFieldValue("DIRECTION", 2)
            End If
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.Payments.add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Gl As CGLDetail
      Set Gl = New CGLDetail
      Call Gl.SetFieldValue("GL_DETAIL_ID", -1)
      Call Gl.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Gl.QueryData(1, m_Rs1, iCount)
      Set Gl = Nothing
      
      Set Ua.GlDetails = Nothing
      Set Ua.GlDetails = New Collection
      While Not m_Rs1.EOF
         Set Gl = New CGLDetail
         Call Gl.PopulateFromRS(1, m_Rs1)
                  
         Gl.Flag = "I"
         Call Ua.GlDetails.add(Gl)
         
         Set Gl = Nothing
         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryBillingDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBillingDoc = False
End Function
Public Function QueryBillingPayment(Ua As CBillingPayment, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
   
   RName = "QueryBillingPayment"
   QueryBillingPayment = False
   
   IsOK = True
   If Len(Ua.DOCUMENT_NO) > 0 Then
      Call Ua.QueryData(1, Rs, ItemCount)
   ElseIf Len(Ua.DOCUMENT_NO_JV) > 0 Then
      Ua.DOCUMENT_TYPE = 112
      Call Ua.QueryData(2, Rs, ItemCount)
   Else
     Call Ua.QueryData(1, Rs, ItemCount)
   End If
   
   If Ua.QueryFlag = 1 Then
   
      Dim Gl As CGLDetail
      Set Gl = New CGLDetail
      Call Gl.SetFieldValue("GL_DETAIL_ID", -1)
      If Ua.BILLING_PAYMENT_ID < 0 Then
         Ua.BILLING_PAYMENT_ID = 0
      End If
      Call Gl.SetFieldValue("BILLING_PAYMENT_ID", Ua.BILLING_PAYMENT_ID)
      Call Gl.QueryData(1, m_Rs1, iCount)
      Set Gl = Nothing

      Set Ua.GlDetails = Nothing
      Set Ua.GlDetails = New Collection
      While Not m_Rs1.EOF
         Set Gl = New CGLDetail
         Call Gl.PopulateFromRS(1, m_Rs1)

         Gl.Flag = "I"
         Call Ua.GlDetails.add(Gl)

         Set Gl = Nothing
         m_Rs1.MoveNext
      Wend
      
    Dim Gl2 As CGLDetail
      Set Gl2 = New CGLDetail
      Call Gl2.SetFieldValue("GL_DETAIL_ID", -1)
      If Ua.BILLING_PAYMENT_ID_REF < 0 Then
         Ua.BILLING_PAYMENT_ID_REF = 0
      End If
      Call Gl2.SetFieldValue("BILLING_PAYMENT_ID", Ua.BILLING_PAYMENT_ID_REF)
      Call Gl2.QueryData(1, m_Rs1, iCount)
      Set Gl2 = Nothing

      Set Ua.GlDetails2 = Nothing
      Set Ua.GlDetails2 = New Collection
      While Not m_Rs1.EOF
         Set Gl2 = New CGLDetail
         Call Gl2.PopulateFromRS(1, m_Rs1)

         Gl2.Flag = "I"
         Call Ua.GlDetails2.add(Gl2)

         Set Gl2 = Nothing
         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryBillingPayment = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBillingPayment = False
End Function
Public Function DeleteBillingDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional PaymentID As Long = -1)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CBillingDoc

   RName = "DeleteBillingDoc"
   DeleteBillingDoc = False
   HasBegin = False

   Set B = New CBillingDoc

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.BILLING_DOC_ID = UID
   B.PAYMENT_ID = PaymentID
   Call B.QueryData(1, m_Rs1, ItemCount)
   
   If Not m_Rs1.EOF Then
      Call B.PopulateFromRS(1, m_Rs1)
   End If
   
   Call B.DeleteData
   If B.INVENTORY_DOC_ID > 0 Then
      Call glbDaily.DeleteInventoryDoc(B.INVENTORY_DOC_ID, IsOK, False, ErrorObj)
   End If
   
   
   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBillingDoc = True
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

   DeleteBillingDoc = True
End Function
Public Function DeleteBillingPayment(UID As Long, UID2 As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CBillingPayment

   RName = "DeleteBillingPayment"
   DeleteBillingPayment = False
   HasBegin = False

   Set B = New CBillingPayment

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
'ลบ PV
   B.BILLING_PAYMENT_ID = UID
   Call B.DeleteData

'ลบ JV
   B.BILLING_PAYMENT_ID = UID2
   Call B.DeleteData
   
   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBillingPayment = True
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

   DeleteBillingPayment = True
End Function
Public Function AddEditBillingDoc(Ug As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim PaymentType As Long
Dim OldPaymentType As Long


   RName = "AddEditBillingDoc"
   AddEditBillingDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   
   Dim Cd As CConfigDoc
   If Ug.AddEditMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ug.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ug.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ug.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   
   Dim CstName As CDoItem
   For Each CstName In Ug.DoItems
      If CstName.Flag = "A" Then
         CstName.DO_ID = Ug.BILLING_DOC_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Dim Sp As CSupItem
   For Each Sp In Ug.SupItems
      If Sp.Flag = "A" Then
         Sp.DO_ID = Ug.BILLING_DOC_ID
         Sp.AddEditMode = SHOW_ADD
         Sp.AddEditData
      ElseIf Sp.Flag = "E" Then
         Sp.AddEditMode = SHOW_EDIT
         Sp.AddEditData
      ElseIf Sp.Flag = "D" Then
         Sp.DeleteData
      End If
   Next Sp
   
   If Ug.DOCUMENT_TYPE = 100 Or Ug.DOCUMENT_TYPE = 101 Or Ug.DOCUMENT_TYPE = 102 Or Ug.DOCUMENT_TYPE = 103 Then
      
      
   End If
   
   Dim SO As CSaleOrder
   For Each SO In Ug.SaleOrders
      If SO.Flag = "A" Then
         SO.DO_ID = Ug.BILLING_DOC_ID
         SO.AddEditMode = SHOW_ADD
         SO.AddEditData
      ElseIf SO.Flag = "E" Then
         SO.AddEditMode = SHOW_EDIT
         SO.AddEditData
      ElseIf SO.Flag = "D" Then
         SO.DeleteData
      End If
   Next SO
   
   Dim RT As CReceiptItem
   For Each RT In Ug.ReceiptItems
      If RT.Flag = "A" Then
         RT.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         RT.AddEditMode = SHOW_ADD
         RT.AddEditData
      ElseIf RT.Flag = "E" Then
         RT.AddEditMode = SHOW_EDIT
         RT.AddEditData
      ElseIf RT.Flag = "D" Then
         RT.DeleteData
      End If
   Next RT
   
   
    Dim RCD As CReceiptChequeDoc
   For Each RCD In Ug.ReceiptChequeDocItems
      If RCD.Flag = "A" Then
         RCD.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         RCD.AddEditMode = SHOW_ADD
         RCD.AddEditData
      ElseIf RCD.Flag = "E" Then
         RCD.AddEditMode = SHOW_EDIT
         RCD.AddEditData
      ElseIf RCD.Flag = "D" Then
         RCD.DeleteData
      End If
   Next RCD
   
   Dim Rtt As CPkglst
   For Each Rtt In Ug.Pkglsts
      If Rtt.Flag = "A" Then
         Rtt.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Rtt.AddEditMode = SHOW_ADD
         Rtt.AddEditData
      ElseIf Rtt.Flag = "E" Then
         Rtt.AddEditMode = SHOW_EDIT
         Rtt.AddEditData
      ElseIf Rtt.Flag = "D" Then
         Rtt.DeleteData
      End If
   Next Rtt

   Dim Memo As CMemoBank
   For Each Memo In Ug.MemoBank
      If Memo.Flag = "A" Then
         Memo.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Memo.AddEditMode = SHOW_ADD
         Memo.AddEditData
      ElseIf Memo.Flag = "E" Then
         Memo.AddEditMode = SHOW_EDIT
         Memo.AddEditData
      ElseIf Memo.Flag = "D" Then
         Memo.DeleteData
      End If
   Next Memo

   Dim Bdc As CBillingDiscount
   For Each Bdc In Ug.BillingDiscounts
      If Bdc.Flag = "A" Then
         Bdc.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Bdc.AddEditMode = SHOW_ADD
         Bdc.AddEditData
      ElseIf Bdc.Flag = "E" Then
         Bdc.AddEditMode = SHOW_EDIT
         Bdc.AddEditData
      ElseIf Bdc.Flag = "D" Then
         Bdc.DeleteData
      End If
   Next Bdc
   
   Dim Bh As CBulkHole
   For Each Bh In Ug.BulkHoles
      If Bh.Flag = "A" Then
         Bh.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Bh.AddEditMode = SHOW_ADD
         Bh.AddEditData
      ElseIf Bh.Flag = "E" Then
         Bh.AddEditMode = SHOW_EDIT
         Bh.AddEditData
      ElseIf Bh.Flag = "D" Then
         Bh.DeleteData
      End If
   Next Bh
   
   Dim Ct As CCashTran
   For Each Ct In Ug.Payments
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ug.BILLING_DOC_ID)
      Call Ct.SetFieldValue("TX_NO", Ug.DOCUMENT_NO)
      Call Ct.SetFieldValue("TX_DATE", Ug.DOCUMENT_DATE)
      Call Ct.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
      Call Ct.SetFieldValue("SUPPLIER_ID", Ug.SUPPLIER_ID)
      
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then 'เช็ค
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
         End If
         Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = 3) Then
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", -1)
            Call Ct.AddEditData
         
            Call Ct.Cheque.DeleteData
         ElseIf (PaymentType = OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If Ct.GetFieldValue("OLD_CHECK_ID") > 0 Then 'เช็ค
            Call Ct.Cheque.DeleteData
         End If
      End If
   Next Ct

   'ทำเพื่อเซต APAR_ID ให้ตรงกันกับที่ BILLING_DOC
   For Each Ct In Ug.Payments
      If Ct.GetFieldValue("CHECK_ID") > 0 Then
         Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
         Call Ct.Cheque.SetFieldValue("APAR_ID", Ug.CUSTOMER_ID)
         Call Ct.Cheque.SetFieldValue("SUPPLIER_ID", Ug.SUPPLIER_ID)
         Call Ct.Cheque.UpdateApArID
      End If
   Next Ct
   
   Dim Gl As CGLDetail
   For Each Gl In Ug.GlDetails
      If Gl.Flag = "A" Then
         Call Gl.SetFieldValue("BILLING_DOC_ID", Ug.BILLING_DOC_ID)
         Gl.ShowMode = SHOW_ADD
         Gl.AddEditData
      ElseIf Gl.Flag = "E" Then
         Gl.ShowMode = SHOW_EDIT
         Gl.AddEditData
      ElseIf Gl.Flag = "D" Then
         Gl.DeleteData
      End If
   Next Gl
   
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditBillingDoc = True
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
   AddEditBillingDoc = False
End Function
Public Function AddEditBillingPayment(Ug As CBillingPayment, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional tType As String = "")
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim O As Object

   RName = "AddEditBillingPayment"
   AddEditBillingPayment = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   

   Dim Cd As CConfigDoc
   If tType = "PV" Then
   
      If Ug.AddEditMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ug.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ug.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ug.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
      End If
   
      Dim Gl As CGLDetail
      For Each Gl In Ug.GlDetails
         If Gl.Flag = "A" Then
            Call Gl.SetFieldValue("BILLING_PAYMENT_ID", Ug.BILLING_PAYMENT_ID)
            Gl.ShowMode = SHOW_ADD
            Gl.AddEditData
         ElseIf Gl.Flag = "E" Then
            Gl.ShowMode = SHOW_EDIT
            Gl.AddEditData
         ElseIf Gl.Flag = "D" Then
            Gl.DeleteData
         End If
      Next Gl
   End If
   
   If tType = "JV" Then
      
      If Ug.AddEditMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ug.RUNNING_NO2)
      Call Cd.SetFieldValue("LAST_NO", Ug.DOCUMENT_NO_JV)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ug.CONFIG_DOC_TYPE2)
      Call Cd.UpdateRunningNo
      End If
      
      Dim Gl2 As CGLDetail
      For Each Gl2 In Ug.GlDetails2
         If Gl2.Flag = "A" Then
            Call Gl2.SetFieldValue("BILLING_PAYMENT_ID", Ug.BILLING_PAYMENT_ID)
            Gl2.ShowMode = SHOW_ADD
            Gl2.AddEditData
         ElseIf Gl2.Flag = "E" Then
            Gl2.ShowMode = SHOW_EDIT
            Gl2.AddEditData
         ElseIf Gl2.Flag = "D" Then
            Gl2.DeleteData
         End If
      Next Gl2
   End If
 
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditBillingPayment = True
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
   AddEditBillingPayment = False
End Function
Public Function CopyBillingDoc(Ug As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, Area As Long, IvdDocType As Long, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset
Dim Ivd As CInventoryDoc

   RName = "CopyBillingDoc"
   CopyBillingDoc = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryBillingDoc(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.BILLING_DOC_ID = -1
   Ug.INVENTORY_DOC_ID = -1
   Ug.DOCUMENT_NO = "**" & Ug.DOCUMENT_NO
   Ug.COMMIT_FLAG = "N" 'จะต้องยังไม่ commit ก่อน
   
   Dim CstName As CDoItem
   For Each CstName In Ug.DoItems
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName
   
   Dim RT As CReceiptItem
   For Each RT In Ug.ReceiptItems
      RT.Flag = "A"
      RT.AddEditMode = SHOW_ADD
   Next RT
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call glbDaily.DO2InventoryDoc(Ug, Ivd, Area, IvdDocType)
   Call AddEditInventoryDoc(Ivd, IsOK, False, ErrorObj)
   Ug.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   Call AddEditBillingDoc(Ug, IsOK, False, ErrorObj)
      
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

   CopyBillingDoc = True
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
   CopyBillingDoc = False
End Function

Public Function CopyPartItem(Ug As CPartItem, IsOK As Boolean, AutoCommit As Boolean, Area As Long, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset
Dim Ivd As CPartItem

   RName = "CopyPartItem"
   CopyPartItem = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryPartItem(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.PART_NO = "**" & Ug.PART_NO
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditPartItem(Ug, IsOK, False, ErrorObj)
      
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

   CopyPartItem = True
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
   CopyPartItem = False
End Function

Public Function CopyJob(Ug As CJob, IsOK As Boolean, AutoCommit As Boolean, Area As Long, ErrorObj As clsErrorLog, Optional ProcessID As Long)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset
Dim Ivd As CInventoryDoc
Dim IvWhd As CInventoryWHDoc
Dim TempIvdID As Long
Dim TempIvdWhID As Long

   RName = "CopyJob"
   CopyJob = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   If ProcessID = 5 Or ProcessID = 4 Then
      Call glbProduction.QueryJob(Ug, TempRs, iCount, IsOK, ErrorObj) 'QueryJob2
   Else
   
   End If
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.JOB_ID = -1
   Ug.INVENTORY_DOC_ID = -1
   Ug.INVENTORY_WH_DOC_ID = -1
   Ug.JOB_NO = "*" & Ug.JOB_NO
   Ug.COMMIT_FLAG = "N" 'จะต้องยังไม่ commit ก่อน
   
   Dim CstName As CJobInput
   For Each CstName In Ug.Inputs
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName

   For Each CstName In Ug.Outputs
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName

   Dim Jp As CJobParameter
   For Each Jp In Ug.Parameters
      Jp.Flag = "A"
      Jp.AddEditMode = SHOW_ADD
   Next Jp
   
   Dim Jv As CJobVerify
   For Each Jv In Ug.Verifies
      Jv.Flag = "A"
      Jv.AddEditMode = SHOW_ADD
   Next Jv
   
   Dim Jr As CJobResource
   For Each Jr In Ug.Peoples
      Jr.Flag = "A"
      Jr.AddEditMode = SHOW_ADD
   Next Jr
   
   For Each Jr In Ug.Machines
      Jr.Flag = "A"
      Jr.AddEditMode = SHOW_ADD
   Next Jr
   
   Dim Iw As CInventoryWHDoc
   If Not (Ug.InventoryWhDoc Is Nothing) Then
   For Each Iw In Ug.InventoryWhDoc
      Iw.Flag = "A"
      Iw.AddEditMode = SHOW_ADD
      Dim Lw As CLotItemWH
      For Each Lw In Iw.C_LotItemsWH
         Lw.Flag = "A"
         Lw.AddEditMode = SHOW_ADD
         Dim LD As CLotDoc
         For Each LD In Lw.C_LotDoc
            LD.Flag = "A"
            LD.AddEditMode = SHOW_ADD
            Dim PD As CPalletDoc
            For Each PD In LD.C_PalletDoc
               PD.Flag = "A"
               PD.AddEditMode = SHOW_ADD
            Next PD
         Next LD
      Next Lw
   Next Iw
   End If
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   If Area = 1 Then
      Call glbDaily.Job2InventoryDoc(Ug, Ivd, Area, 11)
      Call AddEditInventoryDoc(Ivd, IsOK, False, ErrorObj)
      TempIvdID = Ivd.INVENTORY_DOC_ID
      
       If Not (ProcessID = 5 Or ProcessID = 4) Then
         Call glbDaily.Job2InventoryWhDoc(Ug, IvWhd, Area, 11)
         Call AddEditInventoryWhDoc(IvWhd, IsOK, False, ErrorObj)
         TempIvdWhID = IvWhd.INVENTORY_WH_DOC_ID
      End If
   Else
      TempIvdID = -1
      TempIvdWhID = -1
   End If

   Ug.INVENTORY_DOC_ID = TempIvdID
   If Not (ProcessID = 5 Or ProcessID = 4) Then
    Ug.INVENTORY_WH_DOC_ID = TempIvdWhID
   End If
   Call glbProduction.AddEditJob(Ug, IsOK, False, ErrorObj)

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

   CopyJob = True
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
   CopyJob = False
End Function
Public Function CopyJob2(Ug As CJob, IsOK As Boolean, AutoCommit As Boolean, Area As Long, ErrorObj As clsErrorLog, Optional ProcessID As Long)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset
Dim Ivd As CInventoryDoc
Dim IvWhd As CInventoryWHDoc
Dim TempIvdID As Long
Dim TempIvdWhID As Long

   RName = "CopyJob2"
   CopyJob2 = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   If ProcessID = 5 Or ProcessID = 4 Then
      Call glbProduction.QueryJob(Ug, TempRs, iCount, IsOK, ErrorObj) 'QueryJob2
   Else
   
   End If
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.JOB_ID_REF = Ug.JOB_ID
   Ug.JOB_ID = -1
   Ug.INVENTORY_DOC_ID = -1
   Ug.INVENTORY_WH_DOC_ID = -1
   Ug.JOB_NO = Ug.JOB_NO & "-Sp"
   Ug.COMMIT_FLAG = "N" 'จะต้องยังไม่ commit ก่อน
   
   Dim CstName As CJobInput
   For Each CstName In Ug.Inputs
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName

   For Each CstName In Ug.Outputs
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName

   Dim Jp As CJobParameter
   For Each Jp In Ug.Parameters
      Jp.Flag = "A"
      Jp.AddEditMode = SHOW_ADD
   Next Jp
   
   Dim Jv As CJobVerify
   For Each Jv In Ug.Verifies
      Jv.Flag = "A"
      Jv.AddEditMode = SHOW_ADD
   Next Jv
   
   Dim Jr As CJobResource
   For Each Jr In Ug.Peoples
      Jr.Flag = "A"
      Jr.AddEditMode = SHOW_ADD
   Next Jr
   
   For Each Jr In Ug.Machines
      Jr.Flag = "A"
      Jr.AddEditMode = SHOW_ADD
   Next Jr
   
   Dim Iw As CInventoryWHDoc
   If Not (Ug.InventoryWhDoc Is Nothing) Then
   For Each Iw In Ug.InventoryWhDoc
      Iw.Flag = "A"
      Iw.AddEditMode = SHOW_ADD
      Dim Lw As CLotItemWH
      For Each Lw In Iw.C_LotItemsWH
         Lw.Flag = "A"
         Lw.AddEditMode = SHOW_ADD
         Dim LD As CLotDoc
         For Each LD In Lw.C_LotDoc
            LD.Flag = "A"
            LD.AddEditMode = SHOW_ADD
            Dim PD As CPalletDoc
            For Each PD In LD.C_PalletDoc
               PD.Flag = "A"
               PD.AddEditMode = SHOW_ADD
            Next PD
         Next LD
      Next Lw
   Next Iw
   End If
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   If Area = 1 Then
      Call glbDaily.Job2InventoryDoc(Ug, Ivd, Area, 11)
      Call AddEditInventoryDoc(Ivd, IsOK, False, ErrorObj)
      TempIvdID = Ivd.INVENTORY_DOC_ID
      
       If Not (ProcessID = 5 Or ProcessID = 4) Then
         Call glbDaily.Job2InventoryWhDoc(Ug, IvWhd, Area, 11)
         Call AddEditInventoryWhDoc(IvWhd, IsOK, False, ErrorObj)
         TempIvdWhID = IvWhd.INVENTORY_WH_DOC_ID
      End If
   Else
      TempIvdID = -1
      TempIvdWhID = -1
   End If

   Ug.INVENTORY_DOC_ID = TempIvdID
   If Not (ProcessID = 5 Or ProcessID = 4) Then
    Ug.INVENTORY_WH_DOC_ID = TempIvdWhID
   End If
   Call glbProduction.AddEditJob(Ug, IsOK, False, ErrorObj)

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

   CopyJob2 = True
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
   CopyJob2 = False
End Function

Public Function CopyIWDInputToIWD(Ug As CJob, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs1 As ADODB.Recordset

Set TempRs1 = New ADODB.Recordset

   RName = "CopyIWDInputToIWD"
   CopyIWDInputToIWD = False
   HasBegin = False

   Dim IWD As CInventoryWHDoc
   Dim LIW As CLotItemWH
   Dim LTD As CLotDoc
   Dim PD As CPalletDoc
   
   Dim Tiwd As CInventoryWHDoc
   Dim tLIW As CLotItemWH
   Dim Tltd As CLotDoc
   Dim tPD As CPalletDoc
   
   If Not (Ug.InventoryWhDocInput Is Nothing) Then
   Set Ug.tempIWDInput = New Collection
   For Each IWD In Ug.InventoryWhDocInput
      Set Tiwd = New CInventoryWHDoc
      
      Tiwd.DOCUMENT_TYPE = IWD.DOCUMENT_TYPE
      Tiwd.FROM_DATE = IWD.FROM_DATE
      Tiwd.FINISH_DATE = IWD.FINISH_DATE
      
      For Each LIW In IWD.C_LotItemsWH
         Set tLIW = New CLotItemWH
          tLIW.LOT_ITEM_WH_ID = LIW.LOT_ITEM_WH_ID
          tLIW.BIN_NO = LIW.BIN_NO
          tLIW.HEAD_PACK_NO = LIW.HEAD_PACK_NO
          tLIW.LOT_ID = LIW.LOT_ID
          tLIW.LOCK_NO = LIW.LOCK_NO
          tLIW.PACK_AMOUNT = LIW.PACK_AMOUNT
          tLIW.PART_ITEM_ID = LIW.PART_ITEM_ID
          tLIW.PART_NO = LIW.PART_NO
          tLIW.PART_DESC = LIW.PART_DESC
          tLIW.LOCATION_ID = LIW.LOCATION_ID
          tLIW.PRODUCT_TYPE_ID = LIW.PRODUCT_TYPE_ID
          tLIW.START_DATE = LIW.START_DATE
          tLIW.TX_AMOUNT = LIW.TX_AMOUNT
          tLIW.WEIGHT_AMOUNT = LIW.WEIGHT_AMOUNT
          tLIW.WEIGHT_PER_PACK = LIW.WEIGHT_PER_PACK
      
            For Each LTD In LIW.C_LotDoc
              Set Tltd = New CLotDoc
              Tltd.LOT_ID = LTD.LOT_ID
              Tltd.LOCK_NO = LTD.LOCK_NO
              
               For Each PD In LTD.C_PalletDoc
                  Set tPD = New CPalletDoc
                  tPD.CAPACITY_AMOUNT = PD.CAPACITY_AMOUNT
                  tPD.PALLET_DOC_NO = PD.PALLET_DOC_NO
                  Call Tltd.C_PalletDoc.add(tPD)
                  Set tPD = Nothing
               Next PD
               Call tLIW.C_LotDoc.add(Tltd)
               Set Tltd = Nothing
            Next LTD
         Call Tiwd.C_LotItemsWH.add(tLIW)
         Set tLIW = Nothing
      Next LIW
     Call Ug.tempIWDInput.add(Tiwd)
     Set Tiwd = Nothing
   Next IWD
   End If

   HasBegin = True
   CopyIWDInputToIWD = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyIWDInputToIWD = False
End Function
Public Function CopyInventoryDoc(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, Area As Long, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset
Dim TempIvdID As Long

   RName = "CopyInventoryDoc"
   CopyInventoryDoc = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call glbDaily.QueryInventoryDoc(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.INVENTORY_DOC_ID = -1
   Ug.DOCUMENT_NO = "*" & Ug.DOCUMENT_NO
   Ug.COMMIT_FLAG = "N" 'จะต้องยังไม่ commit ก่อน
   
   Dim Li As CLotItem
   For Each Li In Ug.ImportExports
      Li.Flag = "A"
      Li.AddEditMode = SHOW_ADD
   Next Li

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   
   Call glbDaily.AddEditInventoryDoc(Ug, IsOK, False, ErrorObj)

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

   CopyInventoryDoc = True
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
   CopyInventoryDoc = False
End Function

Public Function ImportLegacyUnit(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CUnit
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyUnit"
   ImportLegacyUnit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyUnit(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(1, m_Rs1)
      
      Set U = New CUnit
      U.AddEditMode = SHOW_ADD
      U.UNIT_NO = Format(I, "000")
      U.UNIT_NAME = D.FIELD1
      If CheckUniqueNs(UNIT_NO, U.UNIT_NO, -1) Then
         U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyUnit = True
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
   ImportLegacyUnit = False
End Function

Public Function ImportLegacyPartType(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CPartType
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPartType"
   ImportLegacyPartType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPartType(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(2, m_Rs1)
      
      Set U = New CPartType
      U.AddEditMode = SHOW_ADD
      U.PART_TYPE_NO = D.FIELD1
      U.PART_TYPE_NAME = D.FIELD2
      U.RAW_FLAG = "Y"
      If CheckUniqueNs(PARTTYPE_NO, U.PART_TYPE_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPartType = True
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
   ImportLegacyPartType = False
End Function

Public Function ImportLegacyLocation(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CLocation
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyLocation"
   ImportLegacyLocation = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyLocation(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(3, m_Rs1)
      
      Set U = New CLocation
      U.AddEditMode = SHOW_ADD
      U.LOCATION_NO = "I-" & D.FIELD1
      U.LOCATION_NAME = D.FIELD2
      U.SALE_FLAG = "N"
      U.LOCATION_TYPE = 2
      If CheckUniqueNs(LOCATION_NO, U.LOCATION_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyLocation = True
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
   ImportLegacyLocation = False
End Function

Public Function ImportLegacyHouse(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CLocation
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyHouse"
   ImportLegacyHouse = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyHouse(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(3, m_Rs1)
      
      Set U = New CLocation
      U.AddEditMode = SHOW_ADD
      U.LOCATION_NO = "H-" & D.FIELD1
      U.LOCATION_NAME = D.FIELD2
      If D.FIELD1 = "99" Then
         U.SALE_FLAG = "Y"
   Else
         U.SALE_FLAG = "N"
      End If
      U.LOCATION_TYPE = 1
      If CheckUniqueNs(LOCATION_NO, U.LOCATION_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyHouse = True
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
   ImportLegacyHouse = False
End Function

Public Function ImportLegacyPigType(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CProductType
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPigType"
   ImportLegacyPigType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPigType(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(4, m_Rs1)
      
      Set U = New CProductType
      U.AddEditMode = SHOW_ADD
      U.PRODUCT_TYPE_NO = D.FIELD1
      U.PRODUCT_TYPE_NAME = D.FIELD2
      If CheckUniqueNs(PRODUCTTYPE_NO, U.PRODUCT_TYPE_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPigType = True
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
   ImportLegacyPigType = False
End Function

Public Function ImportLegacyPigStatus(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CProductStatus
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPigStatus"
   ImportLegacyPigStatus = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPigStatus(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(5, m_Rs1)
      
      Set U = New CProductStatus
      U.AddEditMode = SHOW_ADD
      U.PRODUCT_STATUS_NO = D.FIELD1
      U.PRODUCT_STATUS_NAME = D.FIELD2
      If CheckUniqueNs(PRODUCTSTATUS_NO, U.PRODUCT_STATUS_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPigStatus = True
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
   ImportLegacyPigStatus = False
End Function

Private Function LookupPartTypeID(PartTypeNO As String) As Long
Static TempCol As Collection
Dim T As CPartType
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartType(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PART_TYPE_ID
      End If
      
      If T.PART_TYPE_NO = PartTypeNO Then
            TempID = T.PART_TYPE_ID
         Exit For
      End If
   Next T
   
   LookupPartTypeID = TempID
End Function

Private Function LookupPartItemID(PartNo As String) As Long
Static TempCol As Collection
Dim T As CPartItem
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PART_ITEM_ID
      End If
      
      If (Trim(T.PART_NO) = PartNo) Then
            TempID = T.PART_ITEM_ID
         Exit For
      End If
   Next T
   
   LookupPartItemID = TempID
End Function

Private Function LookupPigID(PigNo As String, PigType As String) As Long
Static TempCol As Collection
Dim T As CPartItem
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol, , "Y")
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PART_ITEM_ID
      End If
      
      If (Trim(T.PART_NO) = PigNo) And (T.PIG_TYPE = PigType) Then
            TempID = T.PART_ITEM_ID
         Exit For
      End If
   Next T
   
   LookupPigID = TempID
End Function

Private Function LookupPigStatus(PigStatusNo As String) As Long
Static TempCol As Collection
Dim T As CProductStatus
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadProductStatus(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PRODUCT_STATUS_ID
      End If
      
      If (Trim(T.PRODUCT_STATUS_NO) = PigStatusNo) Then
            TempID = T.PRODUCT_STATUS_ID
         Exit For
      End If
   Next T
   
   LookupPigStatus = TempID
End Function

Private Function LookupLocationID(LocationNo As String, Optional SaleFlag As String = "") As Long
Static TempCol As Collection
Dim T As CLocation
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadLocation(Nothing, TempCol, , SaleFlag)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.LOCATION_ID
      End If
      
      If T.LOCATION_NO = LocationNo Then
            TempID = T.LOCATION_ID
         Exit For
      End If
   Next T
   
   LookupLocationID = TempID
End Function

Private Function LookupSaleLocationID(LocationNo As String) As Long
Static TempCol As Collection
Dim T As CLocation
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadLocation(Nothing, TempCol, , "Y")
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.LOCATION_ID
      End If
      
      If T.LOCATION_NO = LocationNo Then
            TempID = T.LOCATION_ID
         Exit For
      End If
   Next T
   
   LookupSaleLocationID = TempID
End Function

Private Function LookupSupplierID(SupplierNo As String) As Long
Static TempCol As Collection
Dim T As CSupplier
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadSupplier(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.SUPPLIER_ID
      End If
      
      If T.SUPPLIER_CODE = SupplierNo Then
            TempID = T.SUPPLIER_ID
         Exit For
      End If
   Next T
   
   LookupSupplierID = TempID
End Function

Private Function LookupEmpID(EmpCode As String) As Long
Static TempCol As Collection
Dim T As CEmployee
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadEmployee(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.EMP_ID
      End If
      
      If T.EMP_CODE = EmpCode Then
            TempID = T.EMP_ID
         Exit For
      End If
   Next T
   
   LookupEmpID = TempID
End Function

Private Function LookupPartUnitID(UnitName As String) As Long
Static TempCol As Collection
Dim T As CUnit
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadUnit(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.UNIT_ID
      End If
      
      If T.UNIT_NAME = UnitName Then
            TempID = T.UNIT_ID
         Exit For
      End If
   Next T
   
   LookupPartUnitID = TempID
End Function

Private Function LookupGroupID(GroupNo As String) As Long
Static TempCol As Collection
Dim T As CUserGroup
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadUserGroup(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.GROUP_ID
      End If
      
      If T.GROUP_NAME = GroupNo Then
            TempID = T.GROUP_ID
         Exit For
      End If
   Next T
   
   LookupGroupID = TempID
End Function

Public Function ImportLegacyPartItem(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CPartItem
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPartItem"
   ImportLegacyPartItem = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPartItem(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(6, m_Rs1)
      
      Set U = New CPartItem
      U.AddEditMode = SHOW_ADD
      U.PART_NO = D.FIELD2 & "-" & D.FIELD1
      U.PART_DESC = D.FIELD3
      U.PIG_FLAG = "N"
      U.PART_TYPE = LookupPartTypeID(D.FIELD1)
      If CheckUniqueNs(PARTNO_UNIQUE, U.PART_NO, -1) Then
         U.UNIT_COUNT = LookupPartUnitID(D.FIELD4)
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPartItem = True
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
   ImportLegacyPartItem = False
End Function

Public Function ImportLegacySupplier(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CSupplier
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacySupplier"
   ImportLegacySupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacySupplier(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(7, m_Rs1)
      
      Set U = New CSupplier
      U.AddEditMode = SHOW_ADD
      U.SUPPLIER_CODE = D.FIELD1
      U.SUPPLIER_NAME = D.FIELD2
      U.SUPPLIER_GRADE = -1
      U.SUPPLIER_TYPE = -1
      U.SUPPLIER_STATUS = -1
      
      Dim CstName As CSupplierName
      Dim NAME As CName
      If U.CstNames.Count <= 0 Then
         Set CstName = New CSupplierName
         CstName.Flag = "A"
         
         Set NAME = CstName.NAME
         NAME.LONG_NAME = D.FIELD2
         NAME.SHORT_NAME = ""
         NAME.Flag = "A"
         
         Call U.CstNames.add(CstName)
      End If
      
      If CheckUniqueNs(SUPPLIER_UNIQUE, U.SUPPLIER_CODE, -1) Then
         Call glbDaily.AddEditSupplier(U, IsOK, False, ErrorObj)
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacySupplier = True
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
   ImportLegacySupplier = False
End Function

Public Function ImportLegacyEmployee(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CEmployee
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyEmployee"
   ImportLegacyEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyEmployee(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(8, m_Rs1)
      
      Set U = New CEmployee
      U.AddEditMode = SHOW_ADD
      U.EMP_CODE = D.FIELD1
      U.NAME = D.FIELD2
      U.LASTNAME = ""
      U.CURRENT_POSITION = -1
      U.PASS_STATUS = "Y"
      U.EXTERNAL_FLAG = "N"
      
      U.EmpName.AddEditMode = SHOW_ADD
      U.EName.AddEditMode = SHOW_ADD
      U.EName.LONG_NAME = D.FIELD2
      U.EName.LAST_NAME = ""
         
      If CheckUniqueNs(EMPCODE_UNIQUE, U.EMP_CODE, -1) Then
         Call glbDaily.AddEditEmployee(U, IsOK, False, ErrorObj)
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyEmployee = True
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
   ImportLegacyEmployee = False
End Function

Public Function ImportLegacyUser(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CUserAccount
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyUser"
   ImportLegacyUser = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyUser(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(9, m_Rs1)
      
      Set U = New CUserAccount
      U.AddEditMode = SHOW_ADD
      U.USER_NAME = D.FIELD1
      U.USER_PASSWORD = D.FIELD2
      U.USER_STATUS2 = "Y"
      U.EXCEPTION_FLAG = "Y"
      U.GROUP_ID = LookupGroupID("")
      If CheckUniqueNs(USERNAME_UNIQUE, U.USER_NAME, -1) Then
         U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyUser = True
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
   ImportLegacyUser = False
End Function

Public Function ImportLegacyCustomer(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CCustomer
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyCustomer"
   ImportLegacyCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyCustomer(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(10, m_Rs1)
      
      Set U = New CCustomer
      U.AddEditMode = SHOW_ADD
      U.CUSTOMER_CODE = D.FIELD1
      U.CUSTOMER_TYPE = -1
      U.CUSTOMER_NAME = D.FIELD4
      
      Dim CstName As CCustomerName
      Dim NAME As CName
      If U.CstNames.Count <= 0 Then
         Set CstName = New CCustomerName
         CstName.Flag = "A"
         
         Set NAME = CstName.NAME
         NAME.LONG_NAME = D.FIELD4
         NAME.SHORT_NAME = ""
         NAME.Flag = "A"
         
         Call U.CstNames.add(CstName)
      End If
      
      If U.CstAccounts.Count <= 0 Then
         Dim Acc As CAccount
         
         Set Acc = New CAccount
         Acc.ACCOUNT_NO = U.CUSTOMER_CODE
         Acc.Flag = "A"
         
         Call U.CstAccounts.add(Acc)
         Set Acc = Nothing
      End If
   
      If CheckUniqueNs(CUSTCODE_UNIQUE, U.CUSTOMER_CODE, -1) Then
         Call glbDaily.AddEditCustomer(U, IsOK, False, ErrorObj)
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyCustomer = True
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
   ImportLegacyCustomer = False
End Function

Public Function ClearInterim(AutoCommit As Boolean)
On Error GoTo ErrorHandler
Dim RName As String
Dim H As CLegacy_h
Dim HasBegin As Boolean

   RName = "ClearInterim"
   ClearInterim = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set H = New CLegacy_h
   Call H.DeleteAllData
   Set H = Nothing
   
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False

   ClearInterim = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = m_ModuleName
   glbErrorLog.ShowErrorLog (LOG_TO_FILE)

   ClearInterim = False
End Function

Public Function ImportToInterim(AutoCommit As Boolean)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim H As CLegacy_h
Dim D As CLegacy_d
Dim L1 As CLegacy
Dim L2 As CLegacy
Dim HasBegin As Boolean

   RName = "ImportToInterim"
   ImportToInterim = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set L1 = New CLegacy
   Set L2 = New CLegacy
   
   '===
   Call L1.QueryData11(m_Rs1, iCount) 'Feed stock header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(11, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''Debug.Print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 1
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData12(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(12, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
'   ===
   
'   ===
   Call L1.QueryData13(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(13, m_Rs1)
      L2.FIELD1 = L1.FIELD1
'''Debug.Print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 2
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD4
      Call H.AddEditData

      Call L2.QueryData14(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(14, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
'   ===
   
'   ===
   Call L1.QueryData15(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(15, m_Rs1)
      L2.FIELD1 = L1.FIELD1
'''Debug.Print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 3
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData16(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(16, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
'   ===
   
   Call L1.QueryData17(m_Rs1, iCount)
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(17, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''Debug.Print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 4
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD4
      Call H.AddEditData

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
   
'   ===
   Call L1.QueryData18(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(18, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''Debug.Print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 5
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData19(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(19, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend
      Set H = Nothing
      m_Rs1.MoveNext
   Wend

   Call L1.QueryData20(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(20, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''Debug.Print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 0
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData21(m_Rs2, iCount)
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(21, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
   Set H = New CLegacy_h
   Call H.SetDocType
   Set H = Nothing
   '===
   
   Set L2 = Nothing
   Set L1 = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False

   ImportToInterim = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = m_ModuleName
   glbErrorLog.ShowErrorLog (LOG_TO_FILE)

   ImportToInterim = False
End Function

Public Sub TriggerCommit(Col As Collection)
Dim D As Object

   For Each D In Col
      If (D.Flag <> "D") And (D.Flag <> "A") Then
         D.Flag = "E"
      End If
   Next D
End Sub

'===
Public Function QueryLegacyH(Ua As CLegacy_h, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim Exp As CXItem

   RName = "QueryLegacyH"
   QueryLegacyH = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CLegacy_d
      
      Set Yw = New CLegacy_d
      Yw.LEGACY_D_ID = -1
      Yw.LEGACY_H_ID = Ua.LEGACY_H_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.LegacyDetails = Nothing
      Set Ua.LegacyDetails = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CLegacy_d
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.LegacyDetails.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryLegacyH = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLegacyH = False
End Function


Public Sub CreateInventoryDoc2(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag) 'ใบนำเข้า
Dim Lhd As CLegacy_d
Dim II As CLotItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD4)
   Ivd.DELIVERY_FEE = Val(Lh.H_FIELD10)
   Ivd.DELIVERY_ID = LookupSupplierID(Lh.H_FIELD6)
   Ivd.DO_NO = Lh.H_FIELD2
   Ivd.DOCUMENT_TYPE = 1
   Ivd.EMP_ID = -1
   Ivd.SUPPLIER_ID = LookupSupplierID(Lh.H_FIELD3)
   Ivd.TRUCK_NO = Lh.H_FIELD5
   Ivd.BILL_NO = Lh.H_FIELD7
   Ivd.COMMIT_FLAG = CommitFlag
   
   Sum = 0
   For Each Lhd In Lh.LegacyDetails
      Sum = Sum + Val(Lhd.D_FIELD3)
   Next Lhd
   
   For Each Lhd In Lh.LegacyDetails
      Set II = New CLotItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "Y"
      
      II.ACTUAL_UNIT_PRICE = Val(Lhd.D_FIELD4)
      If Sum > 0 Then
         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + (Ivd.DELIVERY_FEE / Sum)
      Else
         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE
      End If
      II.TX_AMOUNT = Val(Lhd.D_FIELD3)
      II.LOCATION_ID = LookupLocationID("I-00")
      II.PART_ITEM_ID = LookupPartItemID(Lhd.D_FIELD2 & "-" & Lhd.D_FIELD7)
      
      Call Ivd.ImportExports.add(II)
      Set II = Nothing
   Next Lhd
End Sub


Public Sub CreateInventoryDoc4(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag As String)  'ใบสุกรคลอด
Dim Lhd As CLegacy_d
Dim II As CLotItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD4)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 5
   Ivd.EMP_ID = -1
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag
   
   Sum = 0
'   For Each Lhd In Lh.LegacyDetails
      Set II = New CLotItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "N"
      
      II.ACTUAL_UNIT_PRICE = 0
      II.INCLUDE_UNIT_PRICE = 0
      II.TX_AMOUNT = Val(Lh.H_FIELD7)
      II.LOCATION_ID = LookupLocationID("H-" & Lh.H_FIELD3)
      II.PART_ITEM_ID = LookupPigID(Lh.H_FIELD10, "M")
      II.FATHER_NO = Lh.H_FIELD6
      II.MOTHER_NO = Lh.H_FIELD5
      II.TOTAL_WEIGHT = Val(Lh.H_FIELD9)
      Call Ivd.ImportExports.add(II)
      Set II = Nothing
'   Next Lhd
End Sub

Public Function VerifySubLot(Li As CLotItem) As Boolean
Dim R As CSubLotItem
Dim Sum As Double

   Sum = 0
   For Each R In Li.SubLotItems
      If R.Flag <> "D" Then
         Sum = Sum + R.SUBLOT_AMT
      End If
   Next R

   If Sum = Li.TX_AMOUNT Then
      VerifySubLot = True
   Else
      VerifySubLot = False
      glbErrorLog.LocalErrorMsg = "จำนวนรวมในแต่ละล็อตจะต้อง = " & Li.TX_AMOUNT
      glbErrorLog.ShowUserError
   End If
End Function
Public Function VerifyDOSubLot(Di As CDoItem) As Boolean
Dim R As CSubLotItem
Dim Sum As Double

   Sum = 0
   For Each R In Di.SubLotItems
      If R.Flag <> "D" Then
         Sum = Sum + R.SUBLOT_AMT
      End If
   Next R

   If Sum = Di.ITEM_AMOUNT Then
      VerifyDOSubLot = True
   Else
      VerifyDOSubLot = False
      glbErrorLog.LocalErrorMsg = "จำนวนรวมในแต่ละล็อตจะต้อง = " & Di.ITEM_AMOUNT
      glbErrorLog.ShowUserError
   End If
End Function

Public Sub GenerateSubLot(Li As CLotItem, SubLotItems As Collection)
Dim TempLi As CLotItem
Dim iCount As Long
Dim TempCols As Collection
Dim D As CLotItem
Dim NeedAmt As Double
Dim C As CSubLotItem

   Set TempCols = New Collection

   Set TempLi = New CLotItem
   TempLi.LOT_ITEM_ID = -1
   TempLi.COMMIT_FLAG = "Y"
   TempLi.PART_ITEM_ID = Li.PART_ITEM_ID
   TempLi.LOCATION_ID = Li.LOCATION_ID
   TempLi.TX_TYPE = "I"
   TempLi.OrderType = 1
   Call LoadSubLotItem(Nothing, TempCols, TempLi)

   If Li.LOT_MANUAL = "Y" Then
      Set Li.SubLotItems = Nothing
      Set Li.SubLotItems = New Collection
      
      For Each C In SubLotItems
         Call Li.SubLotItems.add(C)
      Next C
   Else ' Generate automatically
      For Each C In Li.SubLotItems
         C.Flag = "D"
      Next C
            
      NeedAmt = Li.TX_AMOUNT
      For Each D In TempCols
         If D.LEFT_AMOUNT >= NeedAmt Then
            D.LEFT_AMOUNT = D.LEFT_AMOUNT - NeedAmt
            
            Set C = New CSubLotItem
            C.Flag = "A"
            C.LOT_ITEM_ID = Li.LOT_ITEM_ID
            C.FROM_LOT_ID = D.LOT_ITEM_ID
            C.SUBLOT_AMT = NeedAmt
            C.DOCUMENT_DATE = D.DOCUMENT_DATE
            C.DOCUMENT_NO = D.DOCUMENT_NO
            Call Li.SubLotItems.add(C)
            Set C = Nothing
            
            NeedAmt = 0
            Exit For
         Else
            Set C = New CSubLotItem
            C.Flag = "A"
            C.LOT_ITEM_ID = Li.LOT_ITEM_ID
            C.FROM_LOT_ID = D.LOT_ITEM_ID
            C.SUBLOT_AMT = D.LEFT_AMOUNT
            C.DOCUMENT_DATE = D.DOCUMENT_DATE
            C.DOCUMENT_NO = D.DOCUMENT_NO
            Call Li.SubLotItems.add(C)
            Set C = Nothing
            
            NeedAmt = NeedAmt - D.LEFT_AMOUNT
            D.LEFT_AMOUNT = 0
         End If
      Next D
   End If
   
   Set TempLi = Nothing
   Set TempCols = Nothing
End Sub

Public Sub GenerateDOSubLot(Di As CDoItem, SubLotItems As Collection)
Dim TempLi As CLotItem
Dim iCount As Long
Dim TempCols As Collection
Dim D As CLotItem
Dim NeedAmt As Double
Dim C As CSubLotItem

   Set TempCols = New Collection

   Set TempLi = New CLotItem
   TempLi.LOT_ITEM_ID = -1
   TempLi.COMMIT_FLAG = "Y"
   TempLi.PART_ITEM_ID = Di.PART_ITEM_ID
   TempLi.LOCATION_ID = Di.LOCATION_ID
   TempLi.TX_TYPE = "I"
   TempLi.OrderType = 1
   Call LoadSubLotItem(Nothing, TempCols, TempLi)
Di.LOT_MANUAL = "N"

   If Di.LOT_MANUAL = "Y" Then
      Set Di.SubLotItems = Nothing
      Set Di.SubLotItems = New Collection
      
      For Each C In SubLotItems
         Call Di.SubLotItems.add(C)
      Next C
   Else ' Generate automatically
      For Each C In Di.SubLotItems
         C.Flag = "D"
      Next C
      
      NeedAmt = Di.ITEM_AMOUNT
      For Each D In TempCols
         If D.LEFT_AMOUNT >= NeedAmt Then
            D.LEFT_AMOUNT = D.LEFT_AMOUNT - NeedAmt
            
            Set C = New CSubLotItem
            C.Flag = "A"
            C.LOT_ITEM_ID = Di.LOT_ITEM_ID
            C.FROM_LOT_ID = D.LOT_ITEM_ID
            C.SUBLOT_AMT = NeedAmt
            C.DOCUMENT_DATE = D.DOCUMENT_DATE
            C.DOCUMENT_NO = D.DOCUMENT_NO
            Call Di.SubLotItems.add(C)
            Set C = Nothing
            
            NeedAmt = 0
            Exit For
         Else
            Set C = New CSubLotItem
            C.Flag = "A"
            C.LOT_ITEM_ID = Di.LOT_ITEM_ID
            C.FROM_LOT_ID = D.LOT_ITEM_ID
            C.SUBLOT_AMT = D.LEFT_AMOUNT
            C.DOCUMENT_DATE = D.DOCUMENT_DATE
            C.DOCUMENT_NO = D.DOCUMENT_NO
            Call Di.SubLotItems.add(C)
            Set C = Nothing
            
            NeedAmt = NeedAmt - D.LEFT_AMOUNT
            D.LEFT_AMOUNT = 0
         End If
      Next D
   End If
   
   Set TempLi = Nothing
   Set TempCols = Nothing
End Sub

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
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
   AddEditReportConfig = False
End Function

Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function

Public Function PatchNewPrice(Ug As CLotItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim II As CLotItem
Dim NewPrice As Double

   RName = "PatchNewPrice"
   PatchNewPrice = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Ug.OrderBy = 10
   Ug.OrderType = 1
   Call Ug.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Set II = New CLotItem
      Call II.PopulateFromRS(1, m_Rs1)
      
      NewPrice = (((II.NEW_AMOUNT - II.TX_AMOUNT) * II.PREVIOUS_PRICE) + II.TOTAL_INCLUDE_PRICE) / II.NEW_AMOUNT
      II.NEW_PRICE = NewPrice
      
      II.AddEditMode = SHOW_EDIT
      Call II.AddEditData
      
      Set II = Nothing
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   PatchNewPrice = True
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
   PatchNewPrice = False
End Function

Public Function PatchExtraDiscount(BD As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempBd As CBillingDoc
Dim TempRs1 As ADODB.Recordset
Dim TempRs2  As ADODB.Recordset

   RName = "PatchExtraDiscount"
   PatchExtraDiscount = False
   HasBegin = False

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset

   BD.COMMIT_FLAG = "N"
   Call QueryBillingDoc(BD, TempRs1, iCount, IsOK, glbErrorLog)
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   While Not TempRs1.EOF
      Call BD.PopulateFromRS(1, TempRs1)
            
      Set TempBd = New CBillingDoc
      TempBd.BILLING_DOC_ID = BD.BILLING_DOC_ID
      TempBd.QueryFlag = 1
      Call QueryBillingDoc(TempBd, TempRs2, iCount, IsOK, glbErrorLog)
      If Not TempRs2.EOF Then
         Call TempBd.PopulateFromRS(1, TempRs2)
         Call GenerateExtraDiscount(TempBd)
         
         TempBd.AddEditMode = SHOW_EDIT
         Call AddEditBillingDoc(TempBd, IsOK, AutoCommit, glbErrorLog)
      End If
      Set TempBd = Nothing
      
      TempRs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   If Not TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   If Not TempRs2.State = adStateOpen Then
      TempRs2.Close
   End If
   
   PatchExtraDiscount = True
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
   PatchExtraDiscount = False
End Function

Public Function PatchBankAccount(IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim Pm As CPayment
Dim Mr As CMasterRef
Dim TempRs1 As ADODB.Recordset
Dim I As Long

   RName = "PatchBankAccount"
   PatchBankAccount = False
   HasBegin = False

   Set TempRs1 = New ADODB.Recordset

   Set Pm = New CPayment
   Pm.PAYMENT_ID = -1
   Call Pm.QueryData(2, TempRs1, iCount)
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   I = 0
   While Not TempRs1.EOF
      I = I + 1
      Call Pm.PopulateFromRS(2, TempRs1)
      
      Set Mr = New CMasterRef
      Mr.AddEditMode = SHOW_ADD
      Mr.KEY_CODE = Format(I, "00")
      Mr.KEY_NAME = Pm.ACCOUNT_NO
      Mr.TEMP_ID1 = Pm.TO_BANK_ID
      Mr.TEMP_ID2 = Pm.TO_BANK_BRANCH
      Mr.MASTER_AREA = BANK_ACCOUNT
      Call Mr.AddEditData
      Set Mr = Nothing
      
      TempRs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   If Not TempRs1.State = adStateOpen Then
      TempRs1.Close
   End If
   
   PatchBankAccount = True
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
   PatchBankAccount = False
End Function

Public Sub CreateTransferItems(Ivd As CInventoryDoc)
Dim Ti As CTransferItem
Dim O As Object
Dim Ei As CLotItem
Dim II As CLotItem
Dim I As Long
Dim J As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.ImportExports.Count \ 2
   Count2 = Ivd.ImportExports.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      J = 1
      While J <= Count2
         Set O = Ivd.ImportExports(J)
         If (O.TX_TYPE = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.TX_TYPE = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set Ei = O
            End If
         End If
         J = J + 1
      Wend
         
      Set Ti = New CTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = Ei
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.add(Ti)
      Set Ti = Nothing
   Next I
End Sub

Public Sub MergeImportExport(Imps As Collection, Exps As Collection, ImpExps As Collection)
Dim Count As Long
Dim IP As CLotItem
Dim Exp As CLotItem
Dim O As Object
   
   For Each IP In Imps
      Call ImpExps.add(IP)
   Next IP

   For Each Exp In Exps
      Call ImpExps.add(Exp)
   Next Exp
   
   Call Selectionsort(ImpExps, 1, ImpExps.Count)
End Sub


Private Function GetNextTx(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, Rs3 As ADODB.Recordset, II As CLotItem, Ei As CLotItem, Ri As CLotItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
Dim EofFlag3 As Boolean
Dim TempCol As Collection
Dim O As Object
Dim I As Long
Dim MinSeq As Long
Dim MinIdx As Long
   
   Set TempCol = New Collection
   
   'Import
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call II.PopulateFromRS(1, Rs1)
      Call TempCol.add(II)
   End If

   'Export
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call Ei.PopulateFromRS(1, Rs2)
      Call TempCol.add(Ei)
   End If
      
   'Ro item
'   EofFlag3 = Rs3.EOF
'   If Not Rs3.EOF Then
'      Call Ri.PopulateFromRS(1, Rs3)
'      Call TempCol.Add(Ri)
'   End If
      
   I = 0
   MinIdx = 0
   For Each O In TempCol
      I = I + 1
      If I = 1 Then
         MinSeq = O.TRANSACTION_SEQ
         MinIdx = I
      End If
      
      If O.TRANSACTION_SEQ < MinSeq Then
         MinSeq = O.TRANSACTION_SEQ
         MinIdx = I
      End If
   Next O
   
   If MinIdx > 0 Then
      Set O = TempCol(MinIdx)
      GetNextTx = O.TX_TYPE
   Else
      GetNextTx = ""
   End If
   
   Set TempCol = Nothing
End Function

Public Function QueryAllTransaction(Ivd As CInventoryDoc, IsOK As Boolean, RecordCount As Long, ErrorObj As clsErrorLog, Optional ClearFlag As Boolean = False) As Object
Static Rs1 As ADODB.Recordset
Static Rs2 As ADODB.Recordset
Static Rs3 As ADODB.Recordset
Static II As CLotItem
Static Ei As CLotItem
'Static Ri As CROItem
Dim iCount1 As Long
Dim iCount2 As Long
Dim iCount3 As Long
Static TxCode As String

   If ClearFlag Then
      Set II = Nothing
      Set Ei = Nothing
'      Set Ri = Nothing
   End If
   
   If II Is Nothing Then
      Set II = New CLotItem
      Set Rs1 = New ADODB.Recordset
      
      II.FROM_DATE = Ivd.FROM_DATE
      II.TO_DATE = Ivd.TO_DATE
      II.COMMIT_FLAG = Ivd.COMMIT_FLAG
      II.TX_TYPE = "I"
      II.OrderBy = 11
      II.OrderType = 1
      Call II.QueryData(1, Rs1, iCount1)
   End If
   
   If Ei Is Nothing Then
      Set Ei = New CLotItem
      Set Rs2 = New ADODB.Recordset
      
      Ei.FROM_DATE = Ivd.FROM_DATE
      Ei.TO_DATE = Ivd.TO_DATE
      Ei.COMMIT_FLAG = Ivd.COMMIT_FLAG
      Ei.TX_TYPE = "E"
      Ei.OrderBy = 11
      Ei.OrderType = 1
      Call Ei.QueryData(1, Rs2, iCount2)
   End If
      
'   If Ri Is Nothing Then
'      Set Ri = New CROItem
'      Set Rs3 = New ADODB.Recordset
'
'      Ri.FROM_DATE = Ivd.FROM_DATE
'      Ri.TO_DATE = Ivd.TO_DATE
'      Ri.COMMIT_FLAG = Ivd.COMMIT_FLAG
'      Ri.OrderBy = 10
'      Ri.OrderType = 1
'      Call Ri.QueryData(1, Rs3, iCount3)
'   End If
   
   If ClearFlag Then
      RecordCount = iCount1 + iCount2 + iCount3
      Exit Function
   End If
   
   TxCode = GetNextTx(Rs1, Rs2, Rs3, II, Ei, Nothing)
   If TxCode <> "" Then
      If TxCode = "E" Then
         Set QueryAllTransaction = Ei
         Rs2.MoveNext
      ElseIf TxCode = "I" Then
         Set QueryAllTransaction = II
         Rs1.MoveNext
      ElseIf TxCode = "X" Then
'         Set QueryAllTransaction = Ri
         Rs3.MoveNext
      End If
   Else
      Set QueryAllTransaction = Nothing
   End If
   
   RecordCount = iCount1 + iCount2 + iCount3
End Function

Public Function QueryFeature(Ua As CFeature, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryFeature"
   QueryFeature = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim SocFea As CSocFeature
      Set SocFea = New CSocFeature
      SocFea.FEATURE_ID = Ua.FEATURE_ID
      SocFea.SOC_ID = -1
      SocFea.SOC_FEATURE_ID = -1
      Call SocFea.QueryData(1, m_Rs1, iCount)
      Set SocFea = Nothing

      Set Ua.SocFeatures = Nothing
      Set Ua.SocFeatures = New Collection
      While Not m_Rs1.EOF
         Set SocFea = New CSocFeature
         Call SocFea.PopulateFromRS(1, m_Rs1)

         SocFea.Flag = "I"

         Call Ua.SocFeatures.add(SocFea)
         Set SocFea = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryFeature = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryFeature = False
End Function

Public Function AddEditFeature(Ug As CFeature, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFeature"
   AddEditFeature = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ug.AddEditData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   IsOK = True

   AddEditFeature = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditFeature = False
End Function

Public Function DeleteFeature(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim Fea As CFeature

   RName = "DeleteFeature"
   DeleteFeature = False
   HasBegin = False

   Set Fea = New CFeature
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Fea.FEATURE_ID = UID
   Fea.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   IsOK = True

   DeleteFeature = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
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

   DeleteFeature = True
End Function

Public Function QuerySoc(Ua As CSoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QuerySoc"
   QuerySoc = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
         
   If Ua.QueryFlag = 1 Then
      Dim SocFea As CSocFeature
      Set SocFea = New CSocFeature
      SocFea.SOC_ID = Ua.SOC_ID
      SocFea.FEATURE_ID = -1
      SocFea.SOC_FEATURE_ID = -1
      SocFea.OrderBy = 1
      Call SocFea.QueryData(1, m_Rs1, iCount)
      Set SocFea = Nothing

      Set Ua.SocFeatures = Nothing
      Set Ua.SocFeatures = New Collection
      While Not m_Rs1.EOF
         Set SocFea = New CSocFeature
         Call SocFea.PopulateFromRS(1, m_Rs1)

         Dim Rc As CRcRate
         Set Rc = New CRcRate
         Rc.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         Call Rc.QueryData(m_Rs2, iCount)
         Call Rc.PopulateFromRS(m_Rs2)
         Rc.Flag = "I"

         Dim OC As COcRate
         Set OC = New COcRate
         OC.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         Call OC.QueryData(m_Rs2, iCount)
         Call OC.PopulateFromRS(m_Rs2)
         OC.Flag = "I"

         Dim Ac As CAcRate
         Set Ac = New CAcRate
         Ac.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         Call Ac.QueryData(m_Rs2, iCount)
         Call Ac.PopulateFromRS(m_Rs2)
         Ac.Flag = "I"

         Dim UC As CUcRate
         Set UC = New CUcRate
         UC.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         UC.NULL_FLAG = 0
         Call UC.QueryData(m_Rs2, iCount)
         Call UC.PopulateFromRS(m_Rs2)
         UC.Flag = "I"

         SocFea.Flag = "I"
         Set SocFea.OCRate = OC
         Set SocFea.ACRate = Ac
         Set SocFea.RCRate = Rc
         Set SocFea.UCRate = UC

         Dim StpTierVol As CStpTierVol
         Set StpTierVol = New CStpTierVol

         StpTierVol.STPTIER_VOL_ID = -1
         StpTierVol.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         Call StpTierVol.QueryData(m_Rs2, iCount)
         Set StpTierVol = Nothing
         While Not m_Rs2.EOF
            Set StpTierVol = New CStpTierVol
            Call StpTierVol.PopulateFromRS(m_Rs2)
            StpTierVol.Flag = "I"
            Call SocFea.StepTierVol.add(StpTierVol)
            Set StpTierVol = Nothing
            m_Rs2.MoveNext
         Wend
         Set StpTierVol = Nothing

         Call Ua.SocFeatures.add(SocFea)
         Set SocFea = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QuerySoc = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QuerySoc = False
End Function

Public Function DeleteSoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim SC As CSoc

   RName = "DeleteSoc"
   DeleteSoc = False
   HasBegin = False

   Set SC = New CSoc
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   SC.SOC_ID = UID
   SC.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   
   IsOK = True

   DeleteSoc = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
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

   DeleteSoc = True
End Function
Public Function QueryExWorksPrice(Ua As CExWorksPrice, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryExWorksPrice"
   QueryExWorksPrice = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
         
   If Ua.QueryFlag = 1 Then
      Dim EWPI As CExWorksPriceItem
      Set EWPI = New CExWorksPriceItem
      EWPI.EX_WORKS_PRICE_ID = Ua.EX_WORKS_PRICE_ID
      EWPI.PART_NO = Ua.PART_NO_SEARCH
      EWPI.EX_WORKS_PRICE_ITEM_ID = -1
      EWPI.OrderBy = 1
      Call EWPI.QueryData(1, m_Rs1, iCount)
      Set EWPI = Nothing

      Set Ua.ExWorksPriceItem = Nothing
      Set Ua.ExWorksPriceItem = New Collection
      While Not m_Rs1.EOF
         Set EWPI = New CExWorksPriceItem
         Call EWPI.PopulateFromRS(1, m_Rs1)
          EWPI.Flag = "I"
            Call Ua.ExWorksPriceItem.add(EWPI)
            Set EWPI = Nothing
            m_Rs1.MoveNext
        Wend
        
      Dim EDCI As CExDeliveryCostItem
      Set EDCI = New CExDeliveryCostItem
      EDCI.EX_WORKS_PRICE_ID = Ua.EX_WORKS_PRICE_ID
      EDCI.EX_DELIVERY_COST_ITEM_ID = -1
      EDCI.CUSTOMER_CODE = Ua.CUSTOMER_CODE_SEARCH
      EDCI.OrderBy = 2
      EDCI.OrderType = 1
      Call EDCI.QueryData(1, m_Rs1, iCount)
      Set EDCI = Nothing

      Set Ua.ExDeliveryCost = Nothing
      Set Ua.ExDeliveryCost = New Collection
      While Not m_Rs1.EOF
         Set EDCI = New CExDeliveryCostItem
         Call EDCI.PopulateFromRS(1, m_Rs1)
          EDCI.Flag = "I"
            Call Ua.ExDeliveryCost.add(EDCI, str(EDCI.CUSTOMER_ID) & "-" & str(EDCI.DELIVERY_CUS_ITEM_ID) & "-" & str(EDCI.RATE_TYPE) & "-" & str(EDCI.RATE_TYPE_CUS))
            Set EDCI = Nothing
            m_Rs1.MoveNext
        Wend
        
      Dim EPPI As CExPromotionPartItem
      Set EPPI = New CExPromotionPartItem
      EPPI.EX_WORKS_PRICE_ID = Ua.EX_WORKS_PRICE_ID
      EPPI.EX_PROMOTION_PART_ITEM_ID = -1
      EPPI.PART_NO = Ua.PART_NO_SEARCH
      EPPI.OrderBy = 1
      Call EPPI.QueryData(1, m_Rs1, iCount)
      Set EPPI = Nothing

      Set Ua.ExPromotionPart = Nothing
      Set Ua.ExPromotionPart = New Collection
      While Not m_Rs1.EOF
         Set EPPI = New CExPromotionPartItem
         Call EPPI.PopulateFromRS(1, m_Rs1)
          EPPI.Flag = "I"
            Call Ua.ExPromotionPart.add(EPPI)
            Set EPPI = Nothing
            m_Rs1.MoveNext
        Wend
        
      Dim EPDI As CExPromotionDlcItem
      Set EPDI = New CExPromotionDlcItem
      EPDI.EX_WORKS_PRICE_ID = Ua.EX_WORKS_PRICE_ID
      EPDI.EX_PROMOTION_DLC_ITEM_ID = -1
      EPDI.CUSTOMER_CODE = Ua.CUSTOMER_CODE_SEARCH
      EPDI.OrderBy = 1
      Call EPDI.QueryData(1, m_Rs1, iCount)
      Set EPDI = Nothing

      Set Ua.ExPromotionDlc = Nothing
      Set Ua.ExPromotionDlc = New Collection
      While Not m_Rs1.EOF
         Set EPDI = New CExPromotionDlcItem
         Call EPDI.PopulateFromRS(1, m_Rs1)
          EPDI.Flag = "I"
            Call Ua.ExPromotionDlc.add(EPDI)
            Set EPDI = Nothing
            m_Rs1.MoveNext
        Wend
   End If
   
   QueryExWorksPrice = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryExWorksPrice = False
End Function
Public Function DeleteExWorksPrice(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim EWP As CExWorksPrice

   RName = "DeleteExWorksPrice"
   DeleteExWorksPrice = False
   HasBegin = False

   Set EWP = New CExWorksPrice
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   EWP.EX_WORKS_PRICE_ID = UID
   EWP.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   
   IsOK = True

   DeleteExWorksPrice = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
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

   DeleteExWorksPrice = True
End Function
Private Function PerformUCRate(Ug As CSocFeature, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
Dim S As CStpTierVol
Dim U As CUcRate
Dim Sp As CSystemParam
Dim RateTypeID As Long

   RateTypeID = RATE_FLAT

   If Ug.RATE_TYPE = RateTypeID Then 'Flat
      If Ug.UCRate.Flag = "A" Then
         Ug.UCRate.AddEditMode = SHOW_ADD
         Ug.UCRate.SOC_FEATURE_ID = Ug.SOC_FEATURE_ID
         Call Ug.UCRate.AddEditData
      ElseIf Ug.UCRate.Flag = "E" Then
         Ug.UCRate.AddEditMode = SHOW_EDIT
         Call Ug.UCRate.AddEditData
      ElseIf Ug.UCRate.Flag = "D" Then
         Call Ug.UCRate.DeleteData
      End If
   End If

   Set U = New CUcRate
   For Each S In Ug.StepTierVol
      If S.Flag = "A" Then
         S.AddEditMode = SHOW_ADD
         S.SOC_FEATURE_ID = Ug.SOC_FEATURE_ID
         Call S.AddEditData

         U.AddEditMode = SHOW_ADD
         U.RATE_AMOUNT = S.RATE_AMOUNT
         U.SOC_FEATURE_ID = Ug.SOC_FEATURE_ID
         U.STPTIER_VOL_ID = S.STPTIER_VOL_ID
         Call U.AddEditData
      ElseIf S.Flag = "E" Then
         S.AddEditMode = SHOW_EDIT
         Call S.AddEditData

         U.AddEditMode = SHOW_EDIT
         U.SOC_FEATURE_ID = S.SOC_FEATURE_ID
         U.UC_RATE_ID = S.UC_RATE_ID
         U.STPTIER_VOL_ID = S.STPTIER_VOL_ID
         U.RATE_AMOUNT = S.RATE_AMOUNT
         Call U.AddEditData
      ElseIf S.Flag = "D" Then
         Call U.DeleteData
         Call S.DeleteData
      End If
   Next S
   Set U = Nothing
End Function

Public Function AddEditSoc(Ug As CSoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

Dim SocFea As CSocFeature
Dim OCRate As COcRate
Dim RCRate As CRcRate
Dim UCRate As CUcRate
Dim ACRate As CAcRate

   RName = "AddEditSoc"
   AddEditSoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ug.AddEditData
      
   For Each SocFea In Ug.SocFeatures
      If SocFea.Flag = "A" Then
         SocFea.AddEditMode = SHOW_ADD
         SocFea.SOC_ID = Ug.SOC_ID
         Call SocFea.AddEditData

         Set OCRate = SocFea.OCRate
         OCRate.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         OCRate.AddEditMode = SHOW_ADD
         Call OCRate.AddEditData

         Set RCRate = SocFea.RCRate
         RCRate.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         RCRate.AddEditMode = SHOW_ADD
         Call RCRate.AddEditData

         Set ACRate = SocFea.ACRate
         ACRate.SOC_FEATURE_ID = SocFea.SOC_FEATURE_ID
         ACRate.AddEditMode = SHOW_ADD
         Call ACRate.AddEditData

         Call PerformUCRate(SocFea, IsOK, ErrorObj)
      ElseIf SocFea.Flag = "E" Then
         SocFea.AddEditMode = SHOW_EDIT
         Call SocFea.AddEditData

         Set OCRate = SocFea.OCRate
         OCRate.AddEditMode = SHOW_EDIT
         Call OCRate.AddEditData

         Set RCRate = SocFea.RCRate
         RCRate.AddEditMode = SHOW_EDIT
         Call RCRate.AddEditData

         Set ACRate = SocFea.ACRate
         ACRate.AddEditMode = SHOW_EDIT
         Call ACRate.AddEditData

         Call PerformUCRate(SocFea, IsOK, ErrorObj)
      ElseIf SocFea.Flag = "D" Then
         Call SocFea.DeleteData
      End If
   Next SocFea
         
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   
   IsOK = True

   AddEditSoc = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSoc = False
End Function
Public Function AddEditExWorksPrice(Ug As CExWorksPrice, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

Dim EWPI As CExWorksPriceItem
Dim EDCI As CExDeliveryCostItem
Dim EPPI As CExPromotionPartItem
Dim EPDI As CExPromotionDlcItem
Dim OCRate As COcRate
Dim RCRate As CRcRate
Dim UCRate As CUcRate
Dim ACRate As CAcRate

   RName = "AddEditExWorksPrice"
   AddEditExWorksPrice = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ug.AddEditData
      
   For Each EWPI In Ug.ExWorksPriceItem
      If EWPI.Flag = "A" Then
         EWPI.AddEditMode = SHOW_ADD
         EWPI.EX_WORKS_PRICE_ID = Ug.EX_WORKS_PRICE_ID
         Call EWPI.AddEditData

      ElseIf EWPI.Flag = "E" Then
         EWPI.AddEditMode = SHOW_EDIT
         Call EWPI.AddEditData

      ElseIf EWPI.Flag = "D" Then
         Call EWPI.DeleteData
      End If
   Next EWPI
   
   For Each EDCI In Ug.ExDeliveryCost
      If EDCI.Flag = "A" Then
         EDCI.AddEditMode = SHOW_ADD
         EDCI.EX_WORKS_PRICE_ID = Ug.EX_WORKS_PRICE_ID
         Call EDCI.AddEditData

      ElseIf EDCI.Flag = "E" Then
         EDCI.AddEditMode = SHOW_EDIT
         Call EDCI.AddEditData

      ElseIf EDCI.Flag = "D" Then
         Call EDCI.DeleteData
      End If
   Next EDCI
   
   For Each EPPI In Ug.ExPromotionPart
      If EPPI.Flag = "A" Then
         EPPI.AddEditMode = SHOW_ADD
         EPPI.EX_WORKS_PRICE_ID = Ug.EX_WORKS_PRICE_ID
         Call EPPI.AddEditData

      ElseIf EPPI.Flag = "E" Then
         EPPI.AddEditMode = SHOW_EDIT
         Call EPPI.AddEditData

      ElseIf EPPI.Flag = "D" Then
         Call EPPI.DeleteData
      End If
   Next EPPI
   
   For Each EPDI In Ug.ExPromotionDlc
      If EPDI.Flag = "A" Then
         EPDI.AddEditMode = SHOW_ADD
         EPDI.EX_WORKS_PRICE_ID = Ug.EX_WORKS_PRICE_ID
         Call EPDI.AddEditData

      ElseIf EPDI.Flag = "E" Then
         EPDI.AddEditMode = SHOW_EDIT
         Call EPDI.AddEditData

      ElseIf EPDI.Flag = "D" Then
         Call EPDI.DeleteData
      End If
   Next EPDI
         
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   
   IsOK = True

   AddEditExWorksPrice = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditExWorksPrice = False
End Function
Public Function AddEditPartItemSelect(Ug As CPartItemSelect, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PIS As CPartItemSelect
   RName = "AddEditPartItemSelect"
   AddEditPartItemSelect = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
      
   For Each PIS In Ug.Part_Sel_Coll
      If PIS.Flag = "A" Then
         PIS.AddEditMode = SHOW_ADD
         PIS.PART_ITEM_SELECT_ID = Ug.PART_ITEM_SELECT_ID
         Call PIS.AddEditData

      ElseIf PIS.Flag = "E" Then
         PIS.AddEditMode = SHOW_EDIT
         Call PIS.AddEditData

      ElseIf PIS.Flag = "D" Then
         Call PIS.DeleteData
      End If
   Next PIS
         
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
      HasBegin = False
   End If
   
   IsOK = True

   AddEditPartItemSelect = True
   Exit Function

ErrorHandler:
   If AutoCommit Then
      If HasBegin Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPartItemSelect = False
End Function
Public Function DO2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Area As Long, IvdDocType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDoItem
Dim Ei As CLotItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.DOCUMENT_TYPE = IvdDocType
   Ivd.EXCEPTION_FLAG = "N"
   
   For Each Di In BD.DoItems
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         If Area = 1 Then
            Ei.TX_TYPE = "E"
            Ei.CALCULATE_FLAG = "N"
         ElseIf Area = 2 Then
            Ei.TX_TYPE = "I"
            Ei.CALCULATE_FLAG = "Y"
         End If
         Ei.Flag = "A"
         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
         Ei.LOCATION_ID = Di.LOCATION_ID
         Ei.LOCATION_NAME = Di.LOCATION_NAME
         Ei.TX_AMOUNT = Di.ITEM_AMOUNT
         Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         Ei.SALE_TOT_PRICE = Di.TOTAL_PRICE
'         Ei.INCLUDE_UNIT_PRICE = MyDiff(Ei.SALE_TOT_PRICE, Ei.TX_AMOUNT)
         Ei.LINK_ID = Di.LINK_ID
         
         
         Set Ei.SubLotItems = Di.SubLotItems
         If Di.PART_ITEM_ID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExports.add(Ei)
         End If
         Set Ei = Nothing
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Ei Is Nothing Then
            'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If Di.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItem
               If Area = 1 Then
                  Ei.TX_TYPE = "E"
                  Ei.CALCULATE_FLAG = "N"
               ElseIf Area = 2 Then
                  Ei.TX_TYPE = "I"
                  Ei.CALCULATE_FLAG = "Y"
               End If
               Ei.Flag = "E"
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.ITEM_AMOUNT
               Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
               Ei.SALE_TOT_PRICE = Di.TOTAL_PRICE
               Ei.INCLUDE_UNIT_PRICE = MyDiff(Ei.SALE_TOT_PRICE, Ei.TX_AMOUNT)
               Ei.LINK_ID = Di.LINK_ID
               Set Ei.SubLotItems = Di.SubLotItems
               Call Ivd.ImportExports.add(Ei)
               Set Ei = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If Di.PART_ITEM_ID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Di.Flag
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.ITEM_AMOUNT
               Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
               Ei.SALE_TOT_PRICE = Di.TOTAL_PRICE
               Ei.CALCULATE_FLAG = "N"
               
'               Set Ei.SubLotItems = Di.SubLotItems
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function RT2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Area As Long, IvdDocType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CReceiptItem
Dim Ei As CLotItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.DOCUMENT_TYPE = IvdDocType
   Ivd.EXCEPTION_FLAG = "N"
   
   For Each Di In BD.ReceiptItems
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         If Area = 2 Then
            Ei.TX_TYPE = "E"
            Ei.CALCULATE_FLAG = "N"
         ElseIf Area = 1 Then
            Ei.TX_TYPE = "I"
            Ei.CALCULATE_FLAG = "Y"
         End If
         Ei.Flag = "A"
         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
         Ei.LOCATION_ID = Di.LOCATION_ID
         Ei.LOCATION_NAME = Di.LOCATION_NAME
         Ei.TX_AMOUNT = Di.RETURN_AMOUNT
         'Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         Ei.SALE_TOT_PRICE = Di.RETURN_AMOUNT * Di.RETURN_AVG_PRICE
         Ei.ACTUAL_UNIT_PRICE = Di.RETURN_AVG_PRICE
         Ei.INCLUDE_UNIT_PRICE = Ei.ACTUAL_UNIT_PRICE
         Ei.TOTAL_ACTUAL_PRICE = Di.RETURN_AMOUNT * Di.RETURN_AVG_PRICE
         Ei.TOTAL_INCLUDE_PRICE = Ei.TOTAL_ACTUAL_PRICE
         Ei.LINK_ID = Di.LINK_ID
         
         'Set Ei.SubLotItems = Di.SubLotItems
         If Di.PART_ITEM_ID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExports.add(Ei)
         End If
         Set Ei = Nothing
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Ei Is Nothing Then
            'มีข้อมูลใน ReceiptItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If Di.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItem
               If Area = 2 Then
                  Ei.TX_TYPE = "E"
                  Ei.CALCULATE_FLAG = "N"
               ElseIf Area = 1 Then
                  Ei.TX_TYPE = "I"
                  Ei.CALCULATE_FLAG = "Y"
               End If
               Ei.Flag = "E"
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.RETURN_AMOUNT
               'Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
               Ei.SALE_TOT_PRICE = Di.RETURN_AMOUNT * Di.RETURN_AVG_PRICE
               Ei.ACTUAL_UNIT_PRICE = Di.RETURN_AVG_PRICE
               Ei.INCLUDE_UNIT_PRICE = Ei.ACTUAL_UNIT_PRICE
               Ei.TOTAL_ACTUAL_PRICE = Di.RETURN_AMOUNT * Di.RETURN_AVG_PRICE
               Ei.TOTAL_INCLUDE_PRICE = Ei.TOTAL_ACTUAL_PRICE
               Ei.LINK_ID = Di.LINK_ID
               'Set Ei.SubLotItems = Di.SubLotItems
               Call Ivd.ImportExports.add(Ei)
               Set Ei = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If Di.PART_ITEM_ID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Di.Flag
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.RETURN_AMOUNT
               'Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
               Ei.SALE_TOT_PRICE = Di.RETURN_AMOUNT * Di.RETURN_AVG_PRICE
               Ei.ACTUAL_UNIT_PRICE = Di.RETURN_AVG_PRICE
               Ei.INCLUDE_UNIT_PRICE = Ei.ACTUAL_UNIT_PRICE
               Ei.TOTAL_ACTUAL_PRICE = Di.RETURN_AMOUNT * Di.RETURN_AVG_PRICE
               Ei.TOTAL_INCLUDE_PRICE = Ei.TOTAL_ACTUAL_PRICE

               Ei.CALCULATE_FLAG = "N"
               
'               Set Ei.SubLotItems = Di.SubLotItems
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function SUP2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, IvdDocType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CSupItem
Dim Ei As CLotItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc
   
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      If BD.AddEditMode = SHOW_EDIT Then
         glbErrorLog.LocalErrorMsg = "การค้นหา และบันทึกข้อมูล ERROR กรุณาลบเอกสารเลขที่ " & BD.DOCUMENT_NO & " และ Import ใหม่"
         glbErrorLog.ShowUserError
         Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
         Ivd.AddEditMode = SHOW_ADD
      Else
         Ivd.AddEditMode = SHOW_ADD
      End If
   End If
      
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DO_NO = BD.DO_NO
   Ivd.TRUCK_NO = BD.TRUCK_NO
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.DELIVERY_FEE = BD.DELIVERY_FEE
   Ivd.SENDER_NAME = BD.SENDER_NAME
   Ivd.RECEIVE_NAME = BD.RECEIVE_NAME
   Ivd.SUPPLIER_ID = BD.SUPPLIER_ID
   Ivd.SUPPLIER_ID_TRUE = BD.SUPPLIER_ID_TRUE
   Ivd.DOCUMENT_TYPE = IvdDocType
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.QUE_NO = BD.QUE_NO
   Ivd.DOCUMENT_DESC = BD.NOTE
   Ivd.ENTRY_DATE = BD.ENTRY_DATE
   Ivd.EXIT_DATE = BD.EXIT_DATE
   
   Ivd.EXCEPTION_FLAG = BD.EXCEPTION_FLAG
   Ivd.DEPARTMENT_ID = BD.DEPARTMENT_ID
   Ivd.Credit = BD.Credit
   Ivd.PR_NO = BD.PR_NO
   
'   For Each Ei In Ivd.ImportExports
'      Ei.Flag = "D"
'   Next Ei
   
   For Each Di In BD.SupItems
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         Ei.Flag = "A"
         
         Call Ei.CopyObjectFromSup(1, Di)
         Call Ivd.ImportExports.add(Ei)
         Set Ei = Nothing
         
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Ei Is Nothing Then
            Set Ei = New CLotItem
         
            Ei.Flag = "A"
            
            Call Ei.CopyObjectFromSup(1, Di)
            Call Ivd.ImportExports.add(Ei)
            Set Ei = Nothing
         Else
            Ei.Flag = "E"
            Call Ei.CopyObjectFromSup(1, Di)
         End If
         
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function SUP2InventoryDocEx(BD As CBillingDoc, Ivd As CInventoryDoc, IvdDocType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CSupItem
Dim Ei As CLotItem
Dim II As CLotItem
   
   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DO_NO = BD.DO_NO
   Ivd.TRUCK_NO = BD.TRUCK_NO
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.DELIVERY_FEE = BD.DELIVERY_FEE
   Ivd.SENDER_NAME = BD.SENDER_NAME
   Ivd.RECEIVE_NAME = BD.RECEIVE_NAME
   Ivd.SUPPLIER_ID = BD.SUPPLIER_ID
   Ivd.DOCUMENT_TYPE = IvdDocType
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.QUE_NO = BD.QUE_NO
   Ivd.DOCUMENT_DESC = BD.NOTE
   Ivd.ENTRY_DATE = BD.ENTRY_DATE
   Ivd.EXIT_DATE = BD.EXIT_DATE
   
   Ivd.EXCEPTION_FLAG = BD.EXCEPTION_FLAG
   Ivd.DEPARTMENT_ID = BD.DEPARTMENT_ID
   Ivd.Credit = BD.Credit
   Ivd.PR_NO = BD.PR_NO
   
   For Each Ei In Ivd.ImportExports
      Ei.Flag = "D"
   Next Ei
   
   For Each Di In BD.SupItems
      If Di.Flag <> "D" Then
         Set Ei = New CLotItem
         Set II = New CLotItem
         Ei.Flag = "A"
         II.Flag = "A"
         
         Call Ei.CopyObjectFromSup(1, Di)
         Call II.CopyObjectFromSup(1, Di)
         
         Ei.TX_TYPE = "I"
         II.TX_TYPE = "E"
         
         Call Ivd.ImportExports.add(Ei)
         Call Ivd.ImportExports.add(II)
         Set Ei = Nothing
         Set II = Nothing
      End If
   Next Di
End Function

Public Function Job2InventoryDoc(BD As CJob, Ivd As CInventoryDoc, Area As Long, DocumentType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CJobInput
Dim Ei As CLotItem
Dim TempPartNo As String

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = BD.JOB_DATE
   Ivd.DOCUMENT_NO = BD.JOB_NO
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.DOCUMENT_TYPE = ProcessIDToInventoryDoc(BD.PROCESS_ID)
      
   For Each Di In BD.Inputs
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         Ei.TX_TYPE = "E"
         Ei.CALCULATE_FLAG = "Y"
         Ei.Flag = "A"
         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
         Ei.LOCATION_ID = Di.LOCATION_ID
         Ei.LOCATION_NAME = Di.LOCATION_NAME
         Ei.TX_AMOUNT = Di.TX_AMOUNT
         Ei.TOTAL_WEIGHT = 0
         Ei.TOTAL_INCLUDE_PRICE = 0
         Ei.INCLUDE_UNIT_PRICE = 0
         Ei.LINK_ID = Di.LINK_ID
         
'         Set Ei.SubLotItems = Di.SubLotItems
         If Di.PART_ITEM_ID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExports.add(Ei)
         End If
         Set Ei = Nothing
         

          
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Ei Is Nothing Then
            'มีข้อมูลใน Inputs แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If Di.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItem
               Ei.TX_TYPE = "E"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Di.Flag
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.TX_AMOUNT
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Di.LINK_ID
'               Set Ei.SubLotItems = Di.SubLotItems
               Call Ivd.ImportExports.add(Ei)
               Set Ei = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If Di.PART_ITEM_ID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Di.Flag
               Ei.TX_TYPE = "E"
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.TX_AMOUNT
               Ei.TOTAL_WEIGHT = 0
               Ei.CALCULATE_FLAG = "Y"

'               Set Ei.SubLotItems = Di.SubLotItems
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di

   '===
   For Each Di In BD.Outputs
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         Ei.TX_TYPE = "I"
         Ei.CALCULATE_FLAG = "Y"
         Ei.Flag = "A"
         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
         Ei.LOCATION_ID = Di.LOCATION_ID
         Ei.LOCATION_NAME = Di.LOCATION_NAME
         Ei.TX_AMOUNT = Di.TX_AMOUNT
         Ei.TOTAL_WEIGHT = 0
         Ei.TOTAL_INCLUDE_PRICE = Di.AVG_PRICE * Di.TX_AMOUNT
         Ei.INCLUDE_UNIT_PRICE = Di.AVG_PRICE
         Ei.LINK_ID = Di.LINK_ID
         
'         Set Ei.SubLotItems = Di.SubLotItems
         If Di.PART_ITEM_ID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            If BD.PROCESS_ID = 8 Then
               TempPartNo = "Re-bg-other (" & Di.PART_DESC & ")"
               BD.JOB_DESC = TempPartNo
               Ivd.DOCUMENT_DESC = TempPartNo
            End If
            Call Ivd.ImportExports.add(Ei)
         End If
         Set Ei = Nothing
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Ei Is Nothing Then
            'มีข้อมูลใน Inputs แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If Di.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItem
               Ei.TX_TYPE = "I"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Di.Flag
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.TX_AMOUNT
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Di.LINK_ID
'               Set Ei.SubLotItems = Di.SubLotItems
               Call Ivd.ImportExports.add(Ei)
               Set Ei = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If Di.PART_ITEM_ID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Di.Flag
               Ei.TX_TYPE = "I"
               Ei.PART_ITEM_ID = Di.PART_ITEM_ID
               Ei.LOCATION_ID = Di.LOCATION_ID
               Ei.LOCATION_NAME = Di.LOCATION_NAME
               Ei.TX_AMOUNT = Di.TX_AMOUNT
               Ei.TOTAL_WEIGHT = 0
               Ei.CALCULATE_FLAG = "Y"
            
            If BD.PROCESS_ID = 8 Then
               TempPartNo = "Re-bg-other (" & Di.PART_DESC & ")"
               BD.JOB_DESC = TempPartNo
               Ivd.DOCUMENT_DESC = TempPartNo
            End If

'               Set Ei.SubLotItems = Di.SubLotItems
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function Job2InventoryWhDoc(BD As CJob, Ivd As CInventoryWHDoc, Area As Long, DocumentType As Long, Optional typeInput As Long, Optional TypeCopy As Long = 1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CInventoryWHDoc
Dim Di2 As CInventoryWHDoc
Dim Ei As CLotItemWH
Dim Ei2 As CLotItemWH
Dim Ei3 As CLotItemWH
Dim m_Lot As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim Pdi As CPalletDoc
Dim TempLTD As CLotDoc
Dim TempPD As Collection
Dim TempPartNo As String

   Set Ivd = Nothing
   Set Ivd = New CInventoryWHDoc

   If BD.INVENTORY_WH_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset

      Ivd.INVENTORY_WH_DOC_ID = BD.INVENTORY_WH_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryWhDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)

      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
   
   Ivd.DOCUMENT_DATE = BD.JOB_DATE
   Ivd.DOCUMENT_NO = BD.JOB_NO
   Ivd.DOCUMENT_DESC = BD.JOB_DESC
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.DOCUMENT_TYPE = ProcessIDToInventoryDoc(BD.PROCESS_ID)
   Ivd.BATCH_NO = BD.BATCH_NO
   Ivd.PROCESS_ID = BD.PROCESS_ID
   Ivd.APPROVED_BY = BD.APPROVED_BY
   Ivd.RESPONSE_BY = BD.RESPONSE_BY
   Ivd.START_DATE = BD.START_DATE
   Ivd.FINISH_DATE = BD.FINISH_DATE
   Ivd.SUCCESS_FLAG = BD.COMMIT_FLAG
   Ivd.LOAD_FLAG = "N"
   Ivd.Flag = BD.Flag

     '===
     Dim ID1 As Long
   For Each Di In BD.InventoryWhDoc
     If Di.C_LotItemsWH.Count = 0 Then
      GoTo Continue 'Exit For
     End If
     If Di.Flag = "D" Then
         Set Ivd = Di
         GoTo Continue 'Exit For
     End If
     
     Set Ei2 = Di.C_LotItemsWH.Item(1)
     
     
      If BD.PROCESS_ID <> 4 Then
         If typeInput <> 1 Then
            If BD.PROCESS_ID = 6 Then
              If Ei2.InputRebagToBagType = 3 Then
               TempPartNo = "*ตัดแตก* (" & Ei2.PART_NO & ")(" & Ei2.LOT_NO & ")"
             ElseIf Ei2.InputRebagToBagType = 5 Then
               TempPartNo = "*รับเข้าจากการตัดแตก* (" & Ei2.PART_NO & ")(" & Ei2.LOT_NO & ")"
              ElseIf Ei2.InputRebagToBagType = 1 Then
                TempPartNo = "Re-bg-bg (" & Ei2.PART_NO & ")(" & Ei2.LOT_NO & ")"
               Else
                  TempPartNo = BD.JOB_DESC
               End If
            ElseIf BD.PROCESS_ID = 7 Then
               TempPartNo = "Re-bg-bk (" & Ei2.PART_NO & ")(" & Ei2.LOT_NO & ")"
            Else
              TempPartNo = Ei2.PART_NO & " (" & Ei2.LOT_NO & ")"
            End If
              BD.JOB_DESC = TempPartNo
              Ivd.DOCUMENT_DESC = TempPartNo
         End If
      End If
      
      'กรณี เปลี่ยนแปลง Lot
      Dim LotColls As Collection
      Dim strDate As String
      Set LotColls = New Collection
      For Each LTD In Ei2.C_LotDoc
         If BD.JOB_DATE <> LTD.LOT_DATE Then
               If Ei2.Flag <> "A" Then
                 Ei2.Flag = "E"
                 TypeCopy = 2
               End If
         End If
      Next LTD
           
      If Ei2.Flag = "A" Then
         Set Ei = New CLotItemWH
               Ei.TX_TYPE = "I"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Ei2.Flag
               Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
               Ei.LOCATION_ID = Ei2.LOCATION_ID
               Ei.LOCATION_REST_ID = Ei2.LOCATION_REST_ID
               Ei.LOCATION_LOSE_ID = Ei2.LOCATION_LOSE_ID
               Ei.PRODUCT_TYPE_ID = Ei2.PRODUCT_TYPE_ID
               Ei.LOSE_AMOUNT = Ei2.LOSE_AMOUNT
               Ei.REST_AMOUNT = Ei2.REST_AMOUNT
               Ei.BL_START_DATE = Ei2.BL_START_DATE
               Ei.PACK_DATE = Ei2.PACK_DATE
               Ei.PACK_AMOUNT = Ei2.PACK_AMOUNT
               Ei.TIME_PACK_BEGIN = Ei2.TIME_PACK_BEGIN
               Ei.TIME_PACK_END = Ei2.TIME_PACK_END
               Ei.NOTE = Ei2.NOTE
               Ei.LOCATION_NAME = Ei2.LOCATION_NAME
               If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                  Ei.TX_AMOUNT = Ei2.TX_AMOUNT
               Else
                  Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT * Ei2.WEIGHT_PER_PACK
               End If
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Ei2.LINK_ID
               Ei.BIN_NO = Ei2.BIN_NO
               Ei.PALLET_NO = Ei2.PALLET_NO
               Ei.LOCK_NO = Ei2.LOCK_NO
               Ei.HEAD_PACK_NO = Ei2.HEAD_PACK_NO
               Ei.GOOD_AMOUNT = Ei2.GOOD_AMOUNT
               Ei.WEIGHT_PER_PACK = Ei2.WEIGHT_PER_PACK
               Ei.FULL_PALLET_FROM = Ei2.FULL_PALLET_FROM
               Ei.FULL_PALLET_TO = Ei2.FULL_PALLET_TO
               Ei.FULL_UNIT_PER_PALLET = Ei2.FULL_UNIT_PER_PALLET
               Ei.SCRAP_PALLET = Ei2.SCRAP_PALLET
               Ei.SCRAP_UNIT_PER_PALLET = Ei2.SCRAP_UNIT_PER_PALLET
               Ei.AddEditMode = SHOW_ADD
               
                Set LTD = New CLotDoc
               For Each LTD In Ei2.C_LotDoc
                   Call Ei.C_LotDoc.add(LTD)
               Next LTD
               Call Ivd.C_LotItemsWH.add(Ei) 'เก็บไว้ใน Lot item ด้วย
               Set Ei = Nothing
      ElseIf Ei2.Flag = "E" Then
         Set Ei = GetExportItemWh(Ivd, Ei2.LINK_ID)
         If Ei Is Nothing Then
            If TypeCopy = 2 Then
                  For Each Ei In Ivd.C_LotItemsWH
                     Ei.LOT_ITEM_WH_ID = Ei2.LOT_ITEM_WH_ID
                     Ei.INVENTORY_WH_DOC_ID = Ei2.INVENTORY_WH_DOC_ID
                     Ei.TX_TYPE = "I"
                     Ei.CALCULATE_FLAG = "Y"
                     Ei.Flag = Ei2.Flag
                     Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
                     Ei.LOCATION_ID = Ei2.LOCATION_ID
                     Ei.LOCATION_REST_ID = Ei2.LOCATION_REST_ID
                     Ei.LOCATION_LOSE_ID = Ei2.LOCATION_LOSE_ID
                     Ei.PRODUCT_TYPE_ID = Ei2.PRODUCT_TYPE_ID
                     Ei.LOSE_AMOUNT = Ei2.LOSE_AMOUNT
                     Ei.REST_AMOUNT = Ei2.REST_AMOUNT
                     Ei.PACK_AMOUNT = Ei2.PACK_AMOUNT
                     Ei.BL_START_DATE = Ei2.BL_START_DATE
                     Ei.PACK_DATE = Ei2.PACK_DATE
                     Ei.TIME_PACK_BEGIN = Ei2.TIME_PACK_BEGIN
                     Ei.TIME_PACK_END = Ei2.TIME_PACK_END
                     Ei.NOTE = Ei2.NOTE
                     Ei.LOCATION_NAME = Ei2.LOCATION_NAME
                     If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                        Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT 'Ei2.TX_AMOUNT
                     Else
                        Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT * Ei2.WEIGHT_PER_PACK
                     End If
                     Ei.TOTAL_WEIGHT = 0
                     Ei.TOTAL_INCLUDE_PRICE = 0
                     Ei.INCLUDE_UNIT_PRICE = 0
                     Ei.LINK_ID = Ei2.LINK_ID
                     Ei.BIN_NO = Ei2.BIN_NO
                     Ei.PALLET_NO = Ei2.PALLET_NO
                     Ei.LOCK_NO = Ei2.LOCK_NO
                     Ei.HEAD_PACK_NO = Ei2.HEAD_PACK_NO
                     Ei.GOOD_AMOUNT = Ei2.GOOD_AMOUNT
                     Ei.WEIGHT_PER_PACK = Ei2.WEIGHT_PER_PACK
                     Ei.FULL_PALLET_FROM = Ei2.FULL_PALLET_FROM
                     Ei.FULL_PALLET_TO = Ei2.FULL_PALLET_TO
                     Ei.FULL_UNIT_PER_PALLET = Ei2.FULL_UNIT_PER_PALLET
                     Ei.SCRAP_PALLET = Ei2.SCRAP_PALLET
                     Ei.SCRAP_UNIT_PER_PALLET = Ei2.SCRAP_UNIT_PER_PALLET
                     Ei.AddEditMode = SHOW_EDIT
                     
                     For Each LTD In Ei.C_LotDoc
                         Set TempLTD = Ei2.C_LotDoc(1)
                         LTD.Flag = TempLTD.Flag
                         LTD.LOT_ID = TempLTD.LOT_ID
                         LTD.BIN_NO = TempLTD.BIN_NO
                        If BD.JOB_DATE < LTD.LOT_DATE Then 'กรณี วันที่เอกสารโดยวันที่น้อยกว่าวันที่เดิมเปลี่ยน ชื่อ Lot ต้องเปลี่ยนด้วย
                           Call LoadLotFromLot(Nothing, LotColls, , BD.JOB_DATE, BD.JOB_DATE, 2, , 4)
                           LTD.AddEditMode = SHOW_EDIT
                           LTD.Flag = "E"
                           LTD.LOT_ID = CreateLotAuto(BD.JOB_DATE, Val(Mid(LTD.LOT_NO, 9, 3)), LotColls)
                           LTD.IS_EDIT_LOT = True
                        End If
                       
                        For Each PD In LTD.C_PalletDoc
                           PD.Flag = Ei.Flag
                           If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                              PD.CAPACITY_AMOUNT = Ei.TX_AMOUNT
                           Else 'ถ้าเป็น Bag
                             Set TempPD = Ei2.C_LotDoc(1).C_PalletDoc
                             Set LTD.C_PalletDoc = TempPD
                           End If
                        Next PD
                     Next LTD
                  Next Ei
            ElseIf Ei2.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItemWH
               Ei.LOT_ITEM_WH_ID = Ei2.LOT_ITEM_WH_ID
               Ei.INVENTORY_WH_DOC_ID = Ei2.INVENTORY_WH_DOC_ID
               Ei.TX_TYPE = "I"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Ei2.Flag
               Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
               Ei.LOCATION_ID = Ei2.LOCATION_ID
               Ei.LOCATION_REST_ID = Ei2.LOCATION_REST_ID
               Ei.LOCATION_LOSE_ID = Ei2.LOCATION_LOSE_ID
               Ei.PRODUCT_TYPE_ID = Ei2.PRODUCT_TYPE_ID
               Ei.LOSE_AMOUNT = Ei2.LOSE_AMOUNT
               Ei.REST_AMOUNT = Ei2.REST_AMOUNT
               Ei.PACK_AMOUNT = Ei2.PACK_AMOUNT
               Ei.BL_START_DATE = Ei2.BL_START_DATE
               Ei.PACK_DATE = Ei2.PACK_DATE
               Ei.TIME_PACK_BEGIN = Ei2.TIME_PACK_BEGIN
               Ei.TIME_PACK_END = Ei2.TIME_PACK_END
               Ei.NOTE = Ei2.NOTE
               Ei.LOCATION_NAME = Ei2.LOCATION_NAME
               If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                  Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT 'Ei2.TX_AMOUNT
               Else
                  Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT * Ei2.WEIGHT_PER_PACK
               End If
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Ei2.LINK_ID
               Ei.BIN_NO = Ei2.BIN_NO
               Ei.PALLET_NO = Ei2.PALLET_NO
               Ei.LOCK_NO = Ei2.LOCK_NO
               Ei.HEAD_PACK_NO = Ei2.HEAD_PACK_NO
               Ei.GOOD_AMOUNT = Ei2.GOOD_AMOUNT
               Ei.WEIGHT_PER_PACK = Ei2.WEIGHT_PER_PACK
               Ei.FULL_PALLET_FROM = Ei2.FULL_PALLET_FROM
               Ei.FULL_PALLET_TO = Ei2.FULL_PALLET_TO
               Ei.FULL_UNIT_PER_PALLET = Ei2.FULL_UNIT_PER_PALLET
               Ei.SCRAP_PALLET = Ei2.SCRAP_PALLET
               Ei.SCRAP_UNIT_PER_PALLET = Ei2.SCRAP_UNIT_PER_PALLET
               Ei.AddEditMode = SHOW_EDIT
               
               Set LTD = New CLotDoc
               For Each LTD In Ei2.C_LotDoc
                   Call Ei.C_LotDoc.add(LTD)
                Next LTD
                
               Call Ivd.C_LotItemsWH.add(Ei) 'เก็บไว้ใน Lot item ด้วย
        
               Set Ei = Nothing
               
               Set Ei3 = Nothing
            End If
         End If
      ElseIf Ei2.Flag = "D" Then
          Set Ei = New CLotItemWH

          Ei.Flag = Ei2.Flag
          Ei.INVENTORY_WH_DOC_ID = Ei2.INVENTORY_WH_DOC_ID
          Ei.LOT_ITEM_WH_ID = Ei2.LOT_ITEM_WH_ID
         Set LTD = New CLotDoc
         For Each LTD In Ei2.C_LotDoc
            Call Ei.C_LotDoc.add(LTD)
         Next LTD
         Call Ivd.C_LotItemsWH.add(Ei) 'เก็บไว้ใน Lot item ด้วย
          Set Ei = Nothing
         
      End If
      Set Ei2 = Nothing
Continue:
   Next Di
End Function
Public Function Job2InventoryWhDocInput(BD As CJob, Ivd As CInventoryWHDoc, Area As Long, DocumentType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CInventoryWHDoc
Dim Ei As CLotItemWH
Dim Ei2 As CLotItemWH
Dim Ei3 As CLotItemWH
Dim m_Lot As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim Pdi As CPalletDoc
Dim TempLTD As CLotDoc
Dim TempPD As CPalletDoc
Dim TempPartNo As String

   Set Ivd = Nothing
   Set Ivd = New CInventoryWHDoc

   If BD.INVENTORY_WH_DOC_ID_INPUT > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_WH_DOC_ID = BD.INVENTORY_WH_DOC_ID_INPUT
      Ivd.QueryFlag = 1
'      Call glbDaily.QueryInventoryWhDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
   
   Ivd.DOCUMENT_DESC = ""
   Ivd.DOCUMENT_DATE = BD.JOB_DATE
   Ivd.DOCUMENT_NO = BD.JOB_NO
   Ivd.DOCUMENT_DESC = BD.JOB_DESC
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.DOCUMENT_TYPE = ProcessIDToInventoryDocInput(BD.PROCESS_ID)
   Ivd.BATCH_NO = BD.BATCH_NO
   Ivd.PROCESS_ID = BD.PROCESS_ID
   Ivd.APPROVED_BY = BD.APPROVED_BY
   Ivd.RESPONSE_BY = BD.RESPONSE_BY
   Ivd.START_DATE = BD.START_DATE
   Ivd.FINISH_DATE = BD.FINISH_DATE
   Ivd.SUCCESS_FLAG = BD.COMMIT_FLAG
   Ivd.LOAD_FLAG = "N"
'   Ivd.Flag = Bd.Flag

     '===
     Dim ID1 As Long
   For Each Di In BD.InventoryWhDocInput
   Ivd.Flag = Di.Flag
     For Each Ei2 In Di.C_LotItemsWH
      If Ei2.Flag = "A" Then
         Set Ei = New CLotItemWH
               Ei.TX_TYPE = "E"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Ei2.Flag
               Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
               Ei.LOCATION_ID = Ei2.LOCATION_ID
               Ei.LOCATION_REST_ID = Ei2.LOCATION_REST_ID
               Ei.LOCATION_LOSE_ID = Ei2.LOCATION_LOSE_ID
               Ei.PRODUCT_TYPE_ID = Ei2.PRODUCT_TYPE_ID
               Ei.LOSE_AMOUNT = Ei2.LOSE_AMOUNT
               Ei.REST_AMOUNT = Ei2.REST_AMOUNT
               Ei.PACK_DATE = Ei2.PACK_DATE
               Ei.PACK_AMOUNT = Ei2.PACK_AMOUNT
               Ei.TIME_PACK_BEGIN = Ei2.TIME_PACK_BEGIN
               Ei.TIME_PACK_END = Ei2.TIME_PACK_END
               Ei.NOTE = Ei2.NOTE
               Ei.LOCATION_NAME = Ei2.LOCATION_NAME
               If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                  Ei.TX_AMOUNT = Ei2.TX_AMOUNT
               Else
                  Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT * Ei2.WEIGHT_PER_PACK
               End If
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Ei2.LINK_ID
               Ei.BIN_NO = Ei2.BIN_NO
               Ei.PALLET_NO = Ei2.PALLET_NO
               Ei.LOCK_NO = Ei2.LOCK_NO
               Ei.HEAD_PACK_NO = Ei2.HEAD_PACK_NO
               Ei.GOOD_AMOUNT = Ei2.GOOD_AMOUNT
               Ei.WEIGHT_PER_PACK = Ei2.WEIGHT_PER_PACK
               Ei.FULL_PALLET_FROM = Ei2.FULL_PALLET_FROM
               Ei.FULL_PALLET_TO = Ei2.FULL_PALLET_TO
               Ei.FULL_UNIT_PER_PALLET = Ei2.FULL_UNIT_PER_PALLET
               Ei.SCRAP_PALLET = Ei2.SCRAP_PALLET
               Ei.SCRAP_UNIT_PER_PALLET = Ei2.SCRAP_UNIT_PER_PALLET
               Ei.AddEditMode = SHOW_ADD
               
                Set LTD = New CLotDoc
               For Each LTD In Ei2.C_LotDoc
                   Call Ei.C_LotDoc.add(LTD)
                   For Each PD In LTD.C_PalletDoc
                     PD.TEMP_PALLET_CAP_LAST = 0 'เครียร์ ตัวนี้ให้เป็น 0
                   Next PD
                Next LTD
               Call Ivd.C_LotItemsWH.add(Ei) 'เก็บไว้ใน Lot item ด้วย
               Set Ei = Nothing
      ElseIf Ei2.Flag = "E" Then
         Set Ei = GetExportItemWh(Ivd, Ei2.LINK_ID)
         If Ei Is Nothing Then
            If Ei2.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItemWH
               Ei.LOT_ITEM_WH_ID = Ei2.LOT_ITEM_WH_ID
               Ei.INVENTORY_WH_DOC_ID = Ei2.INVENTORY_WH_DOC_ID
               Ei.TX_TYPE = "E"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Ei2.Flag
               Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
               Ei.LOCATION_ID = Ei2.LOCATION_ID
               Ei.LOCATION_REST_ID = Ei2.LOCATION_REST_ID
               Ei.LOCATION_LOSE_ID = Ei2.LOCATION_LOSE_ID
               Ei.PRODUCT_TYPE_ID = Ei2.PRODUCT_TYPE_ID
               Ei.LOSE_AMOUNT = Ei2.LOSE_AMOUNT
               Ei.REST_AMOUNT = Ei2.REST_AMOUNT
               Ei.PACK_AMOUNT = Ei2.PACK_AMOUNT
               Ei.PACK_DATE = Ei2.PACK_DATE
               Ei.TIME_PACK_BEGIN = Ei2.TIME_PACK_BEGIN
               Ei.TIME_PACK_END = Ei2.TIME_PACK_END
               Ei.NOTE = Ei2.NOTE
               Ei.LOCATION_NAME = Ei2.LOCATION_NAME
               If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                  Ei.TX_AMOUNT = Ei2.TX_AMOUNT
               Else
                  Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT * Ei2.WEIGHT_PER_PACK
               End If
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Ei2.LINK_ID
               Ei.BIN_NO = Ei2.BIN_NO
               Ei.PALLET_NO = Ei2.PALLET_NO
               Ei.LOCK_NO = Ei2.LOCK_NO
               Ei.HEAD_PACK_NO = Ei2.HEAD_PACK_NO
               Ei.GOOD_AMOUNT = Ei2.GOOD_AMOUNT
               Ei.WEIGHT_PER_PACK = Ei2.WEIGHT_PER_PACK
               Ei.FULL_PALLET_FROM = Ei2.FULL_PALLET_FROM
               Ei.FULL_PALLET_TO = Ei2.FULL_PALLET_TO
               Ei.FULL_UNIT_PER_PALLET = Ei2.FULL_UNIT_PER_PALLET
               Ei.SCRAP_PALLET = Ei2.SCRAP_PALLET
               Ei.SCRAP_UNIT_PER_PALLET = Ei2.SCRAP_UNIT_PER_PALLET
               Ei.AddEditMode = SHOW_EDIT
               
               Set LTD = New CLotDoc
               For Each LTD In Ei2.C_LotDoc
                   Call Ei.C_LotDoc.add(LTD)
                    For Each PD In LTD.C_PalletDoc
                     PD.TEMP_PALLET_CAP_LAST = 0 'เครียร์ ตัวนี้ให้เป็น 0
                   Next PD
                Next LTD
               Call Ivd.C_LotItemsWH.add(Ei) 'เก็บไว้ใน Lot item ด้วย
               Set Ei = Nothing
               
               Set Ei3 = Nothing
            End If
         End If
      ElseIf Ei2.Flag = "D" Then
          Set Ei = New CLotItemWH

          Ei.Flag = Ei2.Flag
          Ei.INVENTORY_WH_DOC_ID = Ei2.INVENTORY_WH_DOC_ID
          Ei.LOT_ITEM_WH_ID = Ei2.LOT_ITEM_WH_ID
          Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
         Set LTD = New CLotDoc
         For Each LTD In Ei2.C_LotDoc
            Call Ei.C_LotDoc.add(LTD)
         Next LTD
         Call Ivd.C_LotItemsWH.add(Ei) 'เก็บไว้ใน Lot item ด้วย
          Set Ei = Nothing
      End If
      Next Ei2
'      Set Ei2 = Nothing
Continue:
   Next Di
End Function
Public Function Job2InventoryWhDocForCopy(BD As CJob, Ivd As CInventoryWHDoc, Area As Long, DocumentType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CInventoryWHDoc
Dim Ei As CLotItemWH
Dim Ei2 As CLotItemWH
Dim Ei3 As CLotItemWH
Dim m_Lot As cLot
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim Pdi As CPalletDoc
Dim TempLTD As CLotDoc
Dim TempPD As CPalletDoc

   Set Ivd = Nothing
   Set Ivd = New CInventoryWHDoc

   If BD.INVENTORY_WH_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_WH_DOC_ID = BD.INVENTORY_WH_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryWhDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
   
   
   Ivd.DOCUMENT_DATE = BD.JOB_DATE
   Ivd.DOCUMENT_NO = BD.JOB_NO
   Ivd.DOCUMENT_DESC = BD.JOB_DESC
   Ivd.COMMIT_FLAG = BD.COMMIT_FLAG
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.DOCUMENT_TYPE = ProcessIDToInventoryDoc(BD.PROCESS_ID)
   Ivd.LOT_ID = BD.LOT_ID
   Ivd.BATCH_NO = BD.BATCH_NO
   Ivd.PROCESS_ID = BD.PROCESS_ID
   Ivd.APPROVED_BY = BD.APPROVED_BY
   Ivd.RESPONSE_BY = BD.RESPONSE_BY
   Ivd.START_DATE = BD.START_DATE
   Ivd.FINISH_DATE = BD.FINISH_DATE
   Ivd.SUCCESS_FLAG = BD.COMMIT_FLAG
   Ivd.LOAD_FLAG = "N"
   Ivd.Flag = BD.Flag

     '===
   For Each Di In BD.InventoryWhDoc
     If Di.C_LotItemsWH.Count = 0 Then
      Exit For
     End If
     Set Ei2 = Di.C_LotItemsWH.Item(1) 'GetObject("CLotItemWH", Di.C_LotItemsWH, "", False)ตรงนี้ยังerror
     
      If Ei2.Flag = "A" Then
      ElseIf Ei2.Flag = "E" Then
            If Ei2.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItemWH
               Ei.LOT_ITEM_WH_ID = Ei2.LOT_ITEM_WH_ID
               Ei.INVENTORY_WH_DOC_ID = Ei2.INVENTORY_WH_DOC_ID
               Ei.TX_TYPE = "I"
               Ei.CALCULATE_FLAG = "Y"
               Ei.Flag = Ei2.Flag
               Ei.PART_ITEM_ID = Ei2.PART_ITEM_ID
               Ei.LOCATION_ID = Ei2.LOCATION_ID
               Ei.LOCATION_REST_ID = Ei2.LOCATION_REST_ID
               Ei.LOCATION_LOSE_ID = Ei2.LOCATION_LOSE_ID
               Ei.PRODUCT_TYPE_ID = Ei2.PRODUCT_TYPE_ID
               Ei.LOSE_AMOUNT = Ei2.LOSE_AMOUNT
               Ei.REST_AMOUNT = Ei2.REST_AMOUNT
               Ei.PACK_AMOUNT = Ei2.PACK_AMOUNT
               Ei.PACK_DATE = Ei2.PACK_DATE
               Ei.TIME_PACK_BEGIN = Ei2.TIME_PACK_BEGIN
               Ei.TIME_PACK_END = Ei2.TIME_PACK_END
               Ei.NOTE = Ei2.NOTE
               Ei.LOCATION_NAME = Ei2.LOCATION_NAME
               If Ei2.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
                  Ei.TX_AMOUNT = Ei2.TX_AMOUNT
               Else
                  Ei.TX_AMOUNT = Ei2.GOOD_AMOUNT * Ei2.WEIGHT_PER_PACK
               End If
               Ei.TOTAL_WEIGHT = 0
               Ei.TOTAL_INCLUDE_PRICE = 0
               Ei.INCLUDE_UNIT_PRICE = 0
               Ei.LINK_ID = Ei2.LINK_ID
               Ei.BIN_NO = Ei2.BIN_NO
               Ei.PALLET_NO = Ei2.PALLET_NO
               Ei.LOCK_NO = Ei2.LOCK_NO
               Ei.HEAD_PACK_NO = Ei2.HEAD_PACK_NO
               Ei.GOOD_AMOUNT = Ei2.GOOD_AMOUNT
               Ei.WEIGHT_PER_PACK = Ei2.WEIGHT_PER_PACK
               Ei.FULL_PALLET_FROM = Ei2.FULL_PALLET_FROM
               Ei.FULL_PALLET_TO = Ei2.FULL_PALLET_TO
               Ei.FULL_UNIT_PER_PALLET = Ei2.FULL_UNIT_PER_PALLET
               Ei.SCRAP_PALLET = Ei2.SCRAP_PALLET
               Ei.SCRAP_UNIT_PER_PALLET = Ei2.SCRAP_UNIT_PER_PALLET
               Ei.AddEditMode = SHOW_EDIT
               
               Set LTD = New CLotDoc
               For Each LTD In Ei2.C_LotDoc
                   Call Ei.C_LotDoc.add(LTD)
                Next LTD
               Set Ei = Nothing
               
               Set Ei3 = Nothing
            End If
      ElseIf Ei2.Flag = "D" Then
      End If
      Set Ei2 = Nothing
   Next Di
End Function
Private Function ServiceRating(Ug As CDoItem, Sf As CSocFeature, Quoatas As Collection, UseFlag As Boolean, UnitType As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim RoundedAmount As Double
Dim Sp As CSystemParam
Dim RateTypeID As Long
Dim TierTypeID As Long
Dim StepTypeID As Long
Static SC As CSoc
Dim FoundFlag As Boolean
Dim TempID As Long
Dim Sum As Double
Dim I As Long
Static PrevSocID As Long
Static PrevDate As Date
Dim TempSoc As CSoc

   RName = "ServiceRating"
   ServiceRating = False
   HasBegin = False

   RateTypeID = RATE_FLAT
   TierTypeID = RATE_TIER
   StepTypeID = RATE_STEP
   
   Set TempSoc = New CSoc
   TempSoc.SOC_ID = Sf.SOC_ID
   Call TempSoc.QueryData(1, m_Rs3, ItemCount)
   If Not m_Rs3.EOF Then
      Call TempSoc.PopulateFromRS(1, m_Rs3)
      If (PrevSocID <> Sf.SOC_ID) Or (PrevDate <> TempSoc.MODIFY_DATE) Then
         Set SC = Nothing
         PrevSocID = Sf.SOC_ID
         PrevDate = TempSoc.MODIFY_DATE
      End If
   End If
   Set TempSoc = Nothing
   
   If SC Is Nothing Then
      Set SC = New CSoc
      SC.SOC_ID = Sf.SOC_ID
      SC.QueryFlag = 1
      Call QuerySoc(SC, m_Rs3, ItemCount, IsOK, ErrorObj)
   
      If m_Rs3.EOF Then
         ErrorObj.LocalErrorMsg = "ไม่พบแพคเกจที่ระบุ"
         ErrorObj.ShowUserError
         
         Exit Function
      End If
   End If
   
   FoundFlag = False
   TempID = Sf.SOC_FEATURE_ID
   For Each Sf In SC.SocFeatures
      If Sf.SOC_FEATURE_ID = TempID Then
         FoundFlag = True
         Exit For
      End If
   Next Sf
'   Set Sc = Nothing

   If Sf.AC_FLAG = "Y" Then
      Ug.AC_AMOUNT = Sf.ACRate.RATE_AMOUNT
   Else
      Ug.AC_AMOUNT = 0
   End If
   
   If Sf.UC_FLAG = "N" Then
      Ug.UC_AMOUNT = 0
   End If
   
   ' Minimum
   Ug.MINIMUM_FLAG = Sf.MINIMUM_FLAG
   If Sf.MINIMUM_FLAG = "Y" Then
      If Sf.MINIMUM_UNIT > Ug.ITEM_AMOUNT Then
         Ug.CALCULATED_AMOUNT = Sf.MINIMUM_UNIT
      Else
         Ug.CALCULATED_AMOUNT = Ug.ITEM_AMOUNT
      End If
   Else
      Ug.CALCULATED_AMOUNT = Ug.ITEM_AMOUNT
   End If
   
   ' Round
   If Sf.ROUNDING_FACTOR > 0 Then
      If (Ug.CALCULATED_AMOUNT Mod Sf.ROUNDING_FACTOR) > 0 Then
         RoundedAmount = (Ug.CALCULATED_AMOUNT \ Sf.ROUNDING_FACTOR + 1) * Sf.ROUNDING_FACTOR
      Else
         RoundedAmount = (Ug.CALCULATED_AMOUNT \ Sf.ROUNDING_FACTOR) * Sf.ROUNDING_FACTOR
      End If
      Ug.CALCULATED_AMOUNT = RoundedAmount
   Else
      RoundedAmount = Ug.CALCULATED_AMOUNT
      Ug.CALCULATED_AMOUNT = RoundedAmount
   End If
   
   Dim StpTVl As CStpTierVol
   Dim Left As Double
   If Sf.RATE_TYPE = RateTypeID Then
      If UnitType = 1 Then
         Ug.UC_AMOUNT = Sf.UCRate.PKG_RATE_AMOUNT * Ug.CALCULATED_AMOUNT
         Ug.LOGISTIC_PRICE = Sf.UCRate.LOG_RATE_AMOUNT
      Else
         Ug.UC_AMOUNT = Sf.UCRate.RATE_AMOUNT * Ug.CALCULATED_AMOUNT
         Ug.LOGISTIC_PRICE = Sf.UCRate.LOG_RATE_AMOUNT
      End If
      If Ug.ITEM_AMOUNT <> 0 Then
         Ug.AVG_PRICE = (Ug.UC_AMOUNT + Ug.AC_AMOUNT) / Ug.ITEM_AMOUNT
      Else
         Ug.AVG_PRICE = 0
      End If
   ElseIf Sf.RATE_TYPE = TierTypeID Then
      Left = Ug.CALCULATED_AMOUNT
      For Each StpTVl In Sf.StepTierVol
         If Left <= StpTVl.Width Then
            Exit For
         Else
            Left = Left - StpTVl.Width
         End If
      Next StpTVl
      Ug.UC_AMOUNT = StpTVl.RATE_AMOUNT * Ug.CALCULATED_AMOUNT
      If Ug.ITEM_AMOUNT <> 0 Then
         Ug.AVG_PRICE = (Ug.UC_AMOUNT + Ug.AC_AMOUNT) / Ug.ITEM_AMOUNT
      Else
         Ug.AVG_PRICE = 0
      End If
   ElseIf Sf.RATE_TYPE = StepTypeID Then
      Left = Ug.CALCULATED_AMOUNT
      Sum = 0
      For Each StpTVl In Sf.StepTierVol
         If Left <= StpTVl.Width Then
            Sum = Sum + StpTVl.RATE_AMOUNT * Left
            Left = Left - StpTVl.Width
            Exit For
         Else
            Sum = Sum + StpTVl.RATE_AMOUNT * StpTVl.Width
            Left = Left - StpTVl.Width
         End If
      Next StpTVl
      Ug.UC_AMOUNT = Sum + Minus2Zero(Left) * StpTVl.RATE_AMOUNT
      If Ug.ITEM_AMOUNT <> 0 Then
         Ug.AVG_PRICE = (Ug.UC_AMOUNT + Ug.AC_AMOUNT) / Ug.ITEM_AMOUNT
      Else
         Ug.AVG_PRICE = 0
      End If
   End If

   ServiceRating = True
   Exit Function
   
ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   ServiceRating = True
End Function

Public Function ServiceGuiding(Ug As CDoItem, SocLevel As String, IsOK As Boolean, ErrorObj As clsErrorLog) As CSocFeature
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim Sf As CSocFeature
Dim TempRs As ADODB.Recordset
Dim ServerDate As String
Dim FoundFlag As Boolean

   RName = "ServiceGuiding"
   Set ServiceGuiding = Nothing
   HasBegin = False

   Call glbDatabaseMngr.GetServerDateTime(ServerDate, ErrorObj)
   
   FoundFlag = False
   Set Sf = New CSocFeature
   Set TempRs = New ADODB.Recordset
   
   If SocLevel <> "Y" Then
      Sf.PART_ITEM_ID = Ug.PART_ITEM_ID
      Sf.FEATURE_ID = Ug.FEATURE_ID
      Sf.SUBSCRIBER_ID = Ug.SUBSCRIBER_ID
      Sf.ACCOUNT_ID = Ug.ACCOUNT_ID
      Call Sf.QueryData3(TempRs, ItemCount)
      While (Not TempRs.EOF) And (Not FoundFlag)
         If (Ug.FEATURE_ID = NVLI(TempRs("FEATURE_ID"), -1)) And _
            (ServerDate >= NVLS(TempRs("EFFECTIVE_DATE"), "")) And _
            (ServerDate <= NVLS(TempRs("EXPIRE_DATE"), "")) Then
            
            Call Sf.PopulateFromRS3(TempRs)
            FoundFlag = True
         ElseIf (Ug.PART_ITEM_ID = NVLI(TempRs("PART_ITEM_ID"), -1)) And _
            (ServerDate >= NVLS(TempRs("EFFECTIVE_DATE"), "")) And _
            (ServerDate <= NVLS(TempRs("EXPIRE_DATE"), "")) Then
            
            Call Sf.PopulateFromRS3(TempRs)
            FoundFlag = True
         End If
         
         TempRs.MoveNext
      Wend
   End If
   
   If FoundFlag Then
      Set ServiceGuiding = Sf
      Ug.SOC_ID = Sf.SOC_ID
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      Exit Function
   Else
      Set ServiceGuiding = Nothing
   End If
   
   If SocLevel <> "N" Then
      FoundFlag = False
      Sf.FEATURE_ID = Ug.FEATURE_ID
      Sf.PART_ITEM_ID = Ug.PART_ITEM_ID
      Call Sf.QueryData4(TempRs, ItemCount)
      While (Not TempRs.EOF) And (Not FoundFlag)
         If (Ug.FEATURE_ID = NVLI(TempRs("FEATURE_ID"), -1)) Then
            
            Call Sf.PopulateFromRS3(TempRs)
            FoundFlag = True
         ElseIf (Ug.PART_ITEM_ID = NVLI(TempRs("PART_ITEM_ID"), -1)) Then
            
            Call Sf.PopulateFromRS3(TempRs)
            FoundFlag = True
         End If
         TempRs.MoveNext
      Wend
   End If
   
   If FoundFlag Then
      Ug.SOC_ID = Sf.SOC_ID
      Set ServiceGuiding = Sf
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      Exit Function
   Else
      Set ServiceGuiding = Nothing
   End If
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set Sf = Nothing
   
   Exit Function
   
ErrorHandler:
   IsOK = False

   Set ServiceGuiding = Nothing
End Function

Public Function CalculatePrice(Ug As CDoItem, Quoatas As Collection, UseFlag As Boolean, UnitType As Long, SocLevel As String, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
Dim Sf As CSocFeature

   Set Sf = ServiceGuiding(Ug, SocLevel, IsOK, ErrorObj)
   If Sf Is Nothing Then
'      ErrorObj.LocalErrorMsg = "ไม่พบข้อมูลแพคเกจสำหรับสินค้าและบริการนี้ "
'      ErrorObj.ShowUserError
   Else
      If Not ServiceRating(Ug, Sf, Quoatas, UseFlag, UnitType, IsOK, ErrorObj) Then
      End If
   End If
   CalculatePrice = Not (Sf Is Nothing)
End Function

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub

Public Function QueryEmpReceivable(Ua As CEmpReceivable, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryEmpReceivable"
   QueryEmpReceivable = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
   
   
   QueryEmpReceivable = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmpReceivable = False
End Function

Public Function AddEditEmpReceivable(Ug As CEmpReceivable, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEmpReceivable"
   AddEditEmpReceivable = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditEmpReceivable = True
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
   AddEditEmpReceivable = False
End Function

Public Function DeleteEmpReceivable(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CEmpReceivable

   RName = "DeleteEmpReceivable"
   DeleteEmpReceivable = False
   HasBegin = False

   Set B = New CEmpReceivable


   HasBegin = True
   
   B.EMP_RECEIVABLE_ID = UID
   Call B.QueryData(m_Rs1, ItemCount)
   If Not m_Rs1.EOF Then
      Call B.PopulateFromRS(m_Rs1)
   End If
   
   B.EMP_RECEIVABLE_ID = UID
   Call B.DeleteData

   Set B = Nothing


   HasBegin = False
   IsOK = True

   DeleteEmpReceivable = True
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

   DeleteEmpReceivable = False
End Function


Public Function QueryEmployeeMoney(Ua As CEmployee, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   RName = "QueryEmployeeMoney"
   QueryEmployeeMoney = False

   IsOK = True
   Call Ua.QueryDataMoney(Rs, ItemCount)
   QueryEmployeeMoney = True
   
   
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmployeeMoney = False
End Function
Public Function QuerySliptSalary(Ua As CSliptSalary, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
   RName = "QuerySliptSalary"
   QuerySliptSalary = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)
      If Ua.QueryFlag <> 1 Then
   Dim m_sliptSalary As CSliptSalary
   Set m_sliptSalary = New CSliptSalary
   Call m_sliptSalary.PopulateFromRS(Rs)
   
   If ItemCount <> 0 Then
         Dim SA As CSliptAdd
      Set SA = New CSliptAdd
      SA.SLIPT_SALARY_ID = m_sliptSalary.SLIPT_SALARY_ID
      Call SA.QueryData(m_Rs1, iCount)
      Set SA = Nothing
      
      Set Ua.SliptAdd = Nothing
      Set Ua.SliptAdd = New Collection
      While Not m_Rs1.EOF
         Set SA = New CSliptAdd
         Call SA.PopulateFromRS(m_Rs1)
      
         SA.Flag = "I"
         Call Ua.SliptAdd.add(SA)
         Set SA = Nothing
         m_Rs1.MoveNext
      Wend

 Dim SB As CSliptSub
      Set SB = New CSliptSub
      SB.SLIPT_SALARY_ID = m_sliptSalary.SLIPT_SALARY_ID
      Call SB.QueryData(m_Rs1, iCount)
      Set SB = Nothing
      
      Set Ua.SliptSub = Nothing
      Set Ua.SliptSub = New Collection
      While Not m_Rs1.EOF
         Set SB = New CSliptSub
         Call SB.PopulateFromRS(m_Rs1)
      
         SB.Flag = "I"
         Call Ua.SliptSub.add(SB)
         Set SB = Nothing
         m_Rs1.MoveNext
      Wend
Else
frmSliptSalary.cmdAdd.Enabled = True
frmSliptSalary.cmdEdit.Enabled = True
frmSliptSalary.cmdDelete.Enabled = True

End If
End If
   QuerySliptSalary = True
   
   
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySliptSalary = False
End Function
Public Function QuerySliptSalary1(Ua As CSliptSalary, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
   RName = "QuerySliptSalary1"
   QuerySliptSalary1 = False

   IsOK = True
   Call Ua.QueryData1(Rs, ItemCount)
      QuerySliptSalary1 = True
   
   
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySliptSalary1 = False
End Function
Public Function QuerySliptSalary1_1(Ua As CSliptSalary, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
   RName = "QuerySliptSalary1_1"
   QuerySliptSalary1_1 = False

   IsOK = True
   Call Ua.QueryData2(Rs, ItemCount)
      QuerySliptSalary1_1 = True
   
   
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySliptSalary1_1 = False
End Function

Public Function QuerySliptSalary2(Ua As CSliptSalary, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
   RName = "QuerySliptSalary2"
   QuerySliptSalary2 = False

   IsOK = True
   
        Dim SA As CSliptAdd
      Set SA = New CSliptAdd
      SA.SLIPT_SALARY_ID = Ua.SLIPT_SALARY_ID
      Call SA.QueryData(m_Rs1, iCount)
      Set SA = Nothing
      
      Set Ua.SliptAdd = Nothing
      Set Ua.SliptAdd = New Collection
      While Not m_Rs1.EOF
         Set SA = New CSliptAdd
         Call SA.PopulateFromRS(m_Rs1)
      
         SA.Flag = "I"
         Call Ua.SliptAdd.add(SA)
         Set SA = Nothing
         m_Rs1.MoveNext
      Wend

 Dim SB As CSliptSub
      Set SB = New CSliptSub
      SB.SLIPT_SALARY_ID = Ua.SLIPT_SALARY_ID
      Call SB.QueryData(m_Rs1, iCount)
      Set SB = Nothing
      
      Set Ua.SliptSub = Nothing
      Set Ua.SliptSub = New Collection
      While Not m_Rs1.EOF
         Set SB = New CSliptSub
         Call SB.PopulateFromRS(m_Rs1)
      
         SB.Flag = "I"
         Call Ua.SliptSub.add(SB)
         Set SB = Nothing
         m_Rs1.MoveNext
      Wend

   QuerySliptSalary2 = True
   
   
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QuerySliptSalary2 = False
End Function
Public Function AddEditSliptSalary(Ug As CSliptSalary, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSliptSalary"
   AddEditSliptSalary = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim SA As CSliptAdd
   For Each SA In Ug.SliptAdd
      If SA.Flag = "A" Then
      SA.SLIPT_SALARY_ID = Ug.SLIPT_SALARY_ID
         SA.AddEditMode = SHOW_ADD
         Call SA.AddEditData

      ElseIf SA.Flag = "E" Then
         SA.AddEditMode = SHOW_EDIT
         Call SA.AddEditData
      ElseIf SA.Flag = "D" Then
         Call SA.DeleteData
      End If
   Next SA
   
   Dim SB As CSliptSub
   For Each SB In Ug.SliptSub
      If SB.Flag = "A" Then
      SB.SLIPT_SALARY_ID = Ug.SLIPT_SALARY_ID
         SB.AddEditMode = SHOW_ADD
         Call SB.AddEditData

      ElseIf SB.Flag = "E" Then
         SB.AddEditMode = SHOW_EDIT
         Call SB.AddEditData
      ElseIf SB.Flag = "D" Then
         Call SB.DeleteData
      End If
   Next SB
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditSliptSalary = True
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
   AddEditSliptSalary = False
End Function

Public Function DeleteSliptSalary(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CSliptSalary

   RName = "DeleteSliptSalary"
   DeleteSliptSalary = False
   HasBegin = False

   Set B = New CSliptSalary

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.SLIPT_SALARY_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSliptSalary = True
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

   DeleteSliptSalary = True
End Function

Public Function AddEditEmployeeMoney(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEmployeeMoney"
   AddEditEmployeeMoney = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditDataMoney
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEmployeeMoney = True
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
   AddEditEmployeeMoney = False
End Function

Public Function QueryCurrency(Ua As CCurrency, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

Dim Qp As CCurrency

   RName = "QueryCurrency"
   QueryCurrency = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCurrency = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCurrency = False
End Function

Public Function DeleteCurrency(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CCurrency

   RName = "DeleteCurrency"
   DeleteCurrency = False
   HasBegin = False

   Set B = New CCurrency

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.CURRENCY_EXC_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCurrency = True
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

   DeleteCurrency = True
End Function

Public Function AddEditCurrency(Ug As CCurrency, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCurrency"
   AddEditCurrency = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCurrency = True
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
   AddEditCurrency = False
End Function

Public Function AddEditPayment(Ug As CPayment, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object

   RName = "AddEditPayment"
   AddEditPayment = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CPaymentItem
   For Each CstName In Ug.PaymentItems
      If CstName.Flag = "A" Then
         CstName.PAYMENT_ID = Ug.PAYMENT_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
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

   AddEditPayment = True
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
   AddEditPayment = False
End Function

Public Function AddEditQuantityExtract(Ug As CQuantityExtract, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object

   RName = "AddEditQuantityExtract"
   AddEditQuantityExtract = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CExtractItem
   For Each CstName In Ug.ExtractItems
      If CstName.Flag = "A" Then
         CstName.QUANTITY_EXTRACT_ID = Ug.QUANTITY_EXTRACT_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
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

   AddEditQuantityExtract = True
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
   AddEditQuantityExtract = False
End Function

Public Function QueryPayment(Ua As CPayment, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryPayment"
   QueryPayment = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CPaymentItem
      
      Set Yw = New CPaymentItem
      Yw.PAYMENT_ITEM_ID = -1
      Yw.PAYMENT_ID = Ua.PAYMENT_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.PaymentItems = Nothing
      Set Ua.PaymentItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CPaymentItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.PaymentItems.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryPayment = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPayment = False
End Function

Public Function QueryQuantityExtract(Ua As CQuantityExtract, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryQuantityExtract"
   QueryQuantityExtract = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CExtractItem
      
      Set Yw = New CExtractItem
      Yw.EXTRACT_ITEM_ID = -1
      Yw.QUANTITY_EXTRACT_ID = Ua.QUANTITY_EXTRACT_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.ExtractItems = Nothing
      Set Ua.ExtractItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CExtractItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.ExtractItems.add(Yw, Trim(str(Yw.PART_ITEM_ID)))
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryQuantityExtract = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryQuantityExtract = False
End Function

Public Function DeletePayment(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CPayment

   RName = "DeletePayment"
   DeletePayment = False
   HasBegin = False

   Set B = New CPayment

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PAYMENT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePayment = True
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

   DeletePayment = True
End Function

Public Function QueryCurrencyEx(Ua As CCurrencyEx, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

Dim Qp As CCurrencyEx

   RName = "QueryCurrencyEx"
   QueryCurrencyEx = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCurrencyEx = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCurrencyEx = False
End Function

Public Function DeleteCurrencyEx(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CCurrencyEx

   RName = "DeleteCurrencyEx"
   DeleteCurrencyEx = False
   HasBegin = False

   Set B = New CCurrencyEx

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.CURRENCY_EX_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCurrencyEx = True
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

   DeleteCurrencyEx = True
End Function

Public Function AddEditCurrencyEx(Ug As CCurrencyEx, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCurrencyEx"
   AddEditCurrencyEx = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCurrencyEx = True
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
   AddEditCurrencyEx = False
End Function
Public Sub CopyBalanceAccum(Src As Collection, Dst As Collection)
Dim II As CLotItem
Dim Ba As CBalanceAccum
   
   For Each Ba In Src
'If Ba.LOCATION_ID = 106 And Ba.PART_ITEM_ID = 123 Then
''Debug.Print
'End If
      Set II = New CLotItem
      II.LOCATION_ID = Ba.LOCATION_ID
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
      'If Ba.TOTAL_INCLUDE_PRICE > 1000000 Then
         II.TOTAL_INCLUDE_PRICE = Ba.TOTAL_INCLUDE_PRICE
      'End If
      II.CURRENT_AMOUNT = Ba.BALANCE_AMOUNT
      II.NEW_PRICE = Ba.AVG_PRICE   'MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = Ba.AVG_PRICE
      II.TX_TYPE = "I"
      II.DOCUMENT_DATE = Ba.DOCUMENT_DATE
      II.NEW_AMOUNT = Ba.BALANCE_AMOUNT
      Call Dst.add(II, Ba.LOCATION_ID & "-" & Ba.PART_ITEM_ID)
      Set II = Nothing
   Next Ba
End Sub

Public Sub CopyBalanceAccumByLocation(Src As Collection, Dst As Collection)
Dim II As CLotItem
Dim Ba As CBalanceAccum
   
   For Each Ba In Src
      Set II = New CLotItem
      II.LOCATION_ID = Ba.LOCATION_ID
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
      II.TOTAL_INCLUDE_PRICE = Ba.TOTAL_INCLUDE_PRICE
      II.CURRENT_AMOUNT = Ba.BALANCE_AMOUNT
      II.NEW_PRICE = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = Ba.AVG_PRICE
      II.TX_TYPE = "I"
      II.DOCUMENT_DATE = Ba.DOCUMENT_DATE
      II.NEW_AMOUNT = Ba.BALANCE_AMOUNT
      Call Dst.add(II, Trim(str(Ba.LOCATION_ID)))
      Set II = Nothing
   Next Ba
End Sub

Public Sub CopyBalanceAccumByPartItem(Src As Collection, Dst As Collection)
Dim II As CLotItem
Dim Ba As CBalanceAccum
   
   For Each Ba In Src
      Set II = New CLotItem
      II.LOCATION_ID = Ba.LOCATION_ID
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
      II.TOTAL_INCLUDE_PRICE = Ba.TOTAL_INCLUDE_PRICE
      II.CURRENT_AMOUNT = Ba.BALANCE_AMOUNT
      II.NEW_PRICE = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = Ba.AVG_PRICE
      II.TX_TYPE = "I"
      II.DOCUMENT_DATE = Ba.DOCUMENT_DATE
      II.NEW_AMOUNT = Ba.BALANCE_AMOUNT
      Call Dst.add(II, Trim(str(Ba.PART_ITEM_ID)))
      Set II = Nothing
   Next Ba
End Sub
Public Function PatchDocumentPriority()
Dim Lc As CDocumentPriority
   
   Set Lc = New CDocumentPriority
   
   'Call glbDaily.StartTransaction
   
   Call Lc.ClearData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 1
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 1
   Lc.PRIORITY2 = 1
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 2
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 10
   Lc.PRIORITY2 = 10
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 2
   Lc.PRIORITY2 = 2
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 11
   Lc.PRIORITY2 = 11
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 11
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 3
   Lc.PRIORITY2 = 3
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 11
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 12
   Lc.PRIORITY2 = 12
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 19
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 4
   Lc.PRIORITY2 = 4
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 5
   Lc.PRIORITY2 = 5
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 13
   Lc.PRIORITY2 = 13
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 10
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 14
   Lc.PRIORITY2 = 14
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 21
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 15
   Lc.PRIORITY2 = 15
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 6
   Lc.PRIORITY2 = 6
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 16
   Lc.PRIORITY2 = 16
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 23
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 7
   Lc.PRIORITY2 = 7
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 20
   Lc.PRIORITY2 = 20
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 21
   Lc.PRIORITY2 = 21
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 18
   Lc.PRIORITY2 = 18
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 19
   Lc.PRIORITY2 = 19
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 31
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 9
   Lc.PRIORITY2 = 9
   Lc.Area = 1
   Call Lc.AddEditData
   
   'Call glbDaily.CommitTransaction
   Set Lc = Nothing
End Function

Private Function IsExist(TempID As Long, TempCol As Collection) As Boolean
Dim Hgi As CHGroupItem

   For Each Hgi In TempCol
      If Hgi.LOCATION_ID = TempID Then
         IsExist = True
         Exit Function
      End If
   Next Hgi
   
   IsExist = False
End Function

Public Function DO2Payment(BD As CBillingDoc, Pm As CPayment) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CPaymentItem
Dim TempDI As CDoItem
Dim TotalPrice As Double
Dim RevenueTotalPrice As Double
Dim Ri As CReceiptItem

   If (BD.DOCUMENT_TYPE = 2) And (BD.RECEIPT_TYPE = 1) Then 'ใบเสร็จ
      TotalPrice = 0
      For Each TempDI In BD.DoItems
         If TempDI.Flag <> "D" Then
            TotalPrice = TotalPrice + TempDI.TOTAL_PRICE
         End If
      Next TempDI
      
'      RevenueTotalPrice = 0
'      For Each TempDi In Bd.Revenues
'         If TempDi.Flag <> "D" Then
'            RevenueTotalPrice = RevenueTotalPrice + TempDi.TOTAL_PRICE
'         End If
'      Next TempDi
   ElseIf (BD.DOCUMENT_TYPE = 2) And (BD.RECEIPT_TYPE = 3) Then 'ใบเสร็จรับชำระ
      TotalPrice = 0
      RevenueTotalPrice = 0
      For Each Ri In BD.ReceiptItems
         If Ri.Flag <> "D" Then
            TotalPrice = TotalPrice + Ri.PAID_AMOUNT
         End If
      Next Ri
   End If
   
   Set Pm = Nothing
   Set Pm = New CPayment

   If BD.PAYMENT_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Pm.PAYMENT_ID = BD.PAYMENT_ID
      Pm.QueryFlag = 1
      Call glbDaily.QueryPayment(Pm, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Pm.AddEditMode = SHOW_EDIT
      Set Di = Pm.PaymentItems(1)  'รับประกันว่าต้องมี
      Di.Flag = "E"
   Else
      Pm.AddEditMode = SHOW_ADD
      Set Di = New CPaymentItem
      Di.Flag = "A"
      Call Pm.PaymentItems.add(Di)
   End If

   If BD.PAYMENT_TYPE = BANKTRF_PMT Then
      Pm.PAYMENT_DATE = BD.DOCUMENT_DATE
      Pm.PAYMENT_NO = BD.DOCUMENT_NO
      Pm.COMMIT_FLAG = BD.COMMIT_FLAG
      Pm.TO_BANK_BRANCH = BD.BBRANCH_ID
      Pm.TO_BANK_ID = BD.BANK_ID
      Pm.ACCOUNT_NO = BD.CHECK_NO
      Pm.TX_TYPE = "I" 'In, มีเงินเข้ามา
      Pm.INTERNAL_FLAG = "Y"
      Pm.TOTAL_AMOUNT = TotalPrice + RevenueTotalPrice
      
      Di.PAYMENT_ID = Pm.PAYMENT_ID
      Di.PAYMENT_TYPE = BD.PAYMENT_TYPE
      Di.BANK_BRANCH = -1
      Di.BANK_ID = -1
      Di.BANK_NOTE = BD.BANK_NOTE
      Di.CHECK_NO = ""
      Di.PAY_AMOUNT = TotalPrice + RevenueTotalPrice
   Else
      Pm.PAYMENT_DATE = BD.DOCUMENT_DATE
      Pm.PAYMENT_NO = BD.DOCUMENT_NO
      Pm.COMMIT_FLAG = BD.COMMIT_FLAG
      Pm.TO_BANK_BRANCH = -1
      Pm.TO_BANK_ID = -1
      Pm.ACCOUNT_NO = ""
      Pm.TX_TYPE = "I" 'In, มีเงินเข้ามา
      Pm.INTERNAL_FLAG = "Y"
      Pm.TOTAL_AMOUNT = TotalPrice + RevenueTotalPrice
      
      Di.PAYMENT_ID = Pm.PAYMENT_ID
      Di.PAYMENT_TYPE = BD.PAYMENT_TYPE
      Di.BANK_BRANCH = BD.BBRANCH_ID
      Di.BANK_ID = BD.BANK_ID
      Di.BANK_NOTE = BD.BANK_NOTE
      Di.CHECK_NO = BD.CHECK_NO
      Di.PAY_AMOUNT = TotalPrice + RevenueTotalPrice
   End If
End Function

Private Function IsDiscountMatch(Di As CDoItem, Bdc As CBillingDiscount) As Boolean
   IsDiscountMatch = False
   
   If Bdc.DISCOUNT_TYPE = 1 Then 'สินค้า
      If (Di.PART_ITEM_ID > 0) And (Di.PART_ITEM_ID = Bdc.PART_ITEM_ID) Then
         IsDiscountMatch = True
      ElseIf (Di.PART_ITEM_ID > 0) And (Bdc.PART_ITEM_ID <= 0) Then
         IsDiscountMatch = True
      End If
   ElseIf Bdc.DISCOUNT_TYPE = 2 Then 'บริการ
      If (Di.FEATURE_ID > 0) And (Di.FEATURE_ID = Bdc.FEATURE_ID) Then
         IsDiscountMatch = True
      ElseIf (Di.FEATURE_ID > 0) And (Bdc.FEATURE_ID <= 0) Then
         IsDiscountMatch = True
      End If
   ElseIf Bdc.DISCOUNT_TYPE = 3 Then 'ทั้งหมด
      IsDiscountMatch = True
   End If
End Function

Public Sub GenerateExtraDiscount(BD As CBillingDoc)
Dim Di As CDoItem
Dim Bdc As CBillingDiscount
Dim SumWeight As Double
Dim SumPack As Double
Dim SumPrice As Double
Dim ModifyFlag As Boolean
Dim SumBillDiscount As Double

   SumWeight = 0
   SumPack = 0
   SumPrice = 0
   SumBillDiscount = 0
   
   'คำนวณใหม่เลย
   For Each Bdc In BD.BillingDiscounts
      If Bdc.Flag <> "D" Then
         Bdc.DISCOUNT_AMOUNT = 0
      End If
   Next Bdc
   
   For Each Di In BD.DoItems
      If Di.Flag <> "D" Then
         If Di.Flag <> "A" Then
            Di.Flag = "E"
         End If
         Di.EXTRA_DISCOUNT = 0
         ModifyFlag = False
         
         For Each Bdc In BD.BillingDiscounts
            If Bdc.Flag <> "D" Then
               If IsDiscountMatch(Di, Bdc) Then
                  If Bdc.DSCN_PER_PACK > 0 Then 'ลดต่อถุง
                     Di.EXTRA_DISCOUNT = Di.EXTRA_DISCOUNT + Bdc.DSCN_PER_PACK * Di.PACK_AMOUNT
                  ElseIf Bdc.DSCN_PER_WEIGHT > 0 Then 'ลดต่อน้ำหนัก
                     Di.EXTRA_DISCOUNT = Di.EXTRA_DISCOUNT + Bdc.DSCN_PER_WEIGHT * Di.ITEM_AMOUNT
                  ElseIf Bdc.DSCN_PER_MONEY > 0 Then 'ลดต่อจำนวนเงิน
                     Di.EXTRA_DISCOUNT = Di.EXTRA_DISCOUNT + (Bdc.DSCN_PER_MONEY * Di.TOTAL_PRICE) / 100
                  End If
                  ModifyFlag = True
               End If
               'แปลงเป็น 2 ตำแหน่ง
               Di.EXTRA_DISCOUNT = Val(Format(Di.EXTRA_DISCOUNT, "0.00"))
               Bdc.DISCOUNT_AMOUNT = Bdc.DISCOUNT_AMOUNT + Di.EXTRA_DISCOUNT
               
'               SumBillDiscount = SumBillDiscount + Bdc.DISCOUNT_AMOUNT
               If Bdc.Flag <> "A" Then
                  Bdc.Flag = "E"
               End If
            End If
         Next Bdc
                           
      End If
   Next Di
   
   For Each Bdc In BD.BillingDiscounts
      If Bdc.Flag <> "D" Then
         SumBillDiscount = SumBillDiscount + Bdc.DISCOUNT_AMOUNT
      End If
   Next Bdc
   BD.CD_AMOUNT = SumBillDiscount
End Sub
Public Function DeleteCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCashDoc"
   DeleteCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCashDoc = True
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

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCashDoc = True
End Function
Public Function DeleteChequeDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCheq  As CChequeDoc
   RName = "DeleteChequeDoc"
  DeleteChequeDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

 Set cCheq = New CChequeDoc
   
   cCheq.CHEQUE_DOC_ID = UID
   Call cCheq.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteChequeDoc = True
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

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteChequeDoc = True
End Function
Public Function DeleteReceiptChequeDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCheq  As CReceiptChequeDoc
   RName = "DeleteReceiptChequeDoc"
  DeleteReceiptChequeDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

 Set cCheq = New CReceiptChequeDoc
   
   cCheq.RECEIPT_CHEQUE_DOC_ID = UID
   Call cCheq.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteReceiptChequeDoc = True
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

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteReceiptChequeDoc = True
End Function

Public Function QueryCashDoc(Ua As CCashDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCashDoc"
   QueryCashDoc = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
'      Dim Ji As CJournalItem
'      Set Ji = New CJournalItem
'      Call Ji.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
'      Call Ji.QueryData(1, m_Rs1, iCount)
'      Set Ji = Nothing
'
'      Set Ua.JournalItems = Nothing
'      Set Ua.JournalItems = New Collection
'
'      While Not m_Rs1.EOF
'         Set Ji = New CJournalItem
'         Call Ji.PopulateFromRS(1, m_Rs1)
'         Call Ua.JournalItems.add(Ji)
'         Set Ji = Nothing
'
'         m_Rs1.MoveNext
'      Wend
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.CashTranItems = Nothing
      Set Ua.CashTranItems = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHECK_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.CashTranItems.add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim CDP As CCashDocPost
      Set CDP = New CCashDocPost
      Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing
      
      Set Ua.PostItems = Nothing
      Set Ua.PostItems = New Collection
      
      While Not m_Rs1.EOF
         Set CDP = New CCashDocPost
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         If CDP.GetFieldValue("BILLING_DOC_ID") > 0 Then
            Dim Cp2D As CBillingDoc
            Set Cp2D = New CBillingDoc
            Cp2D.BILLING_DOC_ID = CDP.GetFieldValue("BILLING_DOC_ID")
            Call Cp2D.QueryData(1, m_Rs2, iCount)
            Set Cp2D = Nothing
            
            Set CDP.Post2BD = Nothing
            Set CDP.Post2BD = New Collection
            
            While Not m_Rs2.EOF
               Set Cp2D = New CBillingDoc
               Call Cp2D.PopulateFromRS(1, m_Rs2)
                  
               Call CDP.Post2BD.add(Cp2D)
               Set Cp2D = Nothing
               
               m_Rs2.MoveNext
            Wend
         End If
         
         Call Ua.PostItems.add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend

   End If
   
   QueryCashDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCashDoc = False
End Function
Public Function QueryCashDocEx(Ua As CCashDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
   
   RName = "QueryCashDoc"
   QueryCashDocEx = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 And ItemCount > 0 Then
            
      Call Ua.PopulateFromRS(1, Rs)
      
      Dim CDP As CCashDocPost
      Set CDP = New CCashDocPost
      Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing
      
      Set Ua.PostItems = Nothing
      Set Ua.PostItems = New Collection
      
      While Not m_Rs1.EOF
         Set CDP = New CCashDocPost
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         If CDP.GetFieldValue("BILLING_DOC_ID") > 0 Then
            Dim Cp2D As CBillingDoc
            Set Cp2D = New CBillingDoc
            Cp2D.BILLING_DOC_ID = CDP.GetFieldValue("BILLING_DOC_ID")
            Call Cp2D.QueryData(1, m_Rs2, iCount)
            Set Cp2D = Nothing
            
            Set CDP.Post2BD = Nothing
            Set CDP.Post2BD = New Collection
            
            While Not m_Rs2.EOF
               Set Cp2D = New CBillingDoc
               Call Cp2D.PopulateFromRS(1, m_Rs2)
                  
               Call CDP.Post2BD.add(Cp2D)
               Set Cp2D = Nothing
               
               m_Rs2.MoveNext
            Wend
         End If
         
         Call Ua.PostItems.add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend

   End If
   
   QueryCashDocEx = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCashDocEx = False
End Function
Public Function QueryChequeDoc(Ua As CChequeDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
   
   RName = "QueryChequeDoc"
   QueryChequeDoc = False
   
   IsOK = True
   
   Call Ua.QueryData(8, Rs, ItemCount)
    If Ua.QueryFlag = 1 Then
    Dim RCD As CReceiptChequeDoc
      
      Set RCD = New CReceiptChequeDoc
      RCD.RECEIPT_CHEQUE_DOC_ID = -1
     RCD.CHEQUE_DOC_ID = Ua.CHEQUE_DOC_ID
      RCD.OrderBy = -1
      RCD.OrderType = 1
      Call RCD.QueryData(1, m_Rs1, iCount)
      Set RCD = Nothing
      
      Set Ua.ChequeDoc = Nothing
      Set Ua.ChequeDoc = New Collection
      While Not m_Rs1.EOF
         Set RCD = New CReceiptChequeDoc
         Call RCD.PopulateFromRS(1, m_Rs1)
                  
         RCD.Flag = "I"
         Call Ua.ChequeDoc.add(RCD)
         
         Set RCD = Nothing
         m_Rs1.MoveNext
      Wend
    
    
    '---------------ข้อมูลธนาคารของเช็ค ---------
     Dim BBC As CChequeDoc
      
      Set BBC = New CChequeDoc
      BBC.RECEIPT_CHEQUE_DOC_ID = -1
     BBC.CHEQUE_DOC_ID = Ua.CHEQUE_DOC_ID
'      BBC.OrderBy = -1
'      BBC.OrderType = 1
      Call BBC.QueryData(8, m_Rs1, iCount)
      Set BBC = Nothing
      
      Set Ua.BankInfo = Nothing
      Set Ua.BankInfo = New Collection
      While Not m_Rs1.EOF
         Set BBC = New CChequeDoc
         Call BBC.PopulateFromRS(8, m_Rs1)
                  
         BBC.Flag = "I"
         Call Ua.BankInfo.add(BBC)
         
         Set BBC = Nothing
         m_Rs1.MoveNext
      Wend
    
    
    
    
    End If


   QueryChequeDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryChequeDoc = False
End Function
Public Function QueryChequeDoc2(Ua As CChequeDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   
   RName = "QueryChequeDoc2"
   QueryChequeDoc2 = False
   
   IsOK = True
   
   Call Ua.QueryData(3, Rs, ItemCount)
    If Ua.QueryFlag = 1 Then
    Dim RCD As CReceiptChequeDoc

      Set RCD = New CReceiptChequeDoc
      RCD.RECEIPT_CHEQUE_DOC_ID = -1
     RCD.CHEQUE_DOC_ID = Ua.CHEQUE_DOC_ID
      RCD.OrderBy = 1
      RCD.OrderType = 1
      Call RCD.QueryData(1, m_Rs1, iCount)
      Set RCD = Nothing

      Set Ua.ChequeDoc = Nothing
      Set Ua.ChequeDoc = New Collection
      While Not m_Rs1.EOF
         Set RCD = New CReceiptChequeDoc
         Call RCD.PopulateFromRS(1, m_Rs1)

         RCD.Flag = "I"
         Call Ua.ChequeDoc.add(RCD)

         Set RCD = Nothing
         m_Rs1.MoveNext
      Wend
    
    
    
    End If


   QueryChequeDoc2 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryChequeDoc2 = False
End Function

Public Function QueryReceiptChequeDoc(Ua As CReceiptChequeDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
   
   RName = "QueryReceiptChequeDoc"
   QueryReceiptChequeDoc = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
    If Ua.QueryFlag = 1 Then
      Dim Yw As CBillingDoc
      
      Set Yw = New CBillingDoc
      'Yw.DO_ITEM_ID = -1
      Yw.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.ReceiptCheqDocItems = Nothing
      Set Ua.ReceiptCheqDocItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CBillingDoc
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.ReceiptCheqDocItems.add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
    End If



   QueryReceiptChequeDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryReceiptChequeDoc = False
End Function
Public Function AddEditCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim DocType As CASH_DOC_TYPE

   RName = "AddEditCashDoc"
   AddEditCashDoc = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

'   Dim Ji As CJournalItem
'   For Each Ji In Ua.JournalItems
'      Call Ji.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
'      If Ji.Flag = "A" Then
'         Ji.ShowMode = SHOW_ADD
'         Call Ji.AddEditData
'      ElseIf Ji.Flag = "E" Then
'         Ji.ShowMode = SHOW_EDIT
'         Call Ji.AddEditData
'      ElseIf Ji.Flag = "D" Then
'         Call Ji.DeleteData
'      End If
'   Next Ji

   Dim Ct As CCashTran
   Dim PaymentType As Long
   Dim OldPaymentType As Long
   
   DocType = Ua.GetFieldValue("DOCUMENT_TYPE")
   
   For Each Ct In Ua.CashTranItems
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.SetFieldValue("TX_NO", Ua.GetFieldValue("DOCUMENT_NO"))
      Call Ct.SetFieldValue("TX_DATE", Ua.GetFieldValue("DOCUMENT_DATE"))
      Call Ct.SetFieldValue("CUSTOMER_ID", Ua.GetFieldValue("CUSTOMER_ID"))
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If DocType = CASH_DEPOSIT Then 'ฝากเช็คที่มีอยู่ในระบบแล้ว
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "Y")
            Call Ct.Cheque.SetFieldValue("POST_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then 'เช็ค
               Ct.Cheque.ShowMode = SHOW_ADD
               Call Ct.Cheque.AddEditData
            End If
            Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
         End If
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'เมื่อเป็นการนำฝากจะไม่มีการแก้ไขเกิดขึ้น จะต้องลบแล้วสร้างใหม่
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = 3) Then
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", -1)
            Call Ct.AddEditData
         
            Call Ct.Cheque.DeleteData
         ElseIf (PaymentType = OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If DocType = CASH_DEPOSIT Then
            'ถ้าเป็นการลบ item สำหรับการนำเช็คในมือไปเข้า ธนาคาร ไม่ต้องลบเช็คนั้นทิ้งเพียงแต่ set BANK_FLAG ให้เป็น 'N'
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("OLD_CHECK_ID") > 0 Then 'เช็ค
               Call Ct.Cheque.DeleteData
            End If
         End If
      End If
   Next Ct
   
   For Each Ct In Ua.CashTranItems
      If (DocType = CASH_WITHDRAW) Or (DocType = CASH_WHTHDRAW2) Or (DocType = CASH_TRANSFER) Then
         If Ct.GetFieldValue("CHECK_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.Cheque.UpdateBankBranch
         End If
         
         If Ct.GetFieldValue("TX_TYPE") = "E" Then
            Call Ct.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.SetFieldValue("BANK_ACCOUNT", Ua.GetFieldValue("BANK_ACCOUNT"))
            Call Ct.UpdateBankBranchAccount
         End If
      ElseIf (DocType = CASH_DEPOSIT) Or (DocType = CASH_DEPOSIT2) Then
         If Ct.GetFieldValue("TX_TYPE") = "I" Then
            Call Ct.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.SetFieldValue("BANK_ACCOUNT", Ua.GetFieldValue("BANK_ACCOUNT"))
            Call Ct.UpdateBankBranchAccount
         End If
      End If
   Next Ct
   
   Dim CDP As CCashDocPost
   Dim Ch As CCheque
   Dim BD As CBillingDoc
   Set Ch = New CCheque
   For Each CDP In Ua.PostItems
      If CDP.Flag = "A" Then
         For Each BD In CDP.Post2BD
            If BD.Flag = "A" Then
               BD.AddEditMode = SHOW_ADD
               Call BD.AddEditData
               Call CDP.SetFieldValue("BILLING_DOC_ID", BD.BILLING_DOC_ID)
            ElseIf BD.Flag = "E" Then
               BD.AddEditMode = SHOW_EDIT
               Call BD.AddEditData
            End If
         Next BD
         Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
         CDP.ShowMode = SHOW_ADD
         Call CDP.AddEditData
         
         Call Ch.SetFieldValue("CHEQUE_ID", CDP.GetFieldValue("CHEQUE_ID"))
         If CDP.GetFieldValue("POST_TYPE") = POST_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         ElseIf CDP.GetFieldValue("POST_TYPE") = WAITING_CLEAR Then 'รับแล้วแต่ยังไม่ขึ้นเงิน
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "N")
         ElseIf CDP.GetFieldValue("POST_TYPE") = PASSED_CLEAR Then 'เคลียร์แล้ว
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         End If
         
         Call Ch.UpdateBankFlag
         
      ElseIf CDP.Flag = "E" Then
         For Each BD In CDP.Post2BD
            If BD.Flag = "A" Then
               BD.AddEditMode = SHOW_ADD
               Call BD.AddEditData
               Call CDP.SetFieldValue("BILLING_DOC_ID", BD.BILLING_DOC_ID)
            ElseIf BD.Flag = "E" Then
               BD.AddEditMode = SHOW_EDIT
               Call BD.AddEditData
            End If
         Next BD
         
         CDP.ShowMode = SHOW_EDIT
         Call CDP.AddEditData
      ElseIf CDP.Flag = "D" Then
         Call CDP.DeleteData
      End If
   Next CDP
   
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCashDoc = True
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
   AddEditCashDoc = False
End Function
Public Sub CreateCashTransferItems(Ivd As CCashDoc)
Dim Ti As CCashTransferItem
Dim O As Object
Dim Ei As CCashTran
Dim II As CCashTran
Dim I As Long
Dim J As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.CashTranItems.Count \ 2
   Count2 = Ivd.CashTranItems.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      J = 1
      While J <= Count2
         Set O = Ivd.CashTranItems(J)
         If (O.GetFieldValue("TX_TYPE") = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.GetFieldValue("TX_TYPE") = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set Ei = O
            End If
         End If
         J = J + 1
      Wend
         
      Set Ti = New CCashTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = Ei
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.add(Ti)
      Set Ti = Nothing
   Next I
End Sub
Public Function QueryCheque(Ua As CCheque, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCheque"
   QueryCheque = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryCheque = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCheque = False
End Function
Public Function AddEditCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCheque"
   AddEditCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCheque = True
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
   AddEditCheque = False
End Function
Public Function AddEditChequeDoc(Ua As CChequeDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditChequeDoc"
   AddEditChequeDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   
   Dim RCD  As CReceiptChequeDoc
   For Each RCD In Ua.ChequeDoc
      If RCD.Flag = "A" Then
         RCD.CHEQUE_DOC_ID = Ua.CHEQUE_DOC_ID
         RCD.AddEditMode = SHOW_ADD
         RCD.AddEditData
      ElseIf RCD.Flag = "E" Then
         RCD.AddEditMode = SHOW_EDIT
         RCD.AddEditData
      ElseIf RCD.Flag = "D" Then
         RCD.DeleteData
      End If
   Next RCD

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditChequeDoc = True
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
   AddEditChequeDoc = False
End Function
Public Function AddEditReceiptChequeDoc(Ua As CReceiptChequeDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReceiptChequeDoc"
   AddEditReceiptChequeDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim BD As CBillingDoc
   For Each BD In Ua.ReceiptCheqDocItems
      If BD.Flag = "A" Then
         BD.BILLING_DOC_ID = Ua.BILLING_DOC_ID
         BD.AddEditMode = SHOW_ADD
         BD.AddEditData
      ElseIf BD.Flag = "E" Then
         BD.AddEditMode = SHOW_EDIT
         BD.AddEditData
      ElseIf BD.Flag = "D" Then
         BD.DeleteData
      End If
   Next BD
   
   

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   
   
   
   

   AddEditReceiptChequeDoc = True
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
   AddEditReceiptChequeDoc = False
End Function
Public Function DeleteCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCheque"
   DeleteCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCheque = True
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

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCheque = True
End Function
Public Function QueryMemoNote(Ua As CMemoNote, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryMemoNote"
   QueryMemoNote = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      
   End If
   
   QueryMemoNote = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryMemoNote = False
End Function

Public Function DeleteMemoNote(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CMemoNote

   RName = "DeleteMemoNote"
   DeleteMemoNote = False
   HasBegin = False

   Set B = New CMemoNote

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call B.SetFieldValue("MEMO_NOTE_ID", UID)
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMemoNote = True
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

   DeleteMemoNote = True
End Function

Public Function AddEditMemoNote(Ug As CMemoNote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMemoNote"
   AddEditMemoNote = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMemoNote = True
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
   AddEditMemoNote = False
End Function

Public Function QueryExpense(Ua As CExpense, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryExpense"
   QueryExpense = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim CDP As CExpenseDetail
      Set CDP = New CExpenseDetail
      Call CDP.SetFieldValue("EXPENSE_ID", Ua.GetFieldValue("EXPENSE_ID"))
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing

      Set Ua.ExpenseDetail = Nothing
      Set Ua.ExpenseDetail = New Collection

      While Not m_Rs1.EOF
         Set CDP = New CExpenseDetail
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         Call Ua.ExpenseDetail.add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryExpense = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryExpense = False
End Function
Public Function DeleteExpense(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CExpense

   RName = "DeleteExpense"
   DeleteExpense = False
   HasBegin = False

   Set B = New CExpense

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call B.SetFieldValue("EXPENSE_ID", UID)
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteExpense = True
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

   DeleteExpense = True
End Function
Public Function DeleteLot(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As cLot

   RName = "DeleteLot"
   DeleteLot = False
   HasBegin = False

   Set B = New cLot

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   B.LOT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteLot = True
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

   DeleteLot = False
End Function

Public Function AddEditExpense(Ug As CExpense, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditExpense"
   AddEditExpense = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim CDP As CExpenseDetail
   For Each CDP In Ug.ExpenseDetail
      Call CDP.SetFieldValue("EXPENSE_ID", Ug.GetFieldValue("EXPENSE_ID"))
      If CDP.Flag = "A" Then
         CDP.ShowMode = SHOW_ADD
         Call CDP.AddEditData
      ElseIf CDP.Flag = "E" Then
         CDP.ShowMode = SHOW_EDIT
         Call CDP.AddEditData
      ElseIf CDP.Flag = "D" Then
         Call CDP.DeleteData
      End If
   Next CDP
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditExpense = True
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
   AddEditExpense = False
End Function

Public Function PatchDocumentPriority2()
Dim Lc As CDocumentPriority
   
   Set Lc = New CDocumentPriority
   
   'Call glbDaily.StartTransaction
   
   Call Lc.ClearData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 31
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 21
   Lc.PRIORITY2 = 21
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 2
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 22
   Lc.PRIORITY2 = 22
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 23
   Lc.PRIORITY2 = 23
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 11
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 9
   Lc.PRIORITY2 = 9
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 25
   Lc.PRIORITY2 = 25
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 10
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 26
   Lc.PRIORITY2 = 26
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 21
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 27
   Lc.PRIORITY2 = 27
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 28
   Lc.PRIORITY2 = 28
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 101
   Lc.PRIORITY2 = 101
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 103
   Lc.PRIORITY2 = 103
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 1
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 1
   Lc.PRIORITY2 = 1
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 2
   Lc.PRIORITY2 = 2
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 11
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 20
   Lc.PRIORITY2 = 20
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 19
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 4
   Lc.PRIORITY2 = 4
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 5
   Lc.PRIORITY2 = 5
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 6
   Lc.PRIORITY2 = 6
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 23
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 7
   Lc.PRIORITY2 = 7
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 100
   Lc.PRIORITY2 = 100
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 102
   Lc.PRIORITY2 = 102
   Lc.Area = 1
   Call Lc.AddEditData
   
   'Call glbDaily.CommitTransaction
   Set Lc = Nothing
End Function

Public Function PatchDocumentPriority3()
Dim Lc As CDocumentPriority
   
   Set Lc = New CDocumentPriority
   
   'Call glbDaily.StartTransaction
   
   Call Lc.ClearData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 31
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 21
   Lc.PRIORITY2 = 21
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 2
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 22
   Lc.PRIORITY2 = 22
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 23
   Lc.PRIORITY2 = 23
   Lc.Area = 1
   Call Lc.AddEditData
      
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 25
   Lc.PRIORITY2 = 25
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 10
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 26
   Lc.PRIORITY2 = 26
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 21
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 27
   Lc.PRIORITY2 = 27
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 28
   Lc.PRIORITY2 = 28
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 101
   Lc.PRIORITY2 = 101
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 103
   Lc.PRIORITY2 = 103
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 1
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 1
   Lc.PRIORITY2 = 1
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 2
   Lc.PRIORITY2 = 2
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 19
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 4
   Lc.PRIORITY2 = 4
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 5
   Lc.PRIORITY2 = 5
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 6
   Lc.PRIORITY2 = 6
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 23
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 7
   Lc.PRIORITY2 = 7
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 100
   Lc.PRIORITY2 = 100
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 102
   Lc.PRIORITY2 = 102
   Lc.Area = 1
   Call Lc.AddEditData
   
   'เริ่ม ใบสั่งผลิต แยกมาต่างหาก
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 12
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 11
   Lc.PRIORITY2 = 11
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 12
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 12
   Lc.PRIORITY2 = 12
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 13
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 13
   Lc.PRIORITY2 = 13
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 13
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 14
   Lc.PRIORITY2 = 14
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 14
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 15
   Lc.PRIORITY2 = 15
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 14
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 16
   Lc.PRIORITY2 = 16
   Lc.Area = 1
   Call Lc.AddEditData
   'สิ้นสุด ใบสั่งผลิต แยกมาต่างหาก
      
   'Call glbDaily.CommitTransaction
   Set Lc = Nothing
End Function

Public Function PatchDocumentPriority4()
Dim Lc As CDocumentPriority
   
   Set Lc = New CDocumentPriority
   
   'Call glbDaily.StartTransaction
   
   Call Lc.ClearData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 31
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 21
   Lc.PRIORITY2 = 21
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 2
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 22
   Lc.PRIORITY2 = 22
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 2
   Lc.PRIORITY2 = 2
   Lc.Area = 1
   Call Lc.AddEditData
      
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 25
   Lc.PRIORITY2 = 25
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 10
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 26
   Lc.PRIORITY2 = 26
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 21
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 27
   Lc.PRIORITY2 = 27
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 6
   Lc.PRIORITY2 = 6
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 101
   Lc.PRIORITY2 = 101
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 103
   Lc.PRIORITY2 = 103
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 1
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 1
   Lc.PRIORITY2 = 1
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 3
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 2
   Lc.PRIORITY2 = 2
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 19
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 4
   Lc.PRIORITY2 = 4
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 20
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 5
   Lc.PRIORITY2 = 5
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 22
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 6
   Lc.PRIORITY2 = 6
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 23
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 7
   Lc.PRIORITY2 = 7
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 5
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 100
   Lc.PRIORITY2 = 100
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 4
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 102
   Lc.PRIORITY2 = 102
   Lc.Area = 1
   Call Lc.AddEditData
   
   'เริ่ม ใบสั่งผลิต แยกมาต่างหาก
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 12
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 11
   Lc.PRIORITY2 = 11
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 12
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 12
   Lc.PRIORITY2 = 12
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 13
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 13
   Lc.PRIORITY2 = 13
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 13
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 14
   Lc.PRIORITY2 = 14
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 14
   Lc.TX_TYPE = "E"
   Lc.PRIORITY1 = 15
   Lc.PRIORITY2 = 15
   Lc.Area = 1
   Call Lc.AddEditData
   
   Lc.DOCUMENT_PRIORITY_ID = -1
   Lc.AddEditMode = SHOW_ADD
   Lc.DOCUMENT_TYPE = 14
   Lc.TX_TYPE = "I"
   Lc.PRIORITY1 = 16
   Lc.PRIORITY2 = 16
   Lc.Area = 1
   Call Lc.AddEditData
   'สิ้นสุด ใบสั่งผลิต แยกมาต่างหาก
      
   'Call glbDaily.CommitTransaction
   Set Lc = Nothing
End Function
Public Function AddEditEvaluatePay(Ug As CEvaluatePay, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEvaluatePay"
   AddEditEvaluatePay = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEvaluatePay = True
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
   AddEditEvaluatePay = False
End Function

Public Function QueryEvaluatePay(Ua As CEvaluatePay, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

Dim Qp As CEvaluatePay

   RName = "QueryEvaluatePay"
   QueryEvaluatePay = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryEvaluatePay = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEvaluatePay = False
End Function

Public Function DeleteEvaluatePay(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CEvaluatePay

   RName = "DeleteEvaluatePay"
   DeleteEvaluatePay = False
   HasBegin = False

   Set B = New CEvaluatePay

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call B.SetFieldValue("EVALUATE_PAY_ID", UID)
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEvaluatePay = True
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

   DeleteEvaluatePay = True
End Function

Public Function VerifyReferPo(id As Long, DocumentTypeSet As String, ErrorObj As clsErrorLog) As Boolean
Dim Sp  As CSupItem
Dim m_Rs  As ADODB.Recordset
Dim ItemCount  As Long
   
   VerifyReferPo = False
   
   Set Sp = New CSupItem
   Set m_Rs = New ADODB.Recordset
   Sp.SUP_ITEM_ID = -1
   Sp.DOCUMENT_TYPE_SET = DocumentTypeSet
   Sp.PO_ID = id
   Call Sp.QueryData(26, m_Rs, ItemCount)
   
   If ItemCount > 0 Then
      VerifyReferPo = True
      ErrorObj.LocalErrorMsg = "ไม่สามารถแก้ไขได้เนื่องจากมีการอ้างอิงเอกสารแล้ว"
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
      ErrorObj.ModuleName = m_ModuleName
      ErrorObj.ShowErrorLog (LOG_TO_FILE)
   End If
   
End Function

Public Sub CopyMonthlyAccum(Src As Collection, Dst As Collection)
Dim II As CLotItem
Dim Ma As CMonthlyAccum
   
   For Each Ma In Src
'If Ba.LOCATION_ID = 106 And Ba.PART_ITEM_ID = 123 Then
''Debug.Print
'End If
      Set II = New CLotItem
      II.LOCATION_ID = Ma.LOCATION_ID
      II.PART_ITEM_ID = Ma.PART_ITEM_ID
      
      II.PART_NO = Ma.PART_NO
      II.PART_DESC = Ma.PART_DESC
      
      II.TOTAL_INCLUDE_PRICE = Ma.TOTAL_INCLUDE_PRICE
      II.CURRENT_AMOUNT = Ma.BALANCE_AMOUNT
      II.NEW_PRICE = Ma.AVG_PRICE   'MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = Ma.AVG_PRICE
      II.TX_TYPE = "I"
      II.NEW_AMOUNT = Ma.BALANCE_AMOUNT
      Call Dst.add(II, Ma.LOCATION_ID & "-" & Ma.PART_ITEM_ID)
      Set II = Nothing
   Next Ma
End Sub
Public Sub CopyMonthlyAccumPartItem(Src As Collection, Dst As Collection)
Dim II As CLotItem
Dim Ma As CMonthlyAccum
   
   For Each Ma In Src
      Set II = New CLotItem
      II.PART_ITEM_ID = Ma.PART_ITEM_ID
      II.TOTAL_INCLUDE_PRICE = Ma.TOTAL_INCLUDE_PRICE
      II.CURRENT_AMOUNT = Ma.BALANCE_AMOUNT
      II.NEW_PRICE = MyDiffEx(Ma.TOTAL_INCLUDE_PRICE, Ma.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = II.NEW_PRICE
      II.TX_TYPE = "I"
      II.NEW_AMOUNT = Ma.BALANCE_AMOUNT
      Call Dst.add(II, Trim(str(Ma.PART_ITEM_ID)))
      Set II = Nothing
   Next Ma
End Sub

Public Function AddEditPlanPart(Ug As CPlanPart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPlanPart"
   AddEditPlanPart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPlanPart = True
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
   AddEditPlanPart = False
End Function
Public Function DeletePlanPart(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CPlanPart

   RName = "DeletePlanPart"
   DeletePlanPart = False
   HasBegin = False

   Set B = New CPlanPart

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PLAN_PART_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePlanPart = True
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

   DeletePlanPart = True
End Function

Public Function QueryPlanPart(Ua As CPlanPart, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

Dim Qp As CPlanPart

   RName = "QueryPlanPart"
   QueryPlanPart = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
      
   QueryPlanPart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPlanPart = False
End Function
Public Function QueryPalletDoc(Ua As CPalletDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog, Optional Ind As Long = 1) As Boolean
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

Dim Qp As CPalletDoc

   RName = "QueryPalletDoc"
   QueryPalletDoc = False

   IsOK = True
   Call Ua.QueryData(Ind, Rs, ItemCount)
      
   QueryPalletDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPalletDoc = False
End Function
Public Function AddEditAlertBox(Ug As CAlertBox, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditAlertBox"
   AddEditAlertBox = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Ald As CAlertDetail
   For Each Ald In Ug.CollAlertDetail
      If Ald.Flag = "A" Then
         Ald.ALERT_BOX_ID = Ug.ALERT_BOX_ID
         Ald.AddEditMode = SHOW_ADD
         Call Ald.AddEditData
      ElseIf Ald.Flag = "E" Then
         Ald.AddEditMode = SHOW_EDIT
         Call Ald.AddEditData
      ElseIf Ald.Flag = "D" Then
         Call Ald.DeleteData
      End If
   Next Ald
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditAlertBox = True
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
   AddEditAlertBox = False
End Function
Public Function QueryAlertBox(Ua As CAlertBox, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryAlertBox"
   QueryAlertBox = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      Dim Ald As CAlertDetail
      Set Ald = New CAlertDetail
      Ald.ALERT_BOX_ID = Ua.ALERT_BOX_ID
      Call Ald.QueryData(1, m_Rs1, iCount)
      Set Ald = Nothing
      
      Set Ua.CollAlertDetail = Nothing
      Set Ua.CollAlertDetail = New Collection
      While Not m_Rs1.EOF
         Set Ald = New CAlertDetail
         Call Ald.PopulateFromRS(1, m_Rs1)

         Ald.Flag = "I"
         Call Ua.CollAlertDetail.add(Ald)
         
         Set Ald = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryAlertBox = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAlertBox = False
End Function

Public Function DeleteAlertBox(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim B As CAlertBox

   RName = "DeleteAlertBox"
   DeleteAlertBox = False
   HasBegin = False

   Set B = New CAlertBox

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   B.ALERT_BOX_ID = UID
   Call B.DeleteData
   
   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteAlertBox = True
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

   DeleteAlertBox = True
End Function
Public Function UpdateLogonStatus(Status As Long, IpAddress As String)
Dim RName As String
Dim Ua As CUserAccount

   RName = "UpdateLogonStatus"
   UpdateLogonStatus = False
   
   Set Ua = New CUserAccount
   
   Call Ua.UpdateLogonStatus(Status, IpAddress)
   
   Set Ua = Nothing
   
   UpdateLogonStatus = True
End Function
Public Function QueryLegacyWeight(Ua As CWeight, Rs As ADODB.Recordset, Rs2 As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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
Dim CW As CWeight
Dim tmpRS As ADODB.Recordset
Dim tmpCount As Long

   RName = "QueryLegacyWeight"
   QueryLegacyWeight = False

   IsOK = True
   Set CW = New CWeight
   CW.Date1 = Ua.Date1
   CW.CUST_ID = Ua.CUST_ID
   CW.TRUCK_ID = Ua.TRUCK_ID
   CW.OrderBy = Ua.OrderBy
   CW.OrderType = Ua.OrderType
   
   Call CW.QueryData(1, Rs, tmpCount)
   ItemCount = tmpCount
   Call CW.QueryData(2, Rs2, tmpCount)
   ItemCount = ItemCount + tmpCount
   QueryLegacyWeight = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLegacyWeight = False
End Function
Public Function AddEditWeight(Ua As CWeight, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditWeight"
   AddEditWeight = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditWeight = True
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
   AddEditWeight = False
End Function

Public Function DeleteMasterValid(Ug As CMasterValid, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteMasterValid"
   DeleteMasterValid = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterValid = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   
   IsOK = False

   DeleteMasterValid = True
End Function
Public Function QueryMasterValid(Ua As CMasterValid, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long

   RName = "QueryMasterValid"
   QueryMasterValid = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   Call Ua.PopulateFromRS(1, Rs)
   If Ua.QueryFlag = 1 Then
      Dim Cmss As CCommissionSale
      Set Cmss = New CCommissionSale
      Cmss.MASTER_VALID_ID = Ua.MASTER_VALID_ID
      Call Cmss.QueryData(1, m_Rs1, iCount)
      Set Cmss = Nothing
      
      Set Ua.CollSaleRcp = Nothing
      Set Ua.CollSaleRcp = New Collection
      Set Ua.CollSaleNow = Nothing
      Set Ua.CollSaleNow = New Collection
      Set Ua.CollSaleManagerNow = Nothing
      Set Ua.CollSaleManagerNow = New Collection
      
      While Not m_Rs1.EOF
         Set Cmss = New CCommissionSale
         Call Cmss.PopulateFromRS(1, m_Rs1)
         If Cmss.COMMISSION_SALE_AREA = 1 Then
            Call Ua.CollSaleRcp.add(Cmss)
         ElseIf Cmss.COMMISSION_SALE_AREA = 2 Then
            Call Ua.CollSaleNow.add(Cmss)
         ElseIf Cmss.COMMISSION_SALE_AREA = 3 Then
            Call Ua.CollSaleManagerNow.add(Cmss)
         End If
         Set Cmss = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim CmsCost As CCommissionCost
      Set CmsCost = New CCommissionCost
      CmsCost.MASTER_VALID_ID = Ua.MASTER_VALID_ID
      Call CmsCost.QueryData(1, m_Rs1, iCount)
      Set CmsCost = Nothing
      
      Set Ua.CollCommissionCost = Nothing
      Set Ua.CollCommissionCost = New Collection
      
      While Not m_Rs1.EOF
         Set CmsCost = New CCommissionCost
         Call CmsCost.PopulateFromRS(1, m_Rs1)
         Call Ua.CollCommissionCost.add(CmsCost)
         Set CmsCost = Nothing
         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryMasterValid = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryMasterValid = False
End Function
Public Function CopyMasterValid(Ug As CMasterValid, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyMasterValid"
   CopyMasterValid = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryMasterValid(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   If Ug.MASTER_VALID_TYPE = COMMISSION_BUDGET_CHART Then
      Dim Cm As CCommissionBgChart
      Dim m_CommissionBudgetCharts As Collection
      Set Cm = New CCommissionBgChart
      Set m_CommissionBudgetCharts = New Collection
      Call LoadCommissionBudgetChart(Nothing, m_CommissionBudgetCharts, Ug.MASTER_VALID_ID)
   End If

   
   Ug.AddEditMode = SHOW_ADD
   Ug.MASTER_VALID_ID = -1
   Ug.MASTER_VALID_NO = Ug.MASTER_VALID_NO & "***"
   
   Dim Cmss As CCommissionSale
   For Each Cmss In Ug.CollSaleRcp
      Cmss.Flag = "A"
      Cmss.AddEditMode = SHOW_ADD
   Next Cmss
   For Each Cmss In Ug.CollSaleNow
      Cmss.Flag = "A"
      Cmss.AddEditMode = SHOW_ADD
   Next Cmss
   For Each Cmss In Ug.CollSaleManagerNow
      Cmss.Flag = "A"
      Cmss.AddEditMode = SHOW_ADD
   Next Cmss
   Dim CmsCost As CCommissionCost
   For Each CmsCost In Ug.CollCommissionCost
      CmsCost.Flag = "A"
      CmsCost.AddEditMode = SHOW_ADD
   Next CmsCost
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditMasterValid(Ug, IsOK, False, ErrorObj)
      
   If Ug.MASTER_VALID_TYPE = COMMISSION_BUDGET_CHART Then
      Dim CmC As CCommissionBgChart
      Dim TempSearch As CCommissionBgChart
      
      
      For Each CmC In m_CommissionBudgetCharts
         CmC.AddEditMode = SHOW_ADD
         CmC.MASTER_VALID_ID = Ug.MASTER_VALID_ID
         Call CmC.AddEditData
      Next CmC
      
      For Each CmC In m_CommissionBudgetCharts
         If CmC.PARENT_ID > 0 Then
            CmC.AddEditMode = SHOW_EDIT
            Set TempSearch = GetObject("CCommissionBgChart", m_CommissionBudgetCharts, CmC.PARENT_ID)
            CmC.PARENT_ID = TempSearch.COMMISSION_BUDGET_CHART_ID
            Call CmC.AddEditData
         End If
      Next CmC
   End If
   
      
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

   CopyMasterValid = True
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
   CopyMasterValid = False
End Function
Public Function CopyExWorksPrice(Ug As CExWorksPrice, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyExWorksPrice"
   CopyExWorksPrice = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryExWorksPrice(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If

   Ug.AddEditMode = SHOW_ADD
   Ug.EX_WORKS_PRICE_ID = -1
   Ug.EX_WORKS_PRICE_CODE = Ug.EX_WORKS_PRICE_CODE & "***"
   
   If Ug.EX_WORKS_PRICE_TYPE = 1 Then
   Dim Cmss As CExWorksPriceItem
   For Each Cmss In Ug.ExWorksPriceItem
      Cmss.Flag = "A"
      Cmss.AddEditMode = SHOW_ADD
   Next Cmss
   
   ElseIf Ug.EX_WORKS_PRICE_TYPE = 2 Then
   
   Dim CmsCost As CExDeliveryCostItem
   For Each CmsCost In Ug.ExDeliveryCost
      CmsCost.Flag = "A"
      CmsCost.AddEditMode = SHOW_ADD
   Next CmsCost
   
   ElseIf Ug.EX_WORKS_PRICE_TYPE = 3 Then
   Dim Ceppi As CExPromotionPartItem
   For Each Ceppi In Ug.ExPromotionPart
      Ceppi.Flag = "A"
      Ceppi.AddEditMode = SHOW_ADD
   Next Ceppi
   
   ElseIf Ug.EX_WORKS_PRICE_TYPE = 4 Then
   Dim Cepdi As CExPromotionDlcItem
   For Each Cepdi In Ug.ExPromotionDlc
      Cepdi.Flag = "A"
      Cepdi.AddEditMode = SHOW_ADD
   Next Cepdi
   End If
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditExWorksPrice(Ug, IsOK, False, ErrorObj)
      
'   If Ug.MASTER_VALID_TYPE = COMMISSION_BUDGET_CHART Then
'      Dim CmC As CCommissionBgChart
'      Dim TempSearch As CCommissionBgChart
'
'      For Each CmC In m_CommissionBudgetCharts
'         CmC.AddEditMode = SHOW_ADD
'         CmC.MASTER_VALID_ID = Ug.MASTER_VALID_ID
'         Call CmC.AddEditData
'      Next CmC
'
'      For Each CmC In m_CommissionBudgetCharts
'         If CmC.PARENT_ID > 0 Then
'            CmC.AddEditMode = SHOW_EDIT
'            Set TempSearch = GetObject("CCommissionBgChart", m_CommissionBudgetCharts, CmC.PARENT_ID)
'            CmC.PARENT_ID = TempSearch.COMMISSION_BUDGET_CHART_ID
'            Call CmC.AddEditData
'         End If
'      Next CmC
'   End If
   
      
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

   CopyExWorksPrice = True
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
   CopyExWorksPrice = False
End Function
Public Function AddEditMasterValid(Ua As CMasterValid, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterValid"
   AddEditMasterValid = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Cmss As CCommissionSale
   For Each Cmss In Ua.CollSaleRcp
      Cmss.MASTER_VALID_ID = Ua.MASTER_VALID_ID
      Cmss.COMMISSION_SALE_AREA = 1
      If Cmss.Flag = "A" Then
         Cmss.AddEditMode = SHOW_ADD
         Call Cmss.AddEditData
      ElseIf Cmss.Flag = "E" Then
         Cmss.AddEditMode = SHOW_EDIT
         Call Cmss.AddEditData
      ElseIf Cmss.Flag = "D" Then
         Call Cmss.DeleteData
      End If
   Next Cmss
   
   For Each Cmss In Ua.CollSaleNow
      Cmss.MASTER_VALID_ID = Ua.MASTER_VALID_ID
      Cmss.COMMISSION_SALE_AREA = 2
      If Cmss.Flag = "A" Then
         Cmss.AddEditMode = SHOW_ADD
         Call Cmss.AddEditData
      ElseIf Cmss.Flag = "E" Then
         Cmss.AddEditMode = SHOW_EDIT
         Call Cmss.AddEditData
      ElseIf Cmss.Flag = "D" Then
         Call Cmss.DeleteData
      End If
   Next Cmss
   
   For Each Cmss In Ua.CollSaleManagerNow
      Cmss.MASTER_VALID_ID = Ua.MASTER_VALID_ID
      Cmss.COMMISSION_SALE_AREA = 3
      If Cmss.Flag = "A" Then
         Cmss.AddEditMode = SHOW_ADD
         Call Cmss.AddEditData
      ElseIf Cmss.Flag = "E" Then
         Cmss.AddEditMode = SHOW_EDIT
         Call Cmss.AddEditData
      ElseIf Cmss.Flag = "D" Then
         Call Cmss.DeleteData
      End If
   Next Cmss
   
   Dim CmsCost As CCommissionCost
   For Each CmsCost In Ua.CollCommissionCost
      CmsCost.MASTER_VALID_ID = Ua.MASTER_VALID_ID
      If CmsCost.Flag = "A" Then
         CmsCost.AddEditMode = SHOW_ADD
         Call CmsCost.AddEditData
      ElseIf CmsCost.Flag = "E" Then
         CmsCost.AddEditMode = SHOW_EDIT
         Call CmsCost.AddEditData
      ElseIf CmsCost.Flag = "D" Then
         Call CmsCost.DeleteData
      End If
   Next CmsCost
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   AddEditMasterValid = True
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
   AddEditMasterValid = False
End Function

Public Function DeleteCommissionBudgetChart(Ua As CCommissionBgChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCommissionBudgetChart"
   DeleteCommissionBudgetChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissionBudgetChart = True
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

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCommissionBudgetChart = True
End Function
Public Function AddEditCommissionBudgetChart(Ua As CCommissionBgChart, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissionBudgetChart"
   AddEditCommissionBudgetChart = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditCommissionBudgetChart = True
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
   AddEditCommissionBudgetChart = False
End Function
Public Function QueryCommissionBudgetChart(Ua As CCommissionBgChart, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCommissionBudgetChart"
   QueryCommissionBudgetChart = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryCommissionBudgetChart = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCommissionBudgetChart = False
End Function
Public Function QueryCommissionSubtract(Ua As CCommissionSubtract, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCommissionSubtract"
   QueryCommissionSubtract = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   QueryCommissionSubtract = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCommissionSubtract = False
End Function

Public Function QueryCommissionIncentive(Ua As CCommissionIncentive, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCommissionIncentive"
   QueryCommissionIncentive = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   QueryCommissionIncentive = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCommissionIncentive = False
End Function

Public Function AddEditCommissionSubtract(Ug As CCommissionSubtract, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissionSubtract"
   AddEditCommissionSubtract = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCommissionSubtract = True
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
   AddEditCommissionSubtract = False
End Function

Public Function AddEditCommissionIncentive(Ug As CCommissionIncentive, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissionIncentive"
   AddEditCommissionIncentive = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCommissionIncentive = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   If InStr(Err.DESCRIPTION, "UQ_INCENTIVE") > 0 Then
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION & " " & "มีข้อมูล PC และ สินค้า ที่เลือกอยู่แล้ว"
   Else
      ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   End If
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCommissionIncentive = False
End Function

Public Function DeleteCommissionSubtract(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim Cmsst As CCommissionSubtract

   RName = "DeleteCommissionSubtract"
   DeleteCommissionSubtract = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set Cmsst = New CCommissionSubtract
   
   Cmsst.COMMISSION_SUBTRACT_ID = UID
   Call Cmsst.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissionSubtract = True
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

   DeleteCommissionSubtract = True
End Function

Public Function DeleteCommissionIncentive(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim Cmsst As CCommissionIncentive

   RName = "DeleteCommissionIncentive"
   DeleteCommissionIncentive = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set Cmsst = New CCommissionIncentive
   
   Cmsst.INCENTIVE_ID = UID
   Call Cmsst.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissionIncentive = True
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

   DeleteCommissionIncentive = True
End Function

Public Function QueryCommissionCredit(Ua As CCommissionCredit, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryCommissionCredit"
   QueryCommissionCredit = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   QueryCommissionCredit = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCommissionCredit = False
End Function

Public Function AddEditCommissionCredit(Ug As CCommissionCredit, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCommissionCredit"
   AddEditCommissionCredit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCommissionCredit = True
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
   AddEditCommissionCredit = False
End Function

Public Function DeleteCommissionCredit(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim Cmsst As CCommissionCredit

   RName = "DeleteCommissionCredit"
   DeleteCommissionCredit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set Cmsst = New CCommissionCredit
   
   Cmsst.COMMISSION_CREDIT_ID = UID
   Call Cmsst.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCommissionCredit = True
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

   DeleteCommissionCredit = True
End Function
Public Function QueryLockDate(Ua As CLockDate, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryLockDate"
   QueryLockDate = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryLockDate = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLockDate = False
End Function
Public Function AddEditLockDate(Ua As CLockDate, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLockDate"
   AddEditLockDate = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditLockDate = True
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
   AddEditLockDate = False
End Function
