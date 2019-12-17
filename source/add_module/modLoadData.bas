Attribute VB_Name = "modLoadData"
Option Explicit
' Test test test
Public Sub InitPaymentType3(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentType2Text(1))
   C.ItemData(1) = 1
   '
'   C.AddItem (PaymentType2Text(2))
'   C.ItemData(2) = 2

'   C.AddItem (PaymentType2Text(3))
'   C.ItemData(2) = 3
End Sub

Public Sub InitDiscountType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ลดต่อถุง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ลดต่อน้ำหนัก"))
   C.ItemData(2) = 2

   C.AddItem (MapText("ลดต่อ 100 บาท"))
   C.ItemData(3) = 3
End Sub
Public Sub InitPlanningOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ประมาณ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("จากวันที่ประมาณ"))
   C.ItemData(2) = 2

   C.AddItem (MapText("ถึงวันที่ประมาณ"))
   C.ItemData(3) = 3
End Sub
Public Sub InitInvActOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่นับสต๊อก"))
   C.ItemData(1) = 1

'   C.AddItem (MapText("จากวันที่นับสต๊อก"))
'   C.ItemData(2) = 2
'
'   C.AddItem (MapText("ถึงวันที่นับสต๊อก"))
'   C.ItemData(3) = 3
End Sub
Public Function PlanningArea2Text(Area As Long) As String
   If Area = 1 Then
      PlanningArea2Text = "ประมาณการการใช้วัตถุดิบและผลิตสินค้ารายวัน"
   ElseIf Area = 2 Then
      PlanningArea2Text = "ประมาณการการใช้วัตถุดิบและผลิตสินค้ารายสัปดาห์"
   ElseIf Area = 3 Then
      PlanningArea2Text = "ประมาณการรับเข้าวัตถุดิบรายวันจากซัพพลายเออร์"
   ElseIf Area = 4 Then
      PlanningArea2Text = "ประมาณการการใช้วัตถุดิบและผลิตสินค้ารายเดือน"
   End If
End Function
Public Function InventoryActArea2Text(Area As Long) As String
   If Area = 1 Then
      InventoryActArea2Text = "การตรวจนับจากโกดังวัตถุดิบ"
   ElseIf Area = 2 Then
      InventoryActArea2Text = "การตรวจนับจากห้องยา"
   ElseIf Area = 3 Then
      InventoryActArea2Text = "การตรวจนับจากไซโล"
   End If
End Function
Public Function InventoryWhActArea2Text(Area As Long) As String
   If Area = 14 Then
      InventoryWhActArea2Text = "การตรวจนับจากโกดังสินค้า (BAG)"
   ElseIf Area = 13 Then
      InventoryWhActArea2Text = "การตรวจนับจากโกดังสินค้า (BULK)"
   End If
End Function
Public Function InventoryProductArea2Text(Area As Long) As String
   If Area = 1 Then
      InventoryProductArea2Text = "ข้อมูลการบรรจุอาหารจาก PACKING"
   ElseIf Area = 2 Then
      InventoryProductArea2Text = "ข้อมูลการโหลดอาหารจาก BULK"
   ElseIf Area = 3 Then
      InventoryProductArea2Text = "ข้อมูลคงเหลืออาหารสำเร็จรูป"
   End If
End Function
Public Function InventoryActArea2Text2(Area As Long) As String
   If Area = 1 Then
      InventoryActArea2Text2 = "RAW-MATERIAL"
   ElseIf Area = 2 Then
      InventoryActArea2Text2 = "PHAMACY-ROOM"
   ElseIf Area = 3 Then
      InventoryActArea2Text2 = "SILO"
   End If
End Function
Public Function InventoryWhActArea2Text2(Area As Long) As String
   If Area = 14 Then
      InventoryWhActArea2Text2 = "INVENTORY-WH-BAG"
   ElseIf Area = 13 Then
      InventoryWhActArea2Text2 = "INVENTORY-WH-BULK"
   End If
End Function

Public Sub InitJobOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ใบสั่งผลิต"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่สั่งผลิต"))
   C.ItemData(2) = 2

   C.AddItem (MapText("หมายเลขแบต"))
   C.ItemData(3) = 3
End Sub

Public Sub InitCostProductionOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub
Public Sub InitPackProductionOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Function CalculateTypeToText(Ind As Long) As String
   If Ind = 1 Then
      CalculateTypeToText = "คิดตามน้ำหนักผู้ขาย"
   ElseIf Ind = 2 Then
      CalculateTypeToText = "คิดตามน้ำหนักรวม"
   ElseIf Ind = 3 Then
      CalculateTypeToText = "คิดตามน้ำหนักสุทธิ"
   End If
End Function
Public Function CalculateTypeToText2(Ind As Long) As String
   If Ind = 1 Then
      CalculateTypeToText2 = "ผู้ขาย"
   ElseIf Ind = 2 Then
      CalculateTypeToText2 = "รวม"
   ElseIf Ind = 3 Then
      CalculateTypeToText2 = "สุทธิ"
   End If
End Function

Public Sub InitImportType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("นำเข้าสต็อคยกมา")
   C.ItemData(1) = 1
   
   C.AddItem ("ตั้งยอดสต็อค")
   C.ItemData(2) = 6
   
   C.AddItem ("นำเข้าลูกหนี้ยกมา")
   C.ItemData(3) = 2

   C.AddItem ("นำเข้าข้อมูลลูกหนี้")
   C.ItemData(4) = 3

   C.AddItem ("นำเข้าข้อมูล BALANCE ACCUM")
   C.ItemData(5) = 4

   C.AddItem ("เปลี่ยนรหัสลูกค้า")
   C.ItemData(6) = 5
   
   C.AddItem ("นำเข้าเจ้าหนี้ยกมา")
   C.ItemData(7) = 7
   
End Sub

Public Sub InitImportType2(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("Mapping บัญชีสินค้า")
   C.ItemData(1) = 1
   
   C.AddItem ("Mapping บัญชีงานบริการ")
   C.ItemData(2) = 2

   C.AddItem ("Mapping บัญชีธนาคาร")
   C.ItemData(3) = 3
End Sub

Public Sub InitPatchTable(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("PART_ITEM")
   C.ItemData(1) = 1
   
   C.AddItem ("SUPPLIER")
   C.ItemData(2) = 2
   
   C.AddItem ("CUSTOMER")
   C.ItemData(3) = 3
   
   C.AddItem ("INVENTORY_DOC")
   C.ItemData(4) = 4
   
   C.AddItem ("CHEQUE")
   C.ItemData(5) = 5
   
   C.AddItem ("CASH_DOC")
   C.ItemData(6) = 8

   C.AddItem ("BILLING_DOC")
   C.ItemData(7) = 9
   
   C.AddItem ("FORMULA")
   C.ItemData(8) = 7
   
   C.AddItem ("JOB")
   C.ItemData(9) = 6
End Sub

Public Sub InitCalculateType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (CalculateTypeToText(1))
   C.ItemData(1) = 1

   C.AddItem (CalculateTypeToText(2))
   C.ItemData(2) = 2

   C.AddItem (CalculateTypeToText(3))
   C.ItemData(3) = 3
End Sub

Public Sub InitFamilyStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("มีครอบครัว")
   C.ItemData(1) = 1
   
   C.AddItem ("โสด")
   C.ItemData(2) = 2
   
   C.AddItem ("หม้าย")
   C.ItemData(3) = 3
End Sub

Public Sub LoadBloodSpec(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBloodSpec
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBloodSpec
Dim I As Long

   Set D = New CBloodSpec
   Set Rs = New ADODB.Recordset
   
   D.BLOOD_SPEC_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBloodSpec
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SPEC_NAME)
         C.ItemData(I) = TempData.BLOOD_SPEC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.BLOOD_SPEC_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPeriodDesc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDSheetItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDSheetItem
Dim I As Long

   Set D = New CDSheetItem
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData2(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDSheetItem
      Call TempData.PopulateFromRS2(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PERIOD_DESC)
         C.ItemData(I) = TempData.DSHEET_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadXCollection(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CXCollection
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CXCollection
Dim I As Long

   Set D = New CXCollection
   Set Rs = New ADODB.Recordset
   
   D.X_COLLECTION_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New CXCollection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CXCollection
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.X_COLLECTION_NAME)
         C.ItemData(I) = TempData.X_COLLECTION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.X_COLLECTION_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Public Sub LoadYCollection(C As ComboBox, Optional Cl As Collection = Nothing)
'On Error GoTo ErrorHandler
'Dim ItemCount As Long
'Dim Rs As ADODB.Recordset
'Dim TempData As CYCollection
'Dim i As Long
'
'   Set D = New CYCollection
'   Set Rs = New ADODB.Recordset
'
'   D.Y_COLLECTION_ID = -1
'   Call D.QueryData(Rs, ItemCount)
'
'   If Not (C Is Nothing) Then
'      C.Clear
'      i = 0
'      C.AddItem ("")
'   End If
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New CYCollection
'   End If
'   While Not Rs.EOF
'      i = i + 1
'      Set TempData = New CYCollection
'      Call TempData.PopulateFromRS(Rs)
'
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.Y_COLLECTION_NAME)
'         C.ItemData(i) = TempData.Y_COLLECTION_ID
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.Add(TempData, Str(TempData.Y_COLLECTION_ID))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub

Public Sub InitUserGroupOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(1) = 1
End Sub

Public Sub InitUserStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใช้งานได้")
   C.ItemData(1) = 1

   C.AddItem ("ถูกระงับ")
   C.ItemData(2) = 2
End Sub

Public Sub InitLoginOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ล็อคอิน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อผู้ใช้"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport8_11Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport8_12Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสผลิตภัณฑ์"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport8_13Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport8_14Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสผลิตภัณฑ์"))
   C.ItemData(1) = 1
End Sub
Public Sub InitCboCondition(C As ComboBox, Types As Long)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   If Types = 1 Then
      C.AddItem (MapText("คลุมผ้าใบ"))
      C.ItemData(1) = 1
      
      C.AddItem (MapText("ไม่คลุม"))
      C.ItemData(2) = 2
      
      C.ListIndex = 0
   ElseIf Types = 2 Then
      C.AddItem (MapText("สะอาด/แห้ง"))
      C.ItemData(1) = 1
      
      C.AddItem (MapText("ไม่สะอาด"))
      C.ItemData(2) = 2
      
      C.ListIndex = 0
   ElseIf Types = 3 Then
      C.AddItem (MapText("ปกติ"))
      C.ItemData(1) = 1
      
      C.AddItem (MapText("ชำรุด"))
      C.ItemData(2) = 2
      
      C.ListIndex = 0
   End If
End Sub
Public Sub InitTxType(C As ComboBox)
   C.Clear
   
   C.AddItem (MapText(""))
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รับเข้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("จ่ายออก"))
   C.ItemData(2) = 2
   
   C.ListIndex = 1
End Sub
Public Sub InitReport8_18Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2
   
   C.AddItem (MapText("รหัสสูตร"))
   C.ItemData(1) = 2

   C.AddItem (MapText("วันที่สูตร"))
   C.ItemData(2) = 3

'   C.AddItem (MapText("รหัสผลิตภัณฑ์"))
'   C.ItemData(3) = 3
End Sub

Public Sub InitReport4_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 5

   C.AddItem (MapText("รหัสขาย"))
   C.ItemData(2) = 7
End Sub
Public Sub InitReport2_3_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(1) = 1

End Sub
Public Sub InitReport2_3_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสถานที่จัดส่ง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสลูกค้า และ สถานที่จัดส่ง"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReport2_3_3Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(2) = 2

End Sub
Public Sub InitReport2_3_4Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสถานที่จัดส่ง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสลูกค้า และ สถานที่จัดส่ง"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(3) = 3
End Sub
Public Sub InitReport4_8Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport4_10Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ใบรับวัตถุดิบ"))
   C.ItemData(1) = 8
End Sub

Public Sub InitReport4_11Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 2

   C.AddItem (MapText("รหัสพัสดุ"))
   C.ItemData(2) = 3

End Sub

Public Sub InitReport4_9Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสผู้ขาย"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อผู้ขาย"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport4_A_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสผู้ขาย"))
   C.ItemData(1) = 12

   C.AddItem (MapText("ชื่อผู้ขาย"))
   C.ItemData(2) = 13
End Sub

Public Sub InitReport4_A_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 6
End Sub

Public Sub InitReport4_4Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 7
   
   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 7

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(2) = 8

   C.AddItem (MapText("รหัสพัสดุ"))
   C.ItemData(3) = 15

   C.AddItem (MapText("ชื่อพัสดุ"))
   C.ItemData(4) = 16
End Sub
Public Sub InitReport1_11_5Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(1) = 4

   C.AddItem (MapText("ชื่อวัตถุดิบ"))
   C.ItemData(2) = 5

   C.AddItem (MapText("จำนวน PLAN"))
   C.ItemData(3) = 6
   
   C.ListIndex = 3
End Sub
Public Sub InitReportF1_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 3
   
   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(3) = 3
   
End Sub

Public Sub InitReport4_7Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสประเภทวัตถุดิบ"))
   C.ItemData(1) = 5

   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(2) = 6

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(3) = 7

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(4) = 8

   C.AddItem (MapText("รหัสสถานที่จัดเก็บ"))
   C.ItemData(5) = 9
End Sub

Public Sub InitReport5_2OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReportA_2_6Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 1

   C.AddItem (MapText("ประเภทการชำระเงิน"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReportA_2_7Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReportA_2_8Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("วันที่รับชำระ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่ใบเสร็จ"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReportA_2_4Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 2

   C.AddItem (MapText("ชื่อพนักงานขาย"))
   C.ItemData(2) = 3

   C.AddItem (MapText("รหัสลูกค้า-ชื่อพนักงานขาย"))
   C.ItemData(3) = 4

   C.AddItem (MapText("ชื่อพนักงานขาย-รหัสลูกค้า"))
   C.ItemData(4) = 5
End Sub

Public Sub InitReportA_2_21Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2

End Sub

Public Sub InitReportAP_1_3Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 2

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 2
End Sub

Public Sub InitReportA_2_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 3

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 3

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 4
End Sub

Public Sub InitReport5_3OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่บิลขาย"))
   C.ItemData(1) = 2

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(2) = 4
End Sub
Public Sub InitReportAP_1_11_7Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ชื่อพนักงานขาย"))
   C.ItemData(4) = 4
End Sub
Public Sub InitReport5_6OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 3
   
   C.AddItem (MapText("ทะเบียนรถ"))
   C.ItemData(1) = 3
End Sub
Public Sub InitReport5_7OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(2) = 2
   
'   C.AddItem (MapText("ทะเบียนรถ"))
'   C.ItemData(2) = 2
End Sub

Public Sub InitReportA_1_3OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_4OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ประเภทบรรจุ"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport5_5Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReportA_1_4Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(5) = 5
   
   
End Sub

Public Sub InitReport5_14Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1
End Sub

Public Sub InitReport5_11Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสโรงเรือน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อโรงเรือน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitReport5_12Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("สัปดาห์เกิด"))
   C.ItemData(1) = 1
End Sub

Public Sub InitUserOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อผู้ใช้")
   C.ItemData(1) = 1

   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitDocumentType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใบนำเข้า")
   C.ItemData(1) = 1

   C.AddItem ("ใบเบิกวัตถุดิบ")
   C.ItemData(2) = 2

   C.AddItem ("ใบโอนวัตถุดิบ")
   C.ItemData(3) = 3

   C.AddItem ("ใบปรับยอด")
   C.ItemData(4) = 4
End Sub

Public Sub InitCommitStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("คำนวณแล้ว")
   C.ItemData(1) = 1

   C.AddItem ("ยังไม่คำนวณ")
   C.ItemData(2) = 2
End Sub

Public Sub InitCustomerOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("รหัสพนักงานขาย"))
   C.ItemData(3) = 4
End Sub
'===

Public Sub InitSupplierOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสซัพพลายเออร์"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อซัพพลายเออร์"))
   C.ItemData(2) = 2
End Sub
'===
Public Sub InitSupplierDocNoOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสซัพพลายเออร์"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อซัพพลายเออร์"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(3) = 3

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("รหัส R/M"))
   C.ItemData(5) = 5
   
   C.AddItem (MapText("ชื่อ R/M"))
   C.ItemData(6) = 6
   
End Sub
'===

Public Sub InitEmployeeOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อ"))
   C.ItemData(2) = 2

   C.AddItem (MapText("นามสกุล"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ตำแหน่ง"))
   C.ItemData(4) = 4
End Sub

Public Sub InitFreelanceOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อ"))
   C.ItemData(2) = 2

   C.AddItem (MapText("นามสกุล"))
   C.ItemData(3) = 3

End Sub
'===

Public Sub InitIncentiveOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PC"))
   C.ItemData(1) = 1

   C.AddItem (MapText("สินค้า"))
   C.ItemData(2) = 2

End Sub
Public Sub InitPartItemOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขวัตถุดิบ"))
   C.ItemData(1) = 5

   C.AddItem (MapText("ชื่อวัตถุดิบ"))
   C.ItemData(2) = 2
End Sub

Public Sub InitGoodsOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เบอร์สินค้า"))
   C.ItemData(1) = 5

   C.AddItem (MapText("ชื่อสินค้า"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("รหัสขาย"))
   C.ItemData(3) = 3
End Sub
Public Sub InitGoodsOrderBy2(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เบอร์สินค้า"))
   C.ItemData(1) = 5
   
   C.AddItem (MapText("รหัสขาย"))
   C.ItemData(2) = 3

   C.AddItem (MapText("ชนิดสินค้า"))
   C.ItemData(3) = 2
   
   C.AddItem (MapText("ประเภทสินค้า"))
   C.ItemData(4) = 1
   
   C.AddItem (MapText("วันที่ผลิต"))
   C.ItemData(5) = 4
   
   C.AddItem (MapText("LOT"))
   C.ItemData(6) = 6
   

End Sub
'===
Public Sub InitLoadPartType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("อาหาร BAG"))
   C.ItemData(1) = 14
   
   C.AddItem (MapText("อาหาร BULK"))
   C.ItemData(2) = 13
   
   C.AddItem (MapText("RE-BAG -> BAG"))
   C.ItemData(3) = 17
   
   C.AddItem (MapText("RE-BAG -> BULK"))
   C.ItemData(4) = 18
   
'   C.AddItem (MapText("RE-BAG -> RM(OTHER)"))
'   C.ItemData(5) = 19
End Sub

'===
Public Sub InitLoadPartType2(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("อาหาร BAG"))
   C.ItemData(1) = 14
   
   C.AddItem (MapText("อาหาร BULK"))
   C.ItemData(2) = 13

End Sub

'===
Public Sub InitLoadBalanceType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มียอดยกมา"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("มียอดรับเข้า"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("มียอดจ่ายออก"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("มียอดคงเหลือ"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("ไม่มียอดคงเหลือ"))
   C.ItemData(5) = 5
   
   C.AddItem (MapText("มียอดเคลื่อนไหว"))
   C.ItemData(6) = 6
   
   C.ListIndex = 6

End Sub

Public Sub InitLoadPartPayType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("สินค้า BAG"))
   C.ItemData(1) = 2000
   
   C.AddItem (MapText("สินค้า BULK"))
   C.ItemData(2) = 2001
   
End Sub
Public Sub InitLoadPartPayType2(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ขายอาหาร BAG"))
   C.ItemData(1) = 2000
   
   C.AddItem (MapText("ขายอาหาร BULK"))
   C.ItemData(2) = 2001
   
   C.AddItem (MapText("อาหาร RE-BAG"))
   C.ItemData(3) = 2002
   
   C.AddItem (MapText("บรรจุ BULK TO BAG"))
   C.ItemData(4) = 2003
   
   C.AddItem (MapText("ขายอาหาร อื่นๆ"))
   C.ItemData(5) = 2004
   
End Sub
Public Sub InitLoadProcessType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ขึ้นอาหารเรียบร้อย"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("รอขึ้นอาหาร"))
   C.ItemData(2) = 2
   
   C.ListIndex = 1
End Sub

'===
'
Public Sub InitFeatureOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อสินค้า"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitSocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขแพคเกจ"))
   C.ItemData(1) = 1
End Sub

Public Sub InitPackageOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขแพคเกจ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่ประกาศ"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันที่มีผล"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("วันที่สิ้นสุด"))
   C.ItemData(4) = 4
End Sub
'===

Public Sub InitGoldDailyPriceOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ราคาทอง"))
   C.ItemData(1) = 1
End Sub

Public Sub InitGoldSellBuyOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่บิล"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เลขที่บิล"))
   C.ItemData(2) = 2
End Sub

Public Sub InitGoldWeightOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสทอง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อทอง"))
   C.ItemData(2) = 2
End Sub

'===

Public Sub InitInventoryDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่บิลรับของ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่รับวัตถุดิบ"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitBillingDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(3) = 3
End Sub

Public Sub InitBillingPaymentOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร PV"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("เลขที่เอกสาร JV"))
   C.ItemData(2) = 2

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(3) = 3

End Sub
'===
Public Sub InitBillingDocOtherFilterOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PO ที่ยังไม่อนุมัติ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("PO ที่อนุมัติแล้ว"))
   C.ItemData(2) = 2

   C.AddItem (MapText("เอกสารที่สร้างโดยไม่มี PO และยังไม่อนุมัติ"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("เอกสารที่สร้างโดยไม่มี PO และอนุมัติแล้ว"))
   C.ItemData(4) = 4
End Sub
'===
Public Sub InitBillingDocApproved(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PO ที่ยังไม่อนุมัติ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("PO ที่อนุมัติแล้ว"))
   C.ItemData(2) = 2
   
End Sub
'===
Public Sub InitBillingDocCloseApproved(C As ComboBox, Optional Index As Integer = 0)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PO ที่ยังไม่ปิด"))
   C.ItemData(1) = 1

   C.AddItem (MapText("PO ที่ปิดแล้ว"))
   C.ItemData(2) = 2
   
   C.ListIndex = Index
End Sub
'===

Public Sub InitChequeDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(3) = 3
End Sub
Public Sub InitReport5_1OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2

   C.AddItem (MapText("รหัสสินค้า"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ชื่อสินค้า"))
   C.ItemData(4) = 4
End Sub

Public Sub InitReport5_5_1OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub
'===


Public Sub InitPigDoc1OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ใบเกิด"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เกิด"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitInventoryDoc2OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบเบิก"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เบิก"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitInventoryDoc3OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบโอน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่โอน"))
   C.ItemData(2) = 2
End Sub
'===

Public Sub InitPigDoc3OrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบปรับยอด"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่ปรับยอด"))
   C.ItemData(2) = 2
End Sub
'===
Public Sub initWeightPerPack(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("30"))
   C.ItemData(1) = 1

   C.AddItem (MapText("50"))
   C.ItemData(2) = 2
   
   C.ListIndex = 1
End Sub
Public Sub initSewingThread(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("ขาว"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ขาว-แดง"))
   C.ItemData(2) = 2
   
   C.ListIndex = 1
End Sub
Public Sub InitPigWeekOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปี"))
   C.ItemData(1) = 1
End Sub
'===

Public Sub LoadUserGroup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserGroup
Dim I As Long

   Set D = New CUserGroup
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_NAME)
         C.ItemData(I) = TempData.GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.GROUP_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadCountry(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCountry
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCountry
Dim I As Long

   Set D = New CCountry
   Set Rs = New ADODB.Recordset
   
   D.COUNTRY_ID = -1
   D.CONTINENT_ID = -1
   D.COUNTRY_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCountry
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.COUNTRY_NAME)
         C.ItemData(I) = TempData.COUNTRY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.COUNTRY_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLocationCodeKey(C As ComboBox, Optional Cl As Collection = Nothing, Optional Area As Long = -1, Optional SaleFlag As String = "N", Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLocation
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLocation
Dim I As Long

   Set D = New CLocation
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
'   D.LOCATION_TYPE = Area
'   D.SALE_FLAG = SaleFlag
   D.OrderBy = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLocation
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOCATION_NAME)
         C.ItemData(I) = TempData.LOCATION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & " มีรหัสคลัง " & TempData.LOCATION_NO & " ซ้ำ."
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional Area As Long = -1, Optional SaleFlag As String = "N", Optional LocationID As Long = -1, Optional PartGroupID As Long = -1, Optional KeyType As Long = 1, Optional LocationGroup As String = "")
On Error GoTo ErrorHandler
Dim D As CLocation
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLocation
Dim Area1 As Collection
Dim Area2 As Collection
Dim I As Long
   Set D = New CLocation
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
   D.LOCATION_TYPE = Area
   D.SALE_FLAG = SaleFlag
   D.PART_GROUP_ID = PartGroupID
   D.KEY_CODE = LocationGroup
   D.OrderBy = 1
   If LocationID = -2 Then
      D.LOCATION_ID_SET = "LOCATION_ID=128 OR LOCATION_ID=129 OR LOCATION_ID=130 OR LOCATION_ID=131 OR LOCATION_ID=132"
   ElseIf LocationID = -3 Then
      D.LOCATION_ID_SET = "LOCATION_ID=133 OR LOCATION_ID=134 OR LOCATION_ID=135 OR LOCATION_ID=136 OR LOCATION_ID=137 OR LOCATION_ID=138 OR LOCATION_ID=139 OR LOCATION_ID=140 OR LOCATION_ID=141"
   End If

   Call D.QueryData(Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLocation
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOCATION_NAME)
         C.ItemData(I) = TempData.LOCATION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.LOCATION_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.LOCATION_NO))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPalletDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional LOT_ID As Long = -1, Optional KeyType As Long = 1, Optional NotIn As String = "", Optional Ind As Long, Optional TxType As String, Optional INVENTORY_WH_DOC_ID As Long = -1, Optional LotDocId As Long = -1, Optional HeadPackNo As Long = -1, Optional PartItemID As Long = -1, Optional DocumentType As Long = -1, Optional LotDocIdRef As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim PD As CPalletDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPalletDoc
Dim I As Long

   Set PD = New CPalletDoc
   Set Rs = New ADODB.Recordset
   
   
   PD.LOT_ITEM_WH_ID = -1
   PD.LOT_ID = LOT_ID
   PD.LOT_DOC_ID = LotDocId
   PD.LOT_DOC_ID_REF = LotDocIdRef
   PD.HEAD_PACK_NO = HeadPackNo
   PD.INVENTORY_WH_DOC_ID = INVENTORY_WH_DOC_ID
   PD.NOT_PALLET_DOC_ID = NotIn
   PD.TX_TYPE = TxType
   PD.BALANCE_FLAG = "N"
   PD.PART_ITEM_ID = PartItemID
   PD.LOCATION_ID = LocationID
   
   If DocumentType = 13 Or DocumentType = 16 Or DocumentType = 18 Or DocumentType = 20 Then
        PD.DOCUMENT_TYPE_SET = "(13,16,18,20)"
   ElseIf DocumentType = 14 Or DocumentType = 15 Or DocumentType = 17 Or DocumentType = 19 Then
        PD.DOCUMENT_TYPE_SET = "(14,15,17,19)"
   End If

   PD.OrderBy = 2
   PD.OrderType = 1
   Call PD.QueryData(Ind, Rs, ItemCount)
   Set PD = Nothing
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set PD = New CPalletDoc
      Call PD.PopulateFromRS(Ind, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (PD.PALLET_DOC_NO)
         C.ItemData(I) = PD.PALLET_DOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(PD, Trim(str(PD.PALLET_DOC_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(PD, Trim(PD.PALLET_DOC_NO))
         ElseIf KeyType = 3 Then
            If PD.CAPACITY_AMOUNT > 0 Then 'กรณีมีการปรับยอดที่ปรับค่าให้เป็น 0 ไม่ต้องเอามา
               Call Cl.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(PD.LOT_ID) & "-" & str(PD.HEAD_PACK_NO) & "-" & str(PD.PART_ITEM_ID) & "-" & PD.TX_TYPE & str(PD.BIN_NO)))       'LOT_DOC_ID
           End If
         ElseIf KeyType = 4 Then
           Call Cl.add(PD, Trim(LOT_ID & "-" & PD.TX_TYPE & "-" & PD.HEAD_PACK_NO))
         ElseIf KeyType = 5 Then 'ถ้าเป็นการจ่ายออก ให้เอา LOT_DOC_ID_REF มาเป็น Key
           Set TempData = GetObject("CPalletDoc", Cl, Trim(PD.PALLET_DOC_NO & "-" & str(PD.LOT_ID) & "-" & PD.TX_TYPE & str(PD.BIN_NO)), False)
           If Not TempData Is Nothing Then
               TempData.CAPACITY_AMOUNT = TempData.CAPACITY_AMOUNT + PD.CAPACITY_AMOUNT
            Else
               Call Cl.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(PD.LOT_ID) & "-" & PD.TX_TYPE & str(PD.BIN_NO)))
           End If
          ElseIf KeyType = 6 Then 'สำหรับ Bulk
            Call Cl.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(PD.LOT_DOC_ID) & "-" & "-" & str(PD.PART_ITEM_ID) & "-" & PD.TX_TYPE & str(PD.BIN_NO)))
          ElseIf KeyType = 7 Then 'สำหรับ การแก้ไข pallet กรณีเปลี่ยนชื่อแล้ว PALLET ซ้ำกัน I
            Call Cl.add(PD, Trim(PD.PALLET_DOC_NO & "-" & str(PD.LOT_DOC_ID) & "-" & str(PD.LOT_ID) & PD.TX_TYPE))
         ElseIf KeyType = 8 Then  'สำหรับไว้แสดงผลอย่างเดียว ไม่ต้องใส่ Key
            Call Cl.add(PD)
         End If
      End If
      Set PD = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set PD = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPalletNoFlagYLastBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional LOT_ID As Long = -1, Optional TxType As String)
On Error GoTo ErrorHandler
Dim PD As CPalletDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPalletDoc
Dim I As Long

   Set PD = New CPalletDoc
   Set Rs = New ADODB.Recordset
   
   PD.LOT_ID = LOT_ID
   PD.TX_TYPE = TxType
   PD.BALANCE_FLAG = "Y"
   PD.PALLET_DOC_NO = "999"
   PD.OrderBy = 2
   PD.OrderType = 1
   Call PD.QueryData(8, Rs, ItemCount)
   Set PD = Nothing
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set PD = New CPalletDoc
      Call PD.PopulateFromRS(8, Rs)
   
'      If Not (C Is Nothing) Then
'         C.AddItem (PD.PALLET_DOC_NO)
'         C.ItemData(I) = PD.PALLET_DOC_ID
'      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(PD, Trim(str(PD.INVENTORY_WH_DOC_ID)))
      End If
      Set PD = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set PD = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPalletByLotID(C As ComboBox, Optional Cl As Collection = Nothing, Optional LOT_ID As Long = -1, Optional TxType As String)
On Error GoTo ErrorHandler
Dim PD As CPalletDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPalletDoc
Dim I As Long

   Set PD = New CPalletDoc
   Set Rs = New ADODB.Recordset
   
   PD.LOT_ID = LOT_ID
   PD.TX_TYPE = TxType
   PD.BALANCE_FLAG = "N"
   PD.OrderBy = 2
   PD.OrderType = 1
   Call PD.QueryData(10, Rs, ItemCount)
   Set PD = Nothing
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set PD = New CPalletDoc
      Call PD.PopulateFromRS(10, Rs)
   
'      If Not (C Is Nothing) Then
'         C.AddItem (PD.PALLET_DOC_NO)
'         C.ItemData(I) = PD.PALLET_DOC_ID
'      End If
      
      If Not (Cl Is Nothing) Then
         Set TempData = GetObject("CPalletDoc", Cl, Trim(PD.PALLET_DOC_NO) & "-" & Trim(str(PD.HEAD_PACK_NO)), False)
         If TempData Is Nothing Then
            Call Cl.add(PD, Trim(PD.PALLET_DOC_NO) & "-" & Trim(str(PD.HEAD_PACK_NO)))
         End If
      End If
      Set PD = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set PD = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPalletDocAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional KeyType As Long = 1, Optional NotIn As String = "", Optional Ind As Long, Optional TxType As String, Optional INVENTORY_WH_DOC_ID As Long = -1, Optional Cl2 As Collection = Nothing, Optional LotDocId As Long = -1, Optional HEAD_PACK_NO As Long = -1, Optional LOT_ITEM_WH_ID As Long = -1, Optional DOCUMENT_TYPE As Long = -1, Optional PartItemID As Long = -1, Optional LotDocIdRef As Long = -1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPalletDoc
Dim TempI As Collection
Dim TempE As Collection
Dim LTD As CLotDoc
Dim dTempI As CPalletDoc
Dim dTempE As CPalletDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPalletDoc
Dim I As Long
Dim SumPalletAmount As Double
Dim SumPayOff As Long

   Set D = New CPalletDoc
   Set Rs = New ADODB.Recordset
   Set TempI = New Collection
   Set TempE = New Collection
   
   If DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 16 Or DOCUMENT_TYPE = 18 Then
      Call LoadPalletDoc(Nothing, TempI, LotId, 6, , 5, "I", , , , PartItemID, DOCUMENT_TYPE, , LocationID)
   ElseIf DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 15 Or DOCUMENT_TYPE = 17 Then
      Call LoadPalletDoc(Nothing, TempI, LotId, 3, , 5, "I", , , , PartItemID, DOCUMENT_TYPE, , LocationID)
   End If
   
   Call LoadPalletDoc(Nothing, TempE, LotId, 5, , 5, "E", , , , PartItemID, , LotDocIdRef, LocationID)
   
   D.LOT_ITEM_WH_ID = -1
   D.LOT_ID = LotId
   D.INVENTORY_WH_DOC_ID = INVENTORY_WH_DOC_ID
   D.NOT_PALLET_DOC_ID = NotIn
   D.TX_TYPE = "I"
   D.HEAD_PACK_NO = HEAD_PACK_NO
   D.LOT_ITEM_WH_ID = LOT_ITEM_WH_ID
   
   D.LOT_DOC_ID = LotDocId
   D.BALANCE_FLAG = "N"
   D.PART_ITEM_ID = PartItemID

   D.OrderBy = 2
   D.OrderType = 1
   Call D.QueryData(Ind, Rs, ItemCount)
   
 If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl2 Is Nothing) Then
      Set Cl2 = Nothing
      Set Cl2 = New Collection
   End If
   
   
   SumPayOff = 0
 
 Set LTD = New CLotDoc
   While Not Rs.EOF
      I = I + 1
      SumPalletAmount = 0
      Set TempData = New CPalletDoc
      Call TempData.PopulateFromRS(Ind, Rs)
       If DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 16 Or DOCUMENT_TYPE = 18 Then
         Set dTempI = GetObject("CPalletDoc", TempI, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_DOC_ID) & "-" & "-" & str(TempData.PART_ITEM_ID) & "-" & "I" & str(TempData.BIN_NO)), False)
      ElseIf DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 15 Or DOCUMENT_TYPE = 17 Then
         Set dTempI = GetObject("CPalletDoc", TempI, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_ID) & "-" & str(TempData.HEAD_PACK_NO) & "-" & str(TempData.PART_ITEM_ID) & "-" & "I" & str(TempData.BIN_NO)), False)
      Else
         Set dTempI = GetObject("CPalletDoc", TempI, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_ID) & "-" & str(TempData.HEAD_PACK_NO) & "-" & str(TempData.PART_ITEM_ID) & "-" & "I" & str(TempData.BIN_NO)), False)
      End If
      If Not dTempI Is Nothing Then
         LTD.LOT_AMOUNT = LTD.LOT_AMOUNT + dTempI.CAPACITY_AMOUNT
         LTD.LOT_BAL = LTD.LOT_BAL + dTempI.CAPACITY_AMOUNT
         SumPalletAmount = SumPalletAmount + dTempI.CAPACITY_AMOUNT
      End If
      'เวลาหา TempE ให้ใช้ LOT_DOC_ID เพื่อไปหา LOT_DOC_ID_REF
      Set dTempE = GetObject("CPalletDoc", TempE, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_ID) & "-" & "E" & str(TempData.BIN_NO)), False)  'LOT_DOC_ID
      If Not dTempE Is Nothing Then
         LTD.LOT_PAYOFF = LTD.LOT_PAYOFF + dTempE.CAPACITY_AMOUNT
         LTD.LOT_BAL = LTD.LOT_BAL - dTempE.CAPACITY_AMOUNT
         SumPalletAmount = SumPalletAmount - dTempE.CAPACITY_AMOUNT
      End If
       If (SumPalletAmount > 0) Then 'And (TempData.LOT_DOC_ID = LotDocId)
        ' TempData.CAPACITY_AMOUNT = SumPalletAmount 'ให้เอายอดคงเหลือของแต่ล่ะพาเลทมาเก็บไว้เพื่อนำไปแสดงใน คอมโบบ๊อก
         TempData.PALLET_CAP_LAST = SumPalletAmount 'ให้จำค่าคงเหลือล่าสุดไว้
         If Not (C Is Nothing) Then
               C.AddItem (TempData.PALLET_DOC_NO)
               C.ItemData(I) = TempData.PALLET_DOC_ID
         End If

         If Not (Cl Is Nothing) Then
               If KeyType = 1 Then
                  Call Cl.add(TempData, Trim(str(TempData.PALLET_DOC_ID)))
               ElseIf KeyType = 2 Then
                  Call Cl.add(TempData, Trim(TempData.PALLET_DOC_NO & "-" & TempData.HEAD_PACK_NO))
               ElseIf KeyType = 3 Then
                  Call Cl.add(TempData, Trim(TempData.PALLET_DOC_NO) & "-" & Trim(str(TempData.LOT_ID)) & "-" & Trim(str(TempData.TX_TYPE))) 'LOT_DOC_ID
               End If
         End If
      Else
               I = I - 1
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
    If Not (Cl2 Is Nothing) Then 'เก็บจำนวนรวมของทุก pallet ในแต่ละ Lot
      Call Cl2.add(LTD, Trim(str(LotId)))
      Set LTD = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPalletDocAmountAll(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional KeyType As Long = 1, Optional NotIn As String = "", Optional Ind As Long, Optional TxType As String, Optional INVENTORY_WH_DOC_ID As Long = -1, Optional Cl2 As Collection = Nothing, Optional LotDocId As Long = -1, Optional HEAD_PACK_NO As Long = -1, Optional LOT_ITEM_WH_ID As Long = -1, Optional DOCUMENT_TYPE As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPalletDoc
Dim TempI As Collection
Dim TempE As Collection
Dim LTD As CLotDoc
Dim dTempI As CPalletDoc
Dim dTempE As CPalletDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPalletDoc
Dim I As Long
Dim SumPalletAmount As Double
Dim SumPayOff As Long

   Set D = New CPalletDoc
   Set Rs = New ADODB.Recordset
   Set TempI = New Collection
   Set TempE = New Collection
   
   If DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 16 Or DOCUMENT_TYPE = 18 Then
      Call LoadPalletDoc(Nothing, TempI, LotId, 6, , 5, "I", , , , , DOCUMENT_TYPE)
   ElseIf DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 15 Or DOCUMENT_TYPE = 17 Then
      Call LoadPalletDoc(Nothing, TempI, LotId, 3, , 5, "I", , , , , DOCUMENT_TYPE)
   End If
   
   Call LoadPalletDoc(Nothing, TempE, LotId, 5, , 5, "E")
   
   D.LOT_ITEM_WH_ID = -1
   D.LOT_ID = LotId
   D.INVENTORY_WH_DOC_ID = INVENTORY_WH_DOC_ID
   D.NOT_PALLET_DOC_ID = NotIn
   D.TX_TYPE = "I"
   D.HEAD_PACK_NO = HEAD_PACK_NO
   D.LOT_ITEM_WH_ID = LOT_ITEM_WH_ID

   D.OrderBy = 2
   D.OrderType = 1
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl2 Is Nothing) Then
      Set Cl2 = Nothing
      Set Cl2 = New Collection
   End If
   
   
   SumPayOff = 0
 
 Set LTD = New CLotDoc
   While Not Rs.EOF
      I = I + 1
      SumPalletAmount = 0
      Set TempData = New CPalletDoc
      Call TempData.PopulateFromRS(Ind, Rs)
       If DOCUMENT_TYPE = 13 Or DOCUMENT_TYPE = 18 Then
         Set dTempI = GetObject("CPalletDoc", TempI, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_DOC_ID) & "-" & "-" & str(TempData.PART_ITEM_ID) & "-" & "I" & str(TempData.BIN_NO)), False)
      ElseIf DOCUMENT_TYPE = 14 Or DOCUMENT_TYPE = 17 Then
         Set dTempI = GetObject("CPalletDoc", TempI, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_ID) & "-" & str(TempData.HEAD_PACK_NO) & "-" & str(TempData.PART_ITEM_ID) & "-" & "I" & str(TempData.BIN_NO)), False)
      Else
         Set dTempI = GetObject("CPalletDoc", TempI, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_ID) & "-" & str(TempData.HEAD_PACK_NO) & "-" & str(TempData.PART_ITEM_ID) & "-" & "I" & str(TempData.BIN_NO)), False)
      End If
      If Not dTempI Is Nothing Then
         LTD.LOT_AMOUNT = LTD.LOT_AMOUNT + dTempI.CAPACITY_AMOUNT
         LTD.LOT_BAL = LTD.LOT_BAL + dTempI.CAPACITY_AMOUNT
         SumPalletAmount = SumPalletAmount + dTempI.CAPACITY_AMOUNT
      End If
      'เวลาหา TempE ให้ใช้ LOT_DOC_ID เพื่อไปหา LOT_DOC_ID_REF
      Set dTempE = GetObject("CPalletDoc", TempE, Trim(TempData.PALLET_DOC_NO & "-" & str(TempData.LOT_ID) & "-" & "E" & str(TempData.BIN_NO)), False)  'LOT_DOC_ID
      If Not dTempE Is Nothing Then
         LTD.LOT_PAYOFF = LTD.LOT_PAYOFF + dTempE.CAPACITY_AMOUNT
         LTD.LOT_BAL = LTD.LOT_BAL - dTempE.CAPACITY_AMOUNT
         SumPalletAmount = SumPalletAmount - dTempE.CAPACITY_AMOUNT
      End If
       If (SumPalletAmount > 0) Then 'And (TempData.LOT_DOC_ID = LotDocId)
        ' TempData.CAPACITY_AMOUNT = SumPalletAmount 'ให้เอายอดคงเหลือของแต่ล่ะพาเลทมาเก็บไว้เพื่อนำไปแสดงใน คอมโบบ๊อก
         TempData.PALLET_CAP_LAST = SumPalletAmount 'ให้จำค่าคงเหลือล่าสุดไว้
         If Not (C Is Nothing) Then
               C.AddItem (TempData.PALLET_DOC_NO)
               C.ItemData(I) = TempData.PALLET_DOC_ID
         End If

         If Not (Cl Is Nothing) Then
               If KeyType = 1 Then
                  Call Cl.add(TempData, Trim(str(TempData.PALLET_DOC_ID)))
               ElseIf KeyType = 2 Then
'                  Call Cl.add(TempData, Trim(TempData.PALLET_DOC_NO & "-" & TempData.HEAD_PACK_NO))
                   Call Cl.add(TempData, Trim(TempData.PALLET_DOC_NO & "-" & TempData.HEAD_PACK_NO & "-" & TempData.LOT_ID & "-" & TempData.LOT_DOC_ID & "-" & TempData.LOT_ITEM_WH_ID))
               ElseIf KeyType = 3 Then
                  Call Cl.add(TempData, Trim(TempData.PALLET_DOC_NO) & "-" & Trim(str(TempData.LOT_ID)) & "-" & Trim(str(TempData.TX_TYPE))) 'LOT_DOC_ID
               End If
         End If
      Else
               I = I - 1
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
    If Not (Cl2 Is Nothing) Then 'เก็บจำนวนรวมของทุก pallet ในแต่ละ Lot
      Call Cl2.add(LTD, Trim(str(LotId)))
      Set LTD = Nothing
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPalletDocAmount2(C As ComboBox, Optional Cl As Collection, Optional LotId As Long = -1, Optional KeyType As Long = 1, Optional NotIn As String = "", Optional Ind As Long, Optional TxType As String, Optional INVENTORY_WH_DOC_ID As Long = -1, Optional LotDocId As Long = -1, Optional ColTempPD As Collection = Nothing)
On Error GoTo ErrorHandler
Dim TempPD As CPalletDoc
Dim PD As CPalletDoc
Dim ItemCount As Long
Dim TempData As CPalletDoc
Dim I As Long
Dim SumPalletAmount As Long

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   For Each PD In Cl
      I = I + 1
      SumPalletAmount = 0
      Set TempData = New CPalletDoc
      SumPalletAmount = PD.CAPACITY_AMOUNT
       If SumPalletAmount > 0 Then
        TempData.CAPACITY_AMOUNT = SumPalletAmount 'ให้เอายอดคงเหลือของแต่ล่ะพาเลทมาเก็บไว้เพื่อนำไปแสดงใน คอมโบบ๊อก
         If PD.TEMP_PALLET_CAP_LAST > 0 Then
            If Not (C Is Nothing) Then
                  C.AddItem (PD.PALLET_DOC_NO)
                  C.ItemData(I) = PD.PALLET_DOC_ID
            End If
         Else
             I = I - 1
         End If
      Else
               I = I - 1
      End If
   
      Set TempData = Nothing
   Next PD
     If C.ListCount > 1 Then
         C.ListIndex = 1
      End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLotByPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional TxType As String = "", Optional Cl2 As Collection = Nothing, Optional PackAmount As Long = 0, Optional DOCUMENT_TYPE As Long, Optional AdjusFlag As Boolean = False, Optional LocationID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long
Dim CLotI As Collection
Dim CLotE As Collection
Dim TempLotDoc As CLotDoc
Dim SumLotAmount As Long
Dim m_CollPallet As Collection
Dim TotalAmount As Double

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset
   Set CLotI = New Collection
   Set CLotE = New Collection

   Call LoadLotFromLotDoc(Nothing, CLotI, , FromDate, ToDate, 1, PartItemID, 3, , , "I", , , LocationID)
   Call LoadLotFromLotDoc(Nothing, CLotE, , FromDate, ToDate, 1, PartItemID, 3, , , "E", , , LocationID)
   
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.LOT_ID = LotId
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 2 'OrderBy
   D.OrderType = OrderType
   D.TX_TYPE = TxType
   D.BALANCE_FLAG = "N"
   D.OUT_STOCK_FLAG = "N"
'   If DOCUMENT_TYPE = 14 Then
'      D.VERIFY_FLAG = "Y" 'จะดึงเฉพาะตัวที่ผ่านการตรวจสอบแล้วเท่านั้น
'      Call D.QueryData(12, Rs, ItemCount)
'   Else
'      Call D.QueryData(Ind, Rs, ItemCount)
'   End If
   Call D.QueryData(Ind, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
Dim LTD As CLotDoc
Dim Flag As Boolean
Dim CJ As CJob
Dim m_CollJob As Collection
Set m_CollJob = New Collection
 Call LoadJobByInventoryWhDoc(Nothing, m_CollJob, , , , , DOCUMENT_TYPE)
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(Ind, Rs)
      
      Set CJ = GetObject("CJob", m_CollJob, str(TempData.INVENTORY_WH_DOC_ID), False) 'ถ้ามีแสดงว่า ผ่านการตรวจสอบแล้ว
      If Not (CJ Is Nothing) Then
   
      Set LTD = GetObject("CLotDoc", CLotI, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "I"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = LTD.LOT_AMOUNT
      End If
      Set LTD = GetObject("CLotDoc", CLotE, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "E"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = TempData.LOT_AMOUNT - LTD.LOT_AMOUNT
      End If
      
      
      SumLotAmount = TempData.LOT_AMOUNT
      Flag = True
      If Not Cl2 Is Nothing Then
         For Each LTD In Cl2
            If LTD.LOT_DOC_ID_REF = TempData.LOT_DOC_ID And LTD.Flag <> "D" Then
               Flag = False
            End If
         Next LTD
      End If
      If Flag = True Then
           TotalAmount = TempData.LOT_AMOUNT  'GetTotalAmountPallet(m_CollPallet)
          If Val(TotalAmount) > 0 Then   'ถ้าไม่มีข้อมูลใน pallet แล้วก็ไม่ต้องให้แสดง Lot'If TempData.LOT_AMOUNT > 0 Then
            If SumLotAmount <= PackAmount Then
               SumLotAmount = SumLotAmount + TempData.LOT_AMOUNT
               If Not (C Is Nothing) Then
                     C.AddItem (TempData.LOT_NO & "-" & Format(TempData.TIME_PACK_BEGIN, "HH:mm") & " " & TempData.BIN_NAME & " " & TempData.LOCK_NAME)
                     C.ItemData(I) = TempData.LOT_DOC_ID
               End If
               If Not (Cl Is Nothing) Then
                  If KeyType = 1 Then
                        Call Cl.add(TempData, Trim(str(TempData.LOT_ID) & "-" & str(TempData.LOT_DOC_ID)))
                  End If
               End If
            ElseIf SumLotAmount > PackAmount Then
               If Not (C Is Nothing) Then
                  C.AddItem (TempData.LOT_NO & "-" & Format(TempData.TIME_PACK_BEGIN, "HH:mm") & " " & TempData.BIN_NAME & " " & TempData.LOCK_NAME)
                   C.ItemData(I) = TempData.LOT_ID & TempData.LOT_DOC_ID 'ตัวนี้จำเป็นต้องให้ key ติดกัน
               End If

               If Not (Cl Is Nothing) Then
                  If KeyType = 1 Then
                        TempData.LOT_DOC_ID_REF = TempData.LOT_DOC_ID
                        Call Cl.add(TempData, Trim(str(TempData.LOT_ID & TempData.LOT_DOC_ID))) 'ตัวนี้จำเป็นต้องให้ key ติดกัน
                  End If
               End If
            End If
         Else
         I = I - 1
         
         If TotalAmount = 0 Then
            Call UpdateOutStockFlagInLIW(TempData.LOT_ITEM_WH_ID, "Y") 'หาก lot นั้นๆหมดแล้ว ก็ให้เปลี่ยน สถานะเลย
         End If
         End If
      Else
       I = I - 1
      End If
      
      Else
          I = I - 1
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function CalAdjustByPartItem(PartItemID As Long, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional TxType As String = "", Optional AdjusFlag As Boolean = False) As Boolean
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long
Dim CLotI As Collection
Dim CLotE As Collection
Dim TempLotDoc As CLotDoc
Dim SumLotAmount As Long
Dim m_CollPallet As Collection
Dim TotalAmount As Double
Dim Flag As Boolean

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset
   Set CLotI = New Collection
   Set CLotE = New Collection

   Call LoadLotFromLotDoc(Nothing, CLotI, , , , 1, PartItemID, 3, , , "I", , AdjusFlag)
   Call LoadLotFromLotDoc(Nothing, CLotE, , , , 1, PartItemID, 3, , , "E", , AdjusFlag)
   
   D.PART_ITEM_ID = PartItemID
   D.OrderBy = 2 'OrderBy
   D.OrderType = OrderType
   D.TX_TYPE = TxType
   D.BALANCE_FLAG = "N"
   Call D.QueryData(Ind, Rs, ItemCount)
   
  Dim LTD As CLotDoc
   While Not Rs.EOF
   DoEvents
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(Ind, Rs)
      Set LTD = GetObject("CLotDoc", CLotI, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "I"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = LTD.LOT_AMOUNT
      End If
      Set LTD = GetObject("CLotDoc", CLotE, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "E"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = TempData.LOT_AMOUNT - LTD.LOT_AMOUNT
      End If
      
      
      SumLotAmount = TempData.LOT_AMOUNT
      
      If SumLotAmount > 0 Then
         Call UpdateOutStockFlagInLIW(TempData.LOT_ITEM_WH_ID, "N") 'หาก lot นั้นยังมีก็ให้เปลี่ยนเป็น N
      Else
         Call UpdateOutStockFlagInLIW(TempData.LOT_ITEM_WH_ID, "Y") 'หาก lot นั้นๆหมดแล้ว ก็ให้เปลี่ยน สถานะเลย
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   CalAdjustByPartItem = True
   Set Rs = Nothing
   Set D = Nothing
   Exit Function

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   CalAdjustByPartItem = False
End Function

Public Sub LoadLotFIFOByPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional TxType As String = "", Optional Cl2 As Collection = Nothing, Optional PackAmount As Long = 0, Optional DOCUMENT_TYPE As Long = 0, Optional LIW As CLotItemWH = Nothing)
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long
Dim CLotI As Collection
Dim CLotE As Collection
Dim TempLotDoc As CLotDoc
Dim SumLotAmount As Double
Dim DiffAmount As Double
Dim m_CollPallet As Collection
Dim DOCUMENT_TYPE_INPUT As Long
Dim Flag As Boolean
Dim TotalAmount As Double

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset
   Set CLotI = New Collection
   Set CLotE = New Collection

   Call LoadLotFromLotDoc(Nothing, CLotI, , , , 1, PartItemID, 3, , , "I", , , LIW.LOCATION_ID)
   Call LoadLotFromLotDoc(Nothing, CLotE, , , , 1, PartItemID, 10, , , "E", , , LIW.LOCATION_ID)

   D.PART_ITEM_ID = PartItemID
   D.LOT_ID = LotId
   D.LOCATION_ID = LIW.LOCATION_ID
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 2 'OrderBy
   D.OrderType = OrderType
   D.TX_TYPE = TxType
   D.BALANCE_FLAG = "N" 'เลือกเฉพาะที่ยังไม่ปรับยอด
   D.OUT_STOCK_FLAG = "N" 'เลือกเฉพาะ LotItemWh ที่ ยังมียอด
'   If DOCUMENT_TYPE = 2000 Then
'      D.VERIFY_FLAG = "Y" 'จะดึงเฉพาะตัวที่ผ่านการตรวจสอบแล้วเท่านั้น
'      Call D.QueryData(12, Rs, ItemCount)
'   Else
'      Call D.QueryData(Ind, Rs, ItemCount)
'   End If
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

PackAmount = PackAmount - GetTotalAmount(Cl2) 'เอายอดแพ็คใหม่ลบกับยอดแพ็คเดิมก่อน กรณีเผื่อเป็นการแก้ไขยอดใหม่จาก SO

Dim LTD As CLotDoc
Dim LTD2  As CLotDoc
Dim LastLotAmount As Double
Dim CJ As CJob
Dim m_CollJob As Collection
Set m_CollJob = New Collection
 Call LoadJobByInventoryWhDoc(Nothing, m_CollJob, , , , , DOCUMENT_TYPE)
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(Ind, Rs)

      Set CJ = GetObject("CJob", m_CollJob, str(TempData.INVENTORY_WH_DOC_ID), False) 'ถ้ามีแสดงว่า ผ่านการตรวจสอบแล้ว
      If Not (CJ Is Nothing) Then
      
            Set LTD = GetObject("CLotDoc", CLotI, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "I"), False)
            If Not (LTD Is Nothing) Then
               TempData.LOT_AMOUNT = LTD.LOT_AMOUNT
            End If
            Set LTD = GetObject("CLotDoc", CLotE, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "E"), False)
            If Not (LTD Is Nothing) Then
               TempData.LOT_AMOUNT = TempData.LOT_AMOUNT - LTD.LOT_AMOUNT
            End If
      
            LastLotAmount = 0
            Flag = True
            For Each LTD In Cl2 'เช็คว่า เคยมี lot ใดเคย load ไปแล้วบ้าง
               If LTD.LOT_DOC_ID_REF = TempData.LOT_DOC_ID Then
                  Flag = False
               Else 'หาว่า lot ที่เคย load ก่อนหน้ามีจำนวนเพียงพอหรือยังถ้าพอแล้วก็ไม่ต้องเพิ่ม lot ของ รายการนั้นๆอีก
                  Set LTD2 = GetObject("CLotDoc", CLotI, Trim(str(LTD.LOT_ID) & "-" & Trim(str(LTD.LOT_DOC_ID_REF)) & "-" & "I"), False)
                  If Not (LTD2 Is Nothing) Then
                  LTD.LOT_AMOUNT = LTD2.LOT_AMOUNT
                  End If
                  
                  Set LTD2 = GetObject("CLotDoc", CLotE, Trim(str(LTD.LOT_ID) & "-" & Trim(str(LTD.LOT_DOC_ID_REF)) & "-" & "E"), False)
                  If Not (LTD2 Is Nothing) Then
                  LTD.LOT_AMOUNT = LTD.LOT_AMOUNT - LTD2.LOT_AMOUNT
                  End If
                  LastLotAmount = LastLotAmount + LTD.LOT_AMOUNT
               End If
            Next LTD
            
            If Flag And LastLotAmount <= PackAmount Then
             If Val(TempData.LOT_AMOUNT) > 0 Then 'ถ้าไม่มีข้อมูลใน pallet แล้วก็ไม่ต้องให้แสดง Lot
                  If SumLotAmount <= PackAmount Then
                     SumLotAmount = SumLotAmount + TempData.LOT_AMOUNT
                     
                       If Not (C Is Nothing) Then
                         C.AddItem (TempData.LOT_NO & "-" & Format(TempData.TIME_PACK_BEGIN, "HH:mm"))
                         C.ItemData(I) = TempData.LOT_ID & TempData.LOT_DOC_ID 'ตัวนี้จำเป็นต้องให้ key ติดกัน
                     End If
                     
                     If Not (Cl Is Nothing) Then
                        If KeyType = 1 Then
                              TempData.LOT_DOC_ID_REF = TempData.LOT_DOC_ID
                              Call Cl.add(TempData, Trim(str(TempData.LOT_ID & TempData.LOT_DOC_ID))) 'ตัวนี้จำเป็นต้องให้ key ติดกัน
                        End If
                     End If
                     
                  ElseIf SumLotAmount > PackAmount Then
                     Set TempData = Nothing
                     Set Rs = Nothing
                     Set D = Nothing
                     Exit Sub
                  End If
               Else
               I = I - 1
               If Val(TempData.LOT_AMOUNT) = 0 Then
                  Call UpdateOutStockFlagInLIW(TempData.LOT_ITEM_WH_ID, "Y") 'หาก lot นั้นๆหมดแล้ว ก็ให้เปลี่ยน สถานะเลย
               End If
               End If
            Else
             I = I - 1
            End If
        Else
             I = I - 1
         End If
         Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLotIdByPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional TxType As String = "", Optional Cl2 As Collection = Nothing, Optional TypeFlag As Long = 0, Optional LotNew As cLot = Nothing, Optional PartNo As String = "", Optional LocationID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long
Dim CLotI As Collection
Dim CLotE As Collection
Dim TempLotDoc As CLotDoc

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset
   Set CLotI = New Collection
   Set CLotE = New Collection

   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.PART_NO = PartNo
   D.LOT_ID = LotId
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = OrderBy
   D.OrderType = OrderType
   D.TX_TYPE = TxType
   'D.BALANCE_FLAG = "N"
   Call D.QueryData(Ind, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (LotNew Is Nothing) Then
      I = I + 1
      C.AddItem (LotNew.LOT_NO)
      C.ItemData(I) = LotNew.LOT_ID
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
Dim LTD As CLotDoc
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(Ind, Rs)
            If Not (C Is Nothing) Then
               If TypeFlag = 1 Then
                  C.AddItem (TempData.LOT_NO)
                  C.ItemData(I) = TempData.LOT_ID
               Else
                  C.AddItem (TempData.LOT_NO & "-" & TempData.TIME_PACK_BEGIN)
                  C.ItemData(I) = TempData.LOT_DOC_ID
               End If
            End If
         
            If Not (Cl Is Nothing) Then
               If KeyType = 1 Then
                  If TypeFlag = 1 Then
                     Call Cl.add(TempData, Trim(str(TempData.LOT_ID)))
                  Else
                     Call Cl.add(TempData, Trim(str(TempData.LOT_DOC_ID)))
                  End If
               End If
            End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBinByLotDocId(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotDocId As Long = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.LOT_DOC_ID = LotDocId
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(4, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      'Debug.Print I
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(4, Rs)
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BIN_NAME)
         C.ItemData(I) = TempData.BIN_NO
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
               Call Cl.add(TempData, Trim(str(TempData.BIN_NO)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   If Not (C Is Nothing) Then
       If C.ListCount > 0 Then
         C.ListIndex = 1
       End If
   End If
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLockByLotDocId(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotDocId As Long = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.LOT_DOC_ID = LotDocId
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(5, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(5, Rs)
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOCK_NAME)
         C.ItemData(I) = TempData.LOCK_NO
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
               Call Cl.add(TempData, Trim(str(TempData.LOCK_NO)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   If Not (C Is Nothing) Then
       If C.ListCount > 0 Then
         C.ListIndex = 1
       End If
   End If
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLotFromLotDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional TxType As String = "", Optional OUT_STOCK_FLAG As String = "N", Optional AdjusFlag As Boolean = False, Optional LocationID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim LTD As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset

   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.LOT_ID = LotId
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = OrderBy
   D.OrderType = OrderType
   D.TX_TYPE = TxType
   D.BALANCE_FLAG = "N" 'เลือกเฉพาะที่ยังไม่ปรับยอด

   If Not AdjusFlag Then
      D.OUT_STOCK_FLAG = OUT_STOCK_FLAG 'เลือกเฉพาะ LotItemWh ที่ ยังมียอด
   End If
   Call D.QueryData(Ind, Rs, ItemCount)


   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(Ind, Rs)
      
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOT_NO)
         C.ItemData(I) = TempData.LOT_DOC_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
           If TxType = "I" Then
           ''Debug.Print I
            Call Cl.add(TempData, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & Trim(TempData.TX_TYPE)))
          ElseIf TxType = "E" Then
            Set LTD = GetObject("CLotDoc", Cl, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID_REF)) & "-" & Trim(TempData.TX_TYPE)), False)
            If Not (LTD Is Nothing) Then
               LTD.LOT_AMOUNT = LTD.LOT_AMOUNT + TempData.LOT_AMOUNT
            Else
               Call Cl.add(TempData, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID_REF)) & "-" & Trim(TempData.TX_TYPE)))
            End If
          End If
         End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLotRefExByLotDocId(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LotDocId As Long = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim LTD As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset

   D.LOT_DOC_ID = -1
   D.LOT_DOC_ID_REF = LotDocId
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 1
   D.OrderType = 1
   D.TX_TYPE = "E"
   D.BALANCE_FLAG = "N" 'เลือกเฉพาะที่ยังไม่ปรับยอด

'   If Not AdjusFlag Then
'      D.OUT_STOCK_FLAG = OUT_STOCK_FLAG 'เลือกเฉพาะ LotItemWh ที่ ยังมียอด
'   End If
   Call D.QueryData(11, Rs, ItemCount)


'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(11, Rs)
      
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.LOT_NO)
'         C.ItemData(I) = TempData.LOT_DOC_ID
'      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(TempData.LOT_NO) & "-" & Trim(TempData.PALLET_DOC_NO))
          End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLotFromLot(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional DateIn As String = "", Optional LotNo As String = "")
On Error GoTo ErrorHandler
Dim D As cLot
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As cLot
Dim I As Long

   Set D = New cLot
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = PartItemID
   D.LOT_ID = LotId
   D.LOT_NO = LotNo
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DATE_IN = DateIn
   D.OrderBy = OrderBy
   D.OrderType = OrderType
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New cLot
      Call TempData.PopulateFromRS(Ind, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOT_NO)
         C.ItemData(I) = TempData.LOT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.LOT_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.LOT_NO))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub UpdateOutStockFlagInLIW(LotItemWhId As Long, Optional Flag As String = "N")
   Dim LIW As CLotItemWH
   Set LIW = New CLotItemWH
   LIW.LOT_ITEM_WH_ID = LotItemWhId
   LIW.TX_TYPE = "I"
   LIW.UpdateOutStockFlag (Flag)
End Sub
Public Sub UpdateSuccessFlagInBD(BillingdocID As Long, Optional Flag As String = "N")
   Dim BD As CBillingDoc
   Set BD = New CBillingDoc
   BD.BILLING_DOC_ID = BillingdocID
   BD.UpdateSuccessFlag (Flag)
End Sub

Public Sub UpdateDateLot(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1)
On Error GoTo ErrorHandler
Dim D As cLot
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As cLot
Dim I As Long

   Set D = New cLot
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = -1
   D.TO_DATE = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New cLot
      Call TempData.PopulateFromRS(3, Rs)
      TempData.LOT_DATE = TempData.CREATE_DATE
      Call TempData.UpdateLotDate

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLotInPartIemAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1, Optional OrderBy As Long = 1, Optional OrderType As Long = 1, Optional TxType As String = "", Optional Cl2 As Collection = Nothing, Optional PackAmount As Long = 0, Optional DOCUMENT_TYPE As Long = 0, Optional ModeFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long
Dim CLotI As Collection
Dim CLotE As Collection
Dim TempLotDoc As CLotDoc
Dim SumLotAmount As Long
Dim DiffAmount As Double
Dim m_CollPallet As Collection
Dim DOCUMENT_TYPE_INPUT As Long
Dim Flag As Boolean

   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset
   Set CLotI = New Collection
   Set CLotE = New Collection

If ModeFlag = "D" Then
   Call LoadLotFromLotDoc(Nothing, CLotI, , , , 1, PartItemID, 3, , , "I", "Y") 'ถ้าเป็นการลบจะให้หาเฉพาะตัวที่โดนปิด OUT_STOCK_FLAG ไปแล้วเท่านั้น
   Call LoadLotFromLotDoc(Nothing, CLotE, , , , 1, PartItemID, 10, , , "E")
Else
   Call LoadLotFromLotDoc(Nothing, CLotI, , , , 1, PartItemID, 3, , , "I")
   Call LoadLotFromLotDoc(Nothing, CLotE, , , , 1, PartItemID, 10, , , "E")
End If

   D.PART_ITEM_ID = PartItemID
   D.LOT_ID = LotId
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 2 'OrderBy
   D.OrderType = OrderType
   D.TX_TYPE = TxType
   D.BALANCE_FLAG = "N" 'เลือกเฉพาะที่ยังไม่ปรับยอด
   If ModeFlag = "D" Then
      D.OUT_STOCK_FLAG = "Y" 'เลือกเฉพาะ LotItemWh ที่ ไม่มียอด
   Else
      D.OUT_STOCK_FLAG = "N" 'เลือกเฉพาะ LotItemWh ที่ มียอด
   End If
   Call D.QueryData(Ind, Rs, ItemCount)
   
Dim LTD As CLotDoc
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(Ind, Rs)
      Set LTD = GetObject("CLotDoc", CLotI, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "I"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = LTD.LOT_AMOUNT
      End If
      Set LTD = GetObject("CLotDoc", CLotE, Trim(str(TempData.LOT_ID) & "-" & Trim(str(TempData.LOT_DOC_ID)) & "-" & "E"), False)
      If Not (LTD Is Nothing) Then
         TempData.LOT_AMOUNT = TempData.LOT_AMOUNT - LTD.LOT_AMOUNT
      End If

      Flag = True
      For Each LTD In Cl2
         If LTD.LOT_DOC_ID_REF = TempData.LOT_DOC_ID Then
            Flag = False
         End If
      Next LTD
      If Flag = True Then
        If DOCUMENT_TYPE = 2000 Then
          DOCUMENT_TYPE_INPUT = 14
        ElseIf DOCUMENT_TYPE = 2001 Then
         DOCUMENT_TYPE_INPUT = 13
        End If
      End If

        If Val(TempData.LOT_AMOUNT) > 0 Then 'ถ้าไม่มีข้อมูลใน pallet แล้วก็ไม่ต้องให้แสดง Lot
            Call UpdateOutStockFlagInLIW(TempData.LOT_ITEM_WH_ID, "N")
         Else
            Call UpdateOutStockFlagInLIW(TempData.LOT_ITEM_WH_ID, "Y")
         End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLDByLotDocIdRef(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotDocId As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotDoc
Dim I As Long
   Set D = New CLotDoc
   Set Rs = New ADODB.Recordset
   
   D.LOT_DOC_ID_REF = LotDocId
   Call D.QueryData(7, Rs, ItemCount)
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotDoc
      Call TempData.PopulateFromRS(7, Rs)
      
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.LOT_DOC_ID)
'         C.ItemData(I) = TempData.LOT_DOC_ID
'      End If
      
      If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(str(TempData.LOT_DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBinNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional LotId As Long = -1, Optional FromDate As Date, Optional ToDate As Date, Optional KeyType As Long = 1, Optional PartItemID As Long = -1, Optional Ind As Long = 1)
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.LOT_ID = LotId
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New cLot
      Call TempData.PopulateFromRS(Ind, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LOT_NO)
         C.ItemData(I) = TempData.LOT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.LOT_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.LOT_NO))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLocationEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional HouseGroupID As Long)
On Error GoTo ErrorHandler
Dim D As CHouseGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLocation
Dim I As Long
Dim IsOK As Boolean
Dim iCount As Long
Dim Hgi As CHGroupItem

   Set D = New CHouseGroup
   Set Rs = New ADODB.Recordset
   
   D.HOUSE_GROUP_ID = HouseGroupID
   D.EXTRA_FLAG = ""
   D.QueryFlag = 1
   Call glbMaster.QueryPartGroup(D, Rs, iCount, IsOK, glbErrorLog)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   For Each Hgi In D.HGroupItems
      I = I + 1
      If Hgi.SELECT_FLAG = "Y" Then
         Set TempData = New CLocation
         TempData.LOCATION_ID = Hgi.LOCATION_ID
         TempData.LOCATION_NO = Hgi.LOCATION_NO
         TempData.LOCATION_NAME = Hgi.LOCATION_NAME
      
         If Not (C Is Nothing) Then
            C.AddItem (TempData.LOCATION_NAME)
            C.ItemData(I) = TempData.LOCATION_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(str(TempData.LOCATION_ID)))
         End If
         
         Set TempData = Nothing
      End If
   Next Hgi
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadCustomerGrade(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomerGrade
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerGrade
Dim I As Long

   Set D = New CCustomerGrade
   Set Rs = New ADODB.Recordset
   
   D.CSTGRADE_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerGrade
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CSTGRADE_NAME)
         C.ItemData(I) = TempData.CSTGRADE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CSTGRADE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadFreelance(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CFreelance
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFreelance
Dim I As Long

   Set D = New CFreelance
   Set Rs = New ADODB.Recordset

   D.FREELANCE_ID = -1
   Call D.QueryData(Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFreelance
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.FREELANCE_NAME & " " & TempData.FREELANCE_LASTNAME)
         C.ItemData(I) = TempData.FREELANCE_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.FREELANCE_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadCustomerType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomerType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerType
Dim I As Long

   Set D = New CCustomerType
   Set Rs = New ADODB.Recordset
   
   D.CSTTYPE_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CSTTYPE_NAME)
         C.ItemData(I) = TempData.CSTTYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CSTTYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadSupplierGrade(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplierGrade
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierGrade
Dim I As Long

   Set D = New CSupplierGrade
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_GRADE_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierGrade
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_GRADE_NAME)
         C.ItemData(I) = TempData.SUPPLIER_GRADE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.SUPPLIER_GRADE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplierTransport(C As ComboBox, Optional Cl As Collection = Nothing, Optional Truck As String = "", Optional SupplierCode As String = "")
On Error GoTo ErrorHandler
Dim D As CSupplierTranSport
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierTranSport
Dim I As Long

   Set D = New CSupplierTranSport
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_TRANSPORT_ID = -1
   D.SUPPLIER_TRANSPORT_CODE = Truck
   D.SUPPLIER_CODE = SupplierCode
   D.OrderBy = 2
   D.OrderType = 1
   If Truck <> "" Then
       Call D.QueryData(3, Rs, ItemCount)
   ElseIf SupplierCode <> "" Then
       Call D.QueryData(2, Rs, ItemCount)
   Else
       Call D.QueryData(1, Rs, ItemCount)
   End If
  
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierTranSport
      
      If Truck <> "" Then
          Call TempData.PopulateFromRS(3, Rs)
      ElseIf SupplierCode <> "" Then
          Call TempData.PopulateFromRS(2, Rs)
      Else
          Call TempData.PopulateFromRS(1, Rs)
      End If
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_TRANSPORT_DETAIL)
         C.ItemData(I) = TempData.SUPPLIER_TRANSPORT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If Truck <> "" Then
            Call Cl.add(TempData, Trim(TempData.SUPPLIER_TRANSPORT_CODE))
         ElseIf SupplierCode <> "" Then
            Call Cl.add(TempData, Trim(TempData.SUPPLIER_CODE))
         Else
            Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_TRANSPORT_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplierAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional Truck As String = "", Optional SupplierCode As String = "")
On Error GoTo ErrorHandler
Dim D As CSupplierAccount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierAccount
Dim I As Long

   Set D = New CSupplierAccount
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_ACCOUNT_ID = -1
   D.SUPPLIER_CODE = SupplierCode
   D.OrderBy = 2
   D.OrderType = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierAccount
      
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_ACCOUNT_NO)
         C.ItemData(I) = TempData.SUPPLIER_ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SUPPLIER_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
'==
Public Sub LoadSupplierType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplierType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierType
Dim I As Long

   Set D = New CSupplierType
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_TYPE_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_TYPE_NAME)
         C.ItemData(I) = TempData.SUPPLIER_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.SUPPLIER_TYPE_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadSupplierStatus(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSupplierStatus
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierStatus
Dim I As Long

   Set D = New CSupplierStatus
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_STATUS_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_STATUS_NAME)
         C.ItemData(I) = TempData.SUPPLIER_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.SUPPLIER_STATUS_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadPosition(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CEmpPosition
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmpPosition
Dim I As Long

   Set D = New CEmpPosition
   Set Rs = New ADODB.Recordset
   
   D.POSITION_ID = -1
   D.POSITION_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmpPosition
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.POSITION_DESC)
         C.ItemData(I) = TempData.POSITION_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.POSITION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadYearSeq(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CYearSeq
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearSeq
Dim I As Long

   Set D = New CYearSeq
   Set Rs = New ADODB.Recordset
   
   D.YEAR_SEQ_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearSeq
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.YEAR_NO)
         C.ItemData(I) = TempData.YEAR_SEQ_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigStatus(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CProductStatus
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CProductStatus
Dim I As Long

   Set D = New CProductStatus
   Set Rs = New ADODB.Recordset
   
   D.PRODUCT_STATUS_ID = -1
   D.PRODUCT_STATUS_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CProductStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_STATUS_NAME)
         C.ItemData(I) = TempData.PRODUCT_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStockPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartGroupID As Long, Optional ID As Long)
On Error GoTo ErrorHandler
Dim D As CPartItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long

   Set D = New CPartItem
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.CANCEL_FLAG = "N"
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS2(6, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_DESC)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadPartType(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartGroupID As Long, Optional ID As Long)
On Error GoTo ErrorHandler
Dim D As CPartType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartType
Dim I As Long

   Set D = New CPartType
   Set Rs = New ADODB.Recordset
   
   D.PART_TYPE_ID = -1
   D.PART_TYPE_NAME = ""
   D.PART_GROUP_ID = PartGroupID
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_TYPE_NAME)
         C.ItemData(I) = TempData.PART_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartGroup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPartGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartGroup
Dim I As Long

   Set D = New CPartGroup
   Set Rs = New ADODB.Recordset
   
   D.PART_GROUP_ID = -1
   D.PART_GROUP_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_GROUP_NAME)
         C.ItemData(I) = TempData.PART_GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartMaster(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPartMaster
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartMaster
Dim I As Long

   Set D = New CPartMaster
   Set Rs = New ADODB.Recordset
   
   D.PART_MASTER_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartMaster
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_MASTER_NAME)
         C.ItemData(I) = TempData.PART_MASTER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_MASTER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
'Public Sub LoadExWorksPriceItem(C As ComboBox, Optional Cl As Collection = Nothing)
'On Error GoTo ErrorHandler
'Dim D As CExWorksPriceItem
'Dim ItemCount As Long
'Dim Rs As ADODB.Recordset
'Dim TempData As CExWorksPriceItem
'Dim I As Long
'
'   Set D = New CExWorksPriceItem
'   Set Rs = New ADODB.Recordset
'
'   D.PART_GROUP_ID = -1
'   D.PART_GROUP_NAME = ""
'   Call D.QueryData(Rs, ItemCount)
'
'   If Not (C Is Nothing) Then
'      C.Clear
'      I = 0
'      C.AddItem ("")
'   End If
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CPartGroup
'      Call TempData.PopulateFromRS(1, Rs)
'
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.PART_GROUP_NAME)
'         C.ItemData(I) = TempData.PART_GROUP_ID
'      End If
'
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(str(TempData.PART_GROUP_ID)))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'
'   Set Rs = Nothing
'   Set D = Nothing
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub
Public Sub LoadHouseGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExtraFlag = "")
On Error GoTo ErrorHandler
Dim D As CHouseGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CHouseGroup
Dim I As Long

   Set D = New CHouseGroup
   Set Rs = New ADODB.Recordset
   
   D.HOUSE_GROUP_ID = -1
   D.EXTRA_FLAG = ExtraFlag
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CHouseGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.HOUSE_GROUP_NAME)
         C.ItemData(I) = TempData.HOUSE_GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.HOUSE_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadStatusGroup(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExtraFlag = "")
On Error GoTo ErrorHandler
Dim D As CStatusGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStatusGroup
Dim I As Long

   Set D = New CStatusGroup
   Set Rs = New ADODB.Recordset
   
   D.STATUS_GROUP_ID = -1
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStatusGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STATUS_GROUP_NAME)
         C.ItemData(I) = TempData.STATUS_GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.STATUS_GROUP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadYearWeek(C As ComboBox, Optional Cl As Collection = Nothing, Optional YearSeqID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CYearWeek
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearWeek
Dim I As Long

   Set D = New CYearWeek
   Set Rs = New ADODB.Recordset
   
   D.YEAR_WEEK_ID = -1
   D.YEAR_SEQ_ID = YearSeqID
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearWeek
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.WEEK_NO)
         C.ItemData(I) = TempData.YEAR_WEEK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.YEAR_WEEK_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctYearWeek(C As ComboBox, Optional Cl As Collection = Nothing, Optional YearSeqID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CYearWeek
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CYearWeek
Dim I As Long

   Set D = New CYearWeek
   Set Rs = New ADODB.Recordset
   
   D.YEAR_WEEK_ID = -1
   D.YEAR_SEQ_ID = YearSeqID
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CYearWeek
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.WEEK_NO)
         C.ItemData(I) = TempData.YEAR_WEEK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.WEEK_NO)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Function CompareInt(Key1 As Long, Key2 As Long) As Boolean
   If Key2 <= 0 Then
      CompareInt = True
   Else
      CompareInt = (Key1 = Key2)
   End If
End Function

Public Function CompareStr(Key1 As String, Key2 As String) As Boolean
   If Len(Key2) <= 0 Then
      CompareStr = True
   Else
      CompareStr = (Key1 = Key2)
   End If
End Function
Public Sub LoadPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartTypeID As Long = -1, Optional PigFlag As String = "", Optional PigType As String = "", Optional KeyType As Byte = 1, Optional CancelFlag As String = "N")
'On Error GoTo ErrorHandler
On Error Resume Next
Dim D As CPartItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long
   
   Set D = New CPartItem
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.PART_TYPE = PartTypeID
   D.PIG_FLAG = PigFlag
   D.CANCEL_FLAG = CancelFlag
   D.PIG_TYPE = PigType
   D.OrderBy = 5
   D.OrderType = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
   
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC & "  (" & TempData.PART_NO & ")")
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.PART_NO))
         ElseIf KeyType = 3 Then
            If Len(TempData.NUMBER_PLC_ID) > 0 Then
               Call Cl.add(TempData, Trim(TempData.NUMBER_PLC_ID))
            End If
         ElseIf KeyType = 4 Then
            If Len(TempData.NUMBER_LAB_ID) > 0 Then
               Call Cl.add(TempData, Trim(TempData.NUMBER_LAB_ID))
            End If
         End If
      End If
            
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDeliveryCus(C As ComboBox, Optional Cl As Collection = Nothing, Optional CusID As Long = -1, Optional PigFlag As String = "", Optional PigType As String = "", Optional KeyType As Byte = 1, Optional CancelFlag As String = "N")
On Error GoTo ErrorHandler
Dim D As CDeliveryCus
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDeliveryCus
Dim I As Long

   Set D = New CDeliveryCus
   Set Rs = New ADODB.Recordset

   D.DELIVERY_CUS_ITEM_ID = -1
   D.CUSTOMER_ID = CusID
   Call D.QueryData(Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDeliveryCus
      Call TempData.PopulateFromRS(1, Rs)
     TempData.Flag = "I"
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DELIVERY_CUS_ITEM_NAME)
         C.ItemData(I) = TempData.DELIVERY_CUS_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
      If KeyType = 2 Then
         Call Cl.add(TempData, Trim(TempData.DELIVERY_CUS_ITEM_CODE) & "-" & Trim(str(TempData.CUSTOMER_ID)))
      Else
         Call Cl.add(TempData, Trim(str(TempData.DELIVERY_CUS_ITEM_ID)))
      End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPromotional(C As ComboBox, Optional Cl As Collection = Nothing, Optional CusID As Long = -1, Optional KeyType As Byte = 1)
On Error GoTo ErrorHandler
Dim D As CPromotional
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPromotional
Dim I As Long

   Set D = New CPromotional
   Set Rs = New ADODB.Recordset

   D.PROMOTIONAL_DETAIL_ID = -1
   D.CUSTOMER_ID = CusID
   Call D.QueryData(Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPromotional
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.PROMOTIONAL_DETAIL_NAME)
         C.ItemData(I) = TempData.PROMOTIONAL_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
      If KeyType = 2 Then
         Call Cl.add(TempData, Trim(TempData.PROMOTIONAL_DETAIL_ID) & "-" & Trim(str(TempData.CUSTOMER_ID)))
      Else
         Call Cl.add(TempData, Trim(str(TempData.PROMOTIONAL_ITEM_ID)))
      End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadExWorksPriceItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExWorksPriceId As Long = -1, Optional KeyType As Byte = 1, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CExWorksPrice
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExWorksPrice
Dim I As Long
Dim FUNC_NAME As String
Dim Key As String
FUNC_NAME = "LoadExWorksPriceItem"

   Set D = New CExWorksPrice
   Set Rs = New ADODB.Recordset

   D.EX_WORKS_PRICE_ID = ExWorksPriceId
   D.EX_WORKS_PRICE_TYPE = 1
   D.BETWEEN_DATE = FromDate
   Call D.QueryData(2, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExWorksPrice
      Call TempData.PopulateFromRS(2, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_NO)
         C.ItemData(I) = TempData.EX_WORKS_PRICE_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
              Key = Trim(str(TempData.EX_WORKS_PRICE_ITEM_ID))
             Call Cl.add(TempData, Key)
         ElseIf KeyType = 2 Then
            Key = Trim(str(TempData.PART_ITEM_ID))
            Call Cl.add(TempData, Key)
         End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
  glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & " , " & FUNC_NAME & " , KEY= " & Key & " , KeyType=" & KeyType
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadExDeliveryCusItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExWorksPriceId As Long = -1, Optional KeyType As Byte = 1, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CExWorksPrice
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExWorksPrice
Dim I As Long
Dim FUNC_NAME As String
Dim Key As String
FUNC_NAME = "LoadExDeliveryCusItem"

   Set D = New CExWorksPrice
   Set Rs = New ADODB.Recordset

   D.EX_WORKS_PRICE_ID = ExWorksPriceId
   D.EX_WORKS_PRICE_TYPE = 2
   D.BETWEEN_DATE = FromDate
   Call D.QueryData(3, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExWorksPrice
      Call TempData.PopulateFromRS(3, Rs)

      If Not (C Is Nothing) Then
         If KeyType = 1 Then
            C.AddItem (TempData.DELIVERY_CUS_ITEM_NAME)
         ElseIf KeyType = 2 Then
            C.AddItem (TempData.DELIVERY_CUS_ITEM_NAME & "(" & TempData.WEIGHT_PER_PACK_CUS & ")")
         ElseIf KeyType = 3 Then
            C.AddItem (TempData.DELIVERY_CUS_ITEM_NAME & "(" & TempData.WEIGHT_PER_PACK & ")")
         End If
         C.ItemData(I) = TempData.DELIVERY_CUS_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
             Key = Trim(str(TempData.EX_DELIVERY_COST_ITEM_ID))
             Call Cl.add(TempData, Key)
         ElseIf KeyType = 2 Then
            Key = Trim(str(TempData.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(TempData.WEIGHT_PER_PACK_CUS))
            Call Cl.add(TempData, Key) 'เอาค่าขนส่งที่คิดลูกค้า
         ElseIf KeyType = 3 Then
           Key = Trim(str(TempData.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(TempData.WEIGHT_PER_PACK))
           Call Cl.add(TempData, Key) 'เอาค่าขนส่งที่คิดให้รถรับจ้าง
         ElseIf KeyType = 4 Then
           Key = Trim(str(TempData.CUSTOMER_ID)) & "-" & Trim(str(TempData.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(TempData.RATE_TYPE)) & "-" & Trim(str(TempData.RATE_TYPE_CUS))
           Call Cl.add(TempData, Key)
         End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & " , " & FUNC_NAME & " , KEY= " & Key & " , KeyType=" & KeyType
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadExPromotionPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExWorksPriceId As Long = -1, Optional KeyType As Byte = 1, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CExWorksPrice
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExWorksPrice
Dim I As Long

   Set D = New CExWorksPrice
   Set Rs = New ADODB.Recordset

   D.EX_WORKS_PRICE_ID = ExWorksPriceId
   D.EX_WORKS_PRICE_TYPE = 3
   D.BETWEEN_DATE = FromDate
   Call D.QueryData(4, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExWorksPrice
      Call TempData.PopulateFromRS(4, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_NAME & "-" & TempData.PART_NO)
         C.ItemData(I) = TempData.EX_PROMOTION_PART_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
             Call Cl.add(TempData, Trim(str(TempData.EX_PROMOTION_PART_ITEM_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)) & "-" & Trim(str(TempData.PART_ITEM_ID)))
         End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadExPromotionDlcItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional ExWorksPriceId As Long = -1, Optional KeyType As Byte = 1, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CExWorksPrice
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExWorksPrice
Dim I As Long

   Set D = New CExWorksPrice
   Set Rs = New ADODB.Recordset

   D.EX_WORKS_PRICE_ID = ExWorksPriceId
   D.EX_WORKS_PRICE_TYPE = 4
   D.BETWEEN_DATE = FromDate
   Call D.QueryData(5, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExWorksPrice
      Call TempData.PopulateFromRS(5, Rs)

     If Not (C Is Nothing) Then
         If KeyType = 1 Then
            C.AddItem (TempData.DELIVERY_CUS_ITEM_NAME)
         ElseIf KeyType = 2 Then
            C.AddItem (TempData.DELIVERY_CUS_ITEM_NAME & "(" & TempData.WEIGHT_PER_PACK_CUS & ")")
         End If
         C.ItemData(I) = TempData.DELIVERY_CUS_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
             Call Cl.add(TempData, Trim(str(TempData.EX_PROMOTION_DLC_ITEM_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(str(TempData.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(TempData.WEIGHT_PER_PACK_CUS))) 'เอาค่าขนส่งที่คิดลูกค้า
          ElseIf KeyType = 3 Then
            Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)) & "-" & Trim(str(TempData.DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(TempData.RATE_TYPE_CUS)))
         End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartStockFromLotItemWh(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional PartTypeID As Long = -1, Optional PartID As Long = -1, Optional TxType As String = "", Optional KeyType As Byte = 1, Optional Ind As Long = 1, Optional DOCUMENT_TYPE As Long = 1, Optional DateType As Long = 1, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
'On Error Resume Next
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim m_CollLotItemWh As Collection
Dim LTD As CLotDoc
Dim Key As String

Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset

   D.PART_ITEM_ID = PartID
   D.BALANCE_FLAG = "N"
   D.PART_TYPE = PartTypeID
   D.LOCATION_ID = LocationID
   If TxType = "I" Then
      If DOCUMENT_TYPE = 13 Then 'bulk
         D.DOCUMENT_TYPE_SET = "(13,16,18,21)"
      ElseIf DOCUMENT_TYPE = 14 Then 'bag
         D.DOCUMENT_TYPE_SET = "(14,15,17,20)"
      End If
   End If
   D.TX_TYPE = TxType
   If DateType = 2 Then
      D.FROM_PACK_DATE = FromDate
      D.TO_PACK_DATE = ToDate
   Else
      D.FROM_DATE = FromDate
      D.TO_DATE = ToDate
   End If
   D.OrderBy = 1
   D.OrderType = 1
   
   
   
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(Ind, Rs)
   
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC & "  (" & TempData.PART_NO & ")")
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
'            If Not (Cl Is Nothing) Then
'         If KeyType = 1 Then
'            If TxType = "I" Then
'             If DOCUMENT_TYPE = 13 Then 'BULK
'                Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.LOT_ID) & "-" & TxType & "-" & str(TempData.LOT_DOC_ID))
'             Else 'Bag
'                Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.WEIGHT_PER_PACK) & "-" & str(TempData.LOT_DOC_ID) & "-" & str(TempData.LOT_ID) & "-" & str(TempData.BIN_NO) & "-" & str(TempData.LOCK_NO) & "-" & TxType)
'            End If
'            ElseIf TxType = "E" Then
'               If DOCUMENT_TYPE = 13 Then 'BULK
'                  Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.LOT_ID) & "-" & TxType & "-" & str(TempData.LOT_DOC_ID_REF))
'               Else 'Bag
'                  Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.WEIGHT_PER_PACK) & "-" & str(TempData.LOT_DOC_ID_REF) & "-" & str(TempData.LOT_ID) & "-" & str(TempData.BIN_NO) & "-" & str(TempData.LOCK_NO) & "-" & TxType)
'               End If
'            End If
'            Call Cl.add(TempData, Key)
'         ElseIf KeyType = 2 Then
'            Call Cl.add(TempData, Trim(TempData.PART_NO))
'         ElseIf KeyType = 3 Then
'            Call Cl.add(TempData, Trim(TempData.PART_NO) & "-" & Trim(str(TempData.WEIGHT_PER_PACK)) & "-" & TxType)
'         End If
'      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            If TxType = "I" Then
             If DOCUMENT_TYPE = 13 Then 'BULK
                Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.LOT_ID) & "-" & TxType & "-" & str(TempData.LOT_DOC_ID) & "-" & str(TempData.LOCATION_ID))
             Else 'Bag
                Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.WEIGHT_PER_PACK) & "-" & str(TempData.LOT_DOC_ID) & "-" & str(TempData.LOT_ID) & "-" & str(TempData.BIN_NO) & "-" & str(TempData.LOCK_NO) & "-" & TxType & "-" & str(TempData.LOCATION_ID))
            End If
            ElseIf TxType = "E" Then
               If DOCUMENT_TYPE = 13 Then 'BULK
                  Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.LOT_ID) & "-" & TxType & "-" & str(TempData.LOT_DOC_ID_REF) & "-" & str(TempData.LOCATION_ID))
               Else 'Bag
                  Key = Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.WEIGHT_PER_PACK) & "-" & str(TempData.LOT_DOC_ID_REF) & "-" & str(TempData.LOT_ID) & "-" & str(TempData.BIN_NO) & "-" & str(TempData.LOCK_NO) & "-" & TxType & "-" & str(TempData.LOCATION_ID))
               End If
            End If
            Call Cl.add(TempData, Key)
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.PART_NO))
         ElseIf KeyType = 3 Then
            Call Cl.add(TempData, Trim(TempData.PART_NO) & "-" & Trim(str(TempData.WEIGHT_PER_PACK)) & "-" & TxType)
         End If
      End If
            
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartItemWh(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartTypeID As Long = -1, Optional TxType As String = "", Optional PigType As String = "", Optional KeyType As Byte = 1, Optional CancelFlag As String = "")
On Error GoTo ErrorHandler
'On Error Resume Next
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.TX_TYPE = TxType
   D.OrderBy = 5
   D.OrderType = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC & "  (" & TempData.PART_NO & ")")
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.PART_NO & "-" & TempData.WEIGHT_PER_PACK))
         End If
      End If
            
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function getSoBalanceAmount(BILLING_DOC_ID As Long) As Double
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long
Dim Sum As Double

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.BILLING_DOC_ID = BILLING_DOC_ID
   D.OrderBy = 5
   D.OrderType = 1
   Call D.QueryData(15, Rs, ItemCount)
   
   While Not Rs.EOF
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(15, Rs)
       Sum = Sum + TempData.PACK_AMOUNT
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
    getSoBalanceAmount = Sum
   Set Rs = Nothing
   Set D = Nothing
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function
Public Sub LoadPartItemCodeKey(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartTypeID As Long = -1, Optional PigFlag As String = "N", Optional PigType As String = "")
'On Error GoTo ErrorHandler
On Error Resume Next
Dim D As CPartItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long

   Set D = New CPartItem
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.PART_TYPE = PartTypeID
   D.PIG_FLAG = PigFlag
   D.PIG_TYPE = PigType
   D.OrderBy = 3
   D.OrderType = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC & "  (" & TempData.PART_NO & ")")
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_NO)
      End If
            
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartItemGroupType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPartItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long

   Set D = New CPartItem
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.OrderType = 1
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS(4, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC & "  (" & TempData.PART_NO & ")")
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
            
      Set TempData = Nothing
      
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPigItem(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPartItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItem
Dim I As Long

   Set D = New CPartItem
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.PART_TYPE = -1
   D.PIG_FLAG = "Y"
   D.PIG_TYPE = ""
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CPartItem
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_NO & "-" & TempData.PIG_TYPE)
'''Debug.Print TempData.PART_NO & "-" & TempData.PIG_TYPE
      End If
            
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadUnit(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CUnit
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUnit
Dim I As Long

   Set D = New CUnit
   Set Rs = New ADODB.Recordset
   
   D.UNIT_ID = -1
   D.UNIT_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUnit
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.UNIT_NAME)
         C.ItemData(I) = TempData.UNIT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData)
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.UNIT_ID))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadSumWithDoItemId(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = 1, Optional PartItemID As Long = 1)
On Error GoTo ErrorHandler
Dim D As CUnit
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CUnit
   Set Rs = New ADODB.Recordset
      
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   For Each TempData In Cl
      I = I + 1
      If Not (C Is Nothing) Then
         If TempData.PART_ITEM_ID > 0 And TempData.PART_ITEM_ID <> PartItemID Then 'ไม่เอาค่าขนส่ง
            C.AddItem (TempData.PART_NO)
            C.ItemData(I) = TempData.DO_ITEM_ID
         Else
         I = I - 1
         End If
      End If
   Next TempData

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSellType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSellType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSellType
Dim I As Long

   Set D = New CSellType
   Set Rs = New ADODB.Recordset
   
   D.SELL_TYPE_ID = -1
   D.SELL_TYPE_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSellType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SELL_TYPE_NAME)
         C.ItemData(I) = TempData.SELL_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SELL_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSubLotItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional Li As CLotItem = Nothing)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   Set Rs = New ADODB.Recordset
   Call Li.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.LOT_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalAmountPartItemLotItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromTxType As String, Optional DocumentType As Long) ' Optional LocationID As Long
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim Li As CLotItem

   Set Li = New CLotItem
   Li.LOT_ITEM_ID = -1
   Li.FROM_TX_TYPE = FromTxType
   Li.DOCUMENT_TYPE = DocumentType
'   Li.LOCATION_ID = LocationID
   Set Rs = New ADODB.Recordset
   Call Li.QueryData(33, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(33, Rs)
   
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
        Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID))

'         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.LOCATION_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCostPerUnitByPartNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As String, Optional ToDate As String)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim Li As CLotItem

   Set Li = New CLotItem
   Li.LOT_ITEM_ID = -1
   Li.TX_TYPE = "I"
   Li.FROM_DATE = FromDate
   Li.TO_DATE = ToDate
   Li.DOCUMENT_TYPE_SET = "(12, 13, 14)"
   Set Rs = New ADODB.Recordset
   Call Li.QueryData(17, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(17, Rs)
   
'      If Not (C Is Nothing) Then
'      End If
      
      If Not (Cl Is Nothing) Then
        Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadImportPrice2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   '==
   FromDate1 = DateToStringIntLow(-2)
   ToDate1 = DateToStringIntLow(FromDate)
   '==
   
   D.FROM_DATE = InternalDateToDate(FromDate1)
   D.TO_DATE = InternalDateToDate(ToDate1)
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLeftAmount1(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLeftAmount2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional HiFlag As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = -1
   D.TO_DATE1 = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   D.HiFlag = HiFlag
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice4(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.LOCATION_ID)) & "-" & Trim(str(TempData.DOCUMENT_TYPE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice5(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(13, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)) & "-" & Trim(str(TempData.DOCUMENT_TYPE)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadImportPrice6(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(14, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)) & "-" & Trim(str(TempData.LOCATION_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartTypeID As Long = -1, Optional PartGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
Dim Pt As CPartType

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartType(Nothing, TempCol)
   End If
   
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartTypeID
   D.PART_GROUP_ID = PartGroupID
   D.OrderBy = 1
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pt = TempCol(Trim(str(TempData.PART_TYPE)))
      If Not (Cl Is Nothing) Then
         Call Cl.add(Pt, Trim(str(Pt.PART_TYPE_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartInJob(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional PartTypeID As Long = -1, Optional ProcessID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CJob
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.PART_TYPE = PartTypeID
   D.PROCESS_ID = ProcessID
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
Dim Pt As CPartItem

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol)
   End If
   
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
'D.PART_ITEM_ID = 256
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      Set Pt = TempCol(Trim(str(TempData.PART_ITEM_ID)))
      If Not (Cl Is Nothing) Then
         Call Cl.add(Pt, Trim(str(Pt.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartItemLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartType As Long = -1, Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
Dim Pt As CPartItem

   If TempCol Is Nothing Then
   End If
   
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
'D.PART_ITEM_ID = 123
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PART_GROUP_ID = PartGroup
   D.OrderBy = 1
   Call D.QueryData(14, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartItemLocationBA(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartType As Long = -1, Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = -1 'FromDate
   D.TO_DATE = -1 'ToDate
'   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PART_GROUP_ID = PartGroup
   D.OrderBy = 1
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
'If TempData.PART_ITEM_ID = 314 Then
''Debug.Print
'End If
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctPartItemTx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartType As Long = -1, Optional PartNo As String = "")
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim Pt As CPartItem
   
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PART_NO = PartNo
   D.OrderBy = 1
   Call D.QueryData(22, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(22, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub MergeCollection(Col1 As Collection, Col2 As Collection, SortType As Long)
Dim Ba As CBalanceAccum
Dim TempLi As CLotItem
Dim TempLi2 As CLotItem
Dim Li As CLotItem

   For Each Ba In Col1
'If Ba.PART_NO = "1013" Then
''Debug.Print
'End If
'If Ba.PART_ITEM_ID = 314 Then
''Debug.Print
'End If

      Set TempLi = GetLotItemEx(Col2, Ba.LOCATION_ID & "-" & Ba.PART_ITEM_ID)
      If TempLi Is Nothing Then
         Set TempLi2 = New CLotItem
         TempLi2.PART_ITEM_ID = Ba.PART_ITEM_ID
         TempLi2.LOCATION_ID = Ba.LOCATION_ID
         TempLi2.LOCATION_NO = Ba.LOCATION_NO
         TempLi2.LOCATION_NAME = Ba.LOCATION_NAME
         TempLi2.PART_NO = Ba.PART_NO
         TempLi2.PART_DESC = Ba.PART_DESC
         TempLi2.PART_TYPE_NO = Ba.PART_TYPE_NO
         TempLi2.PART_TYPE_NAME = Ba.PART_TYPE_NAME
         Call Col2.add(TempLi2, Ba.LOCATION_ID & "-" & Ba.PART_ITEM_ID)
'''Debug.Print TempLi2.PART_ITEM_ID & "-" & TempLi2.LOCATION_ID
         Set TempLi2 = Nothing
      End If
'''Debug.Print Ba.PART_ITEM_ID & "-" & Ba.LOCATION_ID
   Next Ba
   
   'Call SelectionsortEx(Col2, 1, Col2.Count, SortType)           'ยกเลิกการใช้งานเพราะรันนานมาก
End Sub

Public Sub LoadPigImportAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(15, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadJobProductRMAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional ProcessID As Long, Optional TxType As String = "", Optional PartItemSet As String = "")
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   D.TX_TYPE = TxType
   D.PART_ITEM_SET = PartItemSet
   D.OrderBy = 1
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.JOB_PART_ITEM_ID & "-" & TempData.TX_TYPE)
'''Debug.Print TempData.JOB_PART_ITEM_ID & "-" & TempData.TX_TYPE & " " & TempData.TOTAL_INCLUDE_PRICE
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadJobProductRMPrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional ProcessID As Long, Optional TxType As String = "", Optional PartGroupID As Long = -1, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   D.TX_TYPE = TxType
   D.PART_GROUP_ID = PartGroupID
   D.PART_TYPE_ID = PartType
   D.OrderBy = 1
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.JOB_PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadJobProductExpenseAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional ProcessID As Long)
On Error GoTo ErrorHandler
Dim D As CJobParameter
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobParameter
Dim I As Long

   Set D = New CJobParameter
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   D.OrderBy = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobParameter
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PARAMETER_PROCESS_ID)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplier(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Byte = 1, Optional SupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CSupplier
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CSupplier
Dim I As Long

   Set D = New CSupplier
   D.SUPPLIER_ID = -1
   D.OrderBy = 1
   D.OrderType = 1
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData2(Rs, ItemCount)
   Else
      If Not Rs.EOF Then
         Rs.MoveFirst
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplier
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_NAME)
         C.ItemData(I) = TempData.SUPPLIER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.SUPPLIER_CODE))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerID As Long = -1, Optional CustManualFlag As String = "", Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CAccount
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CAccount
Dim I As Long

   Set D = New CAccount
   D.ACCOUNT_ID = -1
   D.CUSTOMER_ID = CustomerID
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.OrderType = 1
   D.MANUAL_FLAG = CustManualFlag
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(1, Rs, ItemCount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.ACCOUNT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccountByCustCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerID As Long = -1, Optional CustManualFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CAccount
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CAccount
Dim I As Long

   Set D = New CAccount
   D.ACCOUNT_ID = -1
   D.CUSTOMER_ID = CustomerID
   D.OrderType = 1
   D.MANUAL_FLAG = CustManualFlag
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(1, Rs, ItemCount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
'''Debug.Print TempData.ACCOUNT_NO & "-" & TempData.CUSTOMER_CODE
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerID As Long = -1, Optional ShowFirst As Boolean = True, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long

   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   D.CUSTOMER_ID = CustomerID
   Call D.QueryData3(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
        If KeyType = 1 Then
          Call Cl.add(TempData)
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(str(TempData.ADDRESS_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplierAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long

   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   D.SUPPLIER_ID = SupplierID
   Call D.QueryData4(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupplierAddressGroupID(Cl As Collection, Optional SupplierID As Long = -1, Optional AmPhurSearch As String, Optional ProvinceSearch As String)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long

   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   D.SUPPLIER_ID = SupplierID
   D.AMPHUR = PatchWildCard(AmPhurSearch)
   D.PROVINCE = PatchWildCard(ProvinceSearch)
   
   Call D.QueryData5(Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS5(Rs)
      
      Set D = GetObject("CAddress", Cl, Trim(str(TempData.SUPPLIER_ID)), False)
      If D Is Nothing Then
         Set D = New CAddress
         D.SUPPLIER_ID = TempData.SUPPLIER_ID
         Call Cl.add(D, Trim(str(TempData.SUPPLIER_ID)))
      End If
      
      Call D.collSupAddr.add(TempData)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadEnterpriseAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional EnterpriseID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
Dim TempIndex As Long

   TempIndex = 0
   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = EnterpriseID
   Call D.QueryData2(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CCustomer
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CCustomer
Dim I As Long

   Set D = New CCustomer
   D.CUSTOMER_ID = -1
   D.OrderBy = 2
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData2(Rs, ItemCount)
   Else
      If Not Rs.EOF Then
         Rs.MoveFirst
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomer
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_NAME)
         C.ItemData(I) = TempData.CUSTOMER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.CUSTOMER_CODE))
         Else
            Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomerExt(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CCustomerExt
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CCustomerExt
Dim I As Long

   Set D = New CCustomerExt
   D.CUSTOMER_ID = -1
   D.OrderBy = 2
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData2(Rs, ItemCount)
   Else
      If Not Rs.EOF Then
         Rs.MoveFirst
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerExt
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_NAME)
         C.ItemData(I) = TempData.CUSTOMER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.CUSTOMER_CODE))
         Else
            Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadEmployee(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CEmployee
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmployee
Dim I As Long
   Set D = New CEmployee
   Set Rs = New ADODB.Recordset
   
   D.EMP_ID = -1
   D.CURRENT_POSITION = ID
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmployee
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.NAME & " " & TempData.LASTNAME)
         C.ItemData(I) = TempData.EMP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.EMP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


'==
Public Sub LoadProductStatus(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CProductStatus
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CProductStatus
Dim I As Long

   Set D = New CProductStatus
   Set Rs = New ADODB.Recordset
   
   D.PRODUCT_STATUS_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CProductStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PRODUCT_STATUS_NAME)
         C.ItemData(I) = TempData.PRODUCT_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PRODUCT_STATUS_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadProductStatusEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional StatusGroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CStatusGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSGroupItem
Dim Sgi As CSGroupItem
Dim I As Long
Dim IsOK As Boolean

   Set D = New CStatusGroup
   Set Rs = New ADODB.Recordset
   
   D.STATUS_GROUP_ID = StatusGroupID
   D.QueryFlag = 1
   Call glbMaster.QueryStatusGroup(D, Rs, ItemCount, IsOK, glbErrorLog)
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   For Each Sgi In D.HGroupItems
      I = I + 1
      Set TempData = New CSGroupItem
      Call TempData.CopyField(1, Sgi)
      
      If Not (C Is Nothing) Then
         C.AddItem (Sgi.STATUS_NAME)
         C.ItemData(I) = Sgi.ST_STATUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.ST_STATUS_ID)))
      End If
      
      Set TempData = Nothing
   Next Sgi
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function PigCodeToID(Cd As String) As Long
   If Cd = "N" Then
      PigCodeToID = 1
   ElseIf Cd = "B" Then
      PigCodeToID = 2
   ElseIf Cd = "BT" Then
      PigCodeToID = 3
   ElseIf Cd = "G" Then
      PigCodeToID = 4
   ElseIf Cd = "L" Then
      PigCodeToID = 5
   ElseIf Cd = "R" Then
      PigCodeToID = 6
   End If
End Function

Public Function PigTypeToCode(Cd As Long) As String
   If Cd = 1 Then
      PigTypeToCode = "N"
   ElseIf Cd = 2 Then
      PigTypeToCode = "B"
   ElseIf Cd = 3 Then
      PigTypeToCode = "BT"
   ElseIf Cd = 4 Then
      PigTypeToCode = "G"
   ElseIf Cd = 5 Then
      PigTypeToCode = "L"
   ElseIf Cd = 6 Then
      PigTypeToCode = "R"
   Else
      PigTypeToCode = ""
   End If
End Function

Public Sub PopulatePigType(TempID As Long, PigType As CPigType)
   If TempID = 1 Then
      PigType.PIG_TYPE_ID = 1
      PigType.PIG_TYPE_NO = "N"
      PigType.PIG_TYPE_NAME = "หมูปกติ (N)"
   ElseIf TempID = 2 Then
      PigType.PIG_TYPE_ID = 2
      PigType.PIG_TYPE_NO = "B"
      PigType.PIG_TYPE_NAME = "หมูพ่อพันธ์ (B)"
   ElseIf TempID = 3 Then
      PigType.PIG_TYPE_ID = 3
      PigType.PIG_TYPE_NO = "BT"
      PigType.PIG_TYPE_NAME = "หมูสำรองพ่อ (BT)"
   ElseIf TempID = 4 Then
      PigType.PIG_TYPE_ID = 4
      PigType.PIG_TYPE_NO = "G"
      PigType.PIG_TYPE_NAME = "แม่อุ้มท้อง (G)"
   ElseIf TempID = 5 Then
      PigType.PIG_TYPE_ID = 5
      PigType.PIG_TYPE_NO = "L"
      PigType.PIG_TYPE_NAME = "แม่คลอด (L)"
   ElseIf TempID = 6 Then
      PigType.PIG_TYPE_ID = 6
      PigType.PIG_TYPE_NO = "R"
      PigType.PIG_TYPE_NAME = "สำรองแม่ (R)"
   End If
   
   PigType.KEY_ID = PigType.PIG_TYPE_ID
   PigType.KEY_LOOKUP = PigType.PIG_TYPE_NO
End Sub

'==
Public Sub LoadPigType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim I As Long
Dim TempData As CPigType

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   For I = 1 To 6
      Set TempData = New CPigType
      Call PopulatePigType(I, TempData)
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PIG_TYPE_NAME)
         C.ItemData(I) = TempData.PIG_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PIG_TYPE_ID)))
      End If
      
      Set TempData = Nothing
   Next I
         
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Function DocumentTypeToString(Dt As Long) As String
   If Dt = 1 Then
      DocumentTypeToString = "ใบนำเข้า"
   ElseIf Dt = 2 Then
      DocumentTypeToString = "ใบเบิกวัตถุดิบ"
   ElseIf Dt = 3 Then
      DocumentTypeToString = "ใบโอนวัตถุดิบ"
   ElseIf Dt = 4 Then
      DocumentTypeToString = "ใบปรับยอด"
   End If
End Function
Public Function PoRoTypeToString(DocumentType As Long) As String
   If DocumentType = 100 Then
      PoRoTypeToString = MapText("ใบรับเข้าวัตถุดิบ")
   ElseIf DocumentType = 101 Then
      PoRoTypeToString = MapText("ใบรับเข้าวัสดุอุปกรณ์")
   ElseIf DocumentType = 102 Then
      PoRoTypeToString = MapText("ใบรับเข้าจ่ายออกวัสดุอุปกรณ์")
   ElseIf DocumentType = 103 Then
      PoRoTypeToString = MapText("ใบรับเข้าทั่วไป")
   ElseIf DocumentType = 110 Then
      PoRoTypeToString = MapText("ใบรับคืนสินค้า (ซื้อ)")
   ElseIf DocumentType = 1000 Then
      PoRoTypeToString = MapText("PO สั่งซื้อวัตถุดิบ")
   ElseIf DocumentType = 1001 Then
      PoRoTypeToString = MapText("PO สั่งซื้อวัสดุอุปกรณ์")
   ElseIf DocumentType = 1002 Then
      PoRoTypeToString = MapText("PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์")
   ElseIf DocumentType = 1003 Then
      PoRoTypeToString = MapText("PO สั่งซื้อทั่วไป")
   End If

End Function

   
Public Function CommitTypeToFlag(Ct As Long) As String
   If Ct = 1 Then
      CommitTypeToFlag = "Y"
   ElseIf Ct = 2 Then
      CommitTypeToFlag = "N"
   Else
      CommitTypeToFlag = ""
   End If
End Function

Public Function ID2Orientation(TempID As OrientationSettings) As String
   If TempID = orLandscape Then
      ID2Orientation = "แนวนอน"
   Else
      ID2Orientation = "แนวตั้ง"
   End If
End Function

Public Function ID2PaperSize(TempID As PaperSizeSettings) As String
   If TempID = pprA4 Then
      ID2PaperSize = "A4"
   ElseIf TempID = pprLetter Then
      ID2PaperSize = "Letter"
   ElseIf TempID = pprFanfoldUS Then
      ID2PaperSize = "Us standard"
   ElseIf TempID = 177 Then
      ID2PaperSize = "1/2 Letter"
   Else
      ID2PaperSize = "A4"
   End If
End Function

Public Sub InitPaperSize(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2PaperSize(pprA4))
   C.ItemData(1) = pprA4

   C.AddItem (ID2PaperSize(pprLetter))
   C.ItemData(2) = pprLetter

   C.AddItem (ID2PaperSize(pprFanfoldUS))
   C.ItemData(3) = pprFanfoldUS
   
   C.AddItem (ID2PaperSize(177))
   C.ItemData(4) = 177
End Sub

Public Function PeriodTypeToText(ID As PERIOD_TYPE) As String
   If ID = DAILY_PERIOD Then
      PeriodTypeToText = MapText("รายวัน")
   ElseIf ID = MONTHLY_PERIOD Then
      PeriodTypeToText = MapText("รายเดือน")
   End If
End Function

Public Function RateTypeToText(ID As RATE_TYPE) As String
   If ID = RATE_FLAT Then
      RateTypeToText = MapText("แฟลต")
   ElseIf ID = RATE_STEP Then
      RateTypeToText = MapText("เสตป")
   ElseIf ID = RATE_TIER Then
      RateTypeToText = MapText("เทียร์")
   End If
End Function

Public Function VariableToText(ID As JOB_VARIABLE_TYPE) As String
   If ID = CAN_VAR Then
      VariableToText = MapText("ค่าภาชนะบรรจุ/หน่วย")
   ElseIf ID = LOSS_VAR Then
      VariableToText = MapText("%การสูญเสีย")
   ElseIf ID = OVERHEAD_VAR Then
      VariableToText = MapText("โอเวอร์เฮด/หน่วย")
   ElseIf ID = PMC_VAR Then
      VariableToText = MapText("PMC")
   ElseIf ID = RCM_VAR Then
      VariableToText = MapText("RMC")
   End If
End Function

Public Sub InitRateType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (RateTypeToText(RATE_FLAT))
   C.ItemData(1) = RATE_FLAT

   C.AddItem (RateTypeToText(RATE_STEP))
   C.ItemData(2) = RATE_STEP

   C.AddItem (RateTypeToText(RATE_TIER))
   C.ItemData(3) = RATE_TIER
End Sub

Public Sub InitPeriodType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PeriodTypeToText(DAILY_PERIOD))
   C.ItemData(1) = DAILY_PERIOD

   C.AddItem (PeriodTypeToText(MONTHLY_PERIOD))
   C.ItemData(2) = MONTHLY_PERIOD
End Sub

Public Sub LoadJobVariable(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CFormulaVariable
Dim ItemCount As Long
Dim I As Long
Dim TempData As CFormulaVariable

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   For I = 1 To 3
      Set TempData = New CFormulaVariable
      TempData.VARIABLE_ID = I
      TempData.VARIABLE_NAME = VariableToText(I)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.VARIABLE_NAME)
         C.ItemData(I) = TempData.VARIABLE_ID
      End If
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.VARIABLE_ID)))
      End If
      Set TempData = Nothing
   Next I
      
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitFontName(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("AngsanaUPC")
   C.ItemData(1) = 1
End Sub

Public Sub InitOrientation(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2Orientation(orLandscape))
   C.ItemData(1) = orLandscape

   C.AddItem (ID2Orientation(orPortrait))
   C.ItemData(2) = orPortrait
End Sub

Public Sub LoadAgeRange(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CAgeRange
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAgeRange
Dim I As Long

   Set D = New CAgeRange
   Set Rs = New ADODB.Recordset
   
   D.AGE_RANGE_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCountry
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.AGE_RANGE_NAME)
         C.ItemData(I) = TempData.AGE_RANGE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.AGE_RANGE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GetSystemParam(Cl As Collection, Key As String) As CSystemParam
Dim St As CSystemParam
   For Each St In Cl
      If St.PARAM_NAME = Key Then
         Set GetSystemParam = St
         Exit Function
      End If
   Next St
   
   Set St = New CSystemParam
   St.PARAM_VALUE = ""
   Set GetSystemParam = St
End Function

Public Sub LoadSystemParam(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim Sp As CSystemParam
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempSP As CSystemParam
Dim I As Long

   Set Sp = New CSystemParam
   Set Rs = New ADODB.Recordset
   
   Sp.PARAM_ID = -1
   Sp.OrderBy = 2
   Sp.OrderType = 2
   Call Sp.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempSP = New CSystemParam
      Call TempSP.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempSP.PARAM_NAME)
         C.ItemData(I) = TempSP.PARAM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempSP, TempSP.PARAM_NAME)
      End If
      
      Set TempSP = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set Sp = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAccessRight(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGroupRight
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupRight
Dim I As Long

   Set D = New CGroupRight
   Set Rs = New ADODB.Recordset
   
   D.GROUP_RIGHT_ID = -1
   D.GROUP_ID = GroupID
   Call D.QueryData3(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGroupRight
      Call TempData.PopulateFromRS3(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RIGHT_ITEM_NAME)
         C.ItemData(I) = TempData.GROUP_RIGHT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadFeature(C As ComboBox, Optional Cl As Collection = Nothing, Optional FeatureType As Long = -1, Optional KeyType As Long = -1, Optional FeatureStatus As String = "", Optional BillDirectFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CFeature
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFeature
Dim I As Long

   Set D = New CFeature
   Set Rs = New ADODB.Recordset
   
   D.FEATURE_ID = -1
   D.FEATURE_CODE = ""
   D.QueryFlag = -1
   D.FEATURE_TYPE = FeatureType
   D.FEATURE_STATUS = FeatureStatus
   D.BILL_DIRECT_FLAG = BillDirectFlag
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFeature
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.FEATURE_DESC)
         C.ItemData(I) = TempData.FEATURE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, TempData.FEATURE_CODE)
         Else
            Call Cl.add(TempData, Trim(str(TempData.FEATURE_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadFormulaType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CFormulaType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFormulaType
Dim I As Long

   Set D = New CFormulaType
   Set Rs = New ADODB.Recordset
   
   D.FORMULA_TYPE_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFormulaType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.FORMULA_TYPE_NAME)
         C.ItemData(I) = TempData.FORMULA_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.FORMULA_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'==
Public Sub LoadReason(C As ComboBox, Optional Cl As Collection = Nothing, Optional Area As Long = -1)
On Error GoTo ErrorHandler
Dim D As CReason
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReason
Dim I As Long

   Set D = New CReason
   Set Rs = New ADODB.Recordset
   
   D.REASON_ID = -1
   D.REASON_NAME = ""
   D.Area = Area
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReason
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.REASON_NAME)
         C.ItemData(I) = TempData.REASON_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLayout(C As ComboBox, Optional Cl As Collection = Nothing, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLayout
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLayout
Dim I As Long

   Set D = New CLayout
   Set Rs = New ADODB.Recordset
   
   D.LAY_OUT_ID = -1
   D.LAY_OUT_NAME = ""
   D.LOCATION_ID = LocationID
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLayout
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.LAY_OUT_NAME)
         C.ItemData(I) = TempData.LAY_OUT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.LAY_OUT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDoType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDoType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoType
Dim I As Long

   Set D = New CDoType
   Set Rs = New ADODB.Recordset
   
   D.DO_TYPE_ID = -1
   D.DO_TYPE_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DO_TYPE_NAME)
         C.ItemData(I) = TempData.DO_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_TYPE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadFeatureType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CFeatureType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFeatureType
Dim I As Long

   Set D = New CFeatureType
   Set Rs = New ADODB.Recordset
   
   D.FEATURE_TYPE_ID = -1
   D.FEATURE_TYPE_NO = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFeatureType
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.FEATURE_TYPE_NAME)
         C.ItemData(I) = TempData.FEATURE_TYPE_ID
      End If
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.FEATURE_TYPE_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub CopySubLotItem(SourceCol As Collection, DestColl As Collection)
Dim R As CSubLotItem

Set DestColl = Nothing
Set DestColl = New Collection

   For Each R In SourceCol
      If R.Flag <> "D" Then
         Call DestColl.add(R)
      End If
   Next R
End Sub

Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CLotItem
Dim Ei As CLotItem

   For Each Ei In Ivd.ImportExports
      If Ei.LINK_ID = GuiID Then
         Set GetExportItem = Ei
         Exit Function
      End If
   Next Ei
End Function
Public Function GetExportItemWh(Ivd As CInventoryWHDoc, GuiID As Long) As CLotItemWH
Dim Ei As CLotItemWH

   For Each Ei In Ivd.ImportExports
      If Ei.LINK_ID = GuiID Then
         Set GetExportItemWh = Ei
         Exit Function
      End If
   Next Ei
End Function
Public Function GetLotItemsWH(Ivd As CInventoryWHDoc, GuiID As Long) As CLotItemWH
Dim Ei As CLotItemWH

   For Each Ei In Ivd.C_LotItemsWH
      If Ei.LINK_ID = GuiID Then
         Set GetLotItemsWH = Ei
         Exit Function
      End If
   Next Ei
End Function


Public Sub LoadSoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional SocLevel As String = "")
On Error GoTo ErrorHandler
Dim D As CSoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSoc
Dim I As Long

   Set D = New CSoc
   Set Rs = New ADODB.Recordset
   
   D.SOC_ID = -1
   D.SOC_CODE = ""
   D.SOC_LEVEL = SocLevel
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         Call C.AddItem(TempData.SOC_DESC)
         C.ItemData(I) = TempData.SOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadResource(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CResource
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CResource
Dim I As Long

   Set D = New CResource
   Set Rs = New ADODB.Recordset
   
   D.RESOURCE_ID = -1
   D.RESOURCE_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CResource
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RESOURCE_NAME)
         C.ItemData(I) = TempData.RESOURCE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.RESOURCE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSex(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSex
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSex
Dim I As Long

   Set D = New CSex
   Set Rs = New ADODB.Recordset
   
   D.SEX_ID = -1
   D.SEX_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSex
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SEX_NAME)
         C.ItemData(I) = TempData.SEX_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadWorkStatus(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CWorkStatus
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CWorkStatus
Dim I As Long

   Set D = New CWorkStatus
   Set Rs = New ADODB.Recordset
   
   D.WORK_ID = -1
   D.WORK_NAME = ""
   D.OrderType = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CWorkStatus
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.WORK_NAME)
         C.ItemData(I) = TempData.WORK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBloodGroup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBloodGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBloodGroup
Dim I As Long

   Set D = New CBloodGroup
   Set Rs = New ADODB.Recordset
   
   D.BLOODGRP_ID = -1
   D.BLOODGRP_NAME = ""
D.OrderType = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBloodGroup
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BLOODGRP_NAME)
         C.ItemData(I) = TempData.BLOODGRP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReligious(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CReligious
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReligious
Dim I As Long

   Set D = New CReligious
   Set Rs = New ADODB.Recordset
   
   D.RELIGIOUS_ID = -1
   D.RELIGIOUS_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReligious
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RELIGIOUS_NAME)
         C.ItemData(I) = TempData.RELIGIOUS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMarital(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CMarital
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMarital
Dim I As Long

   Set D = New CMarital
   Set Rs = New ADODB.Recordset
   
   D.MARITAL_ID = -1
   D.MARITAL_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMarital
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.MARITAL_NAME)
         C.ItemData(I) = TempData.MARITAL_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMilitary(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CMilitary
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMilitary
Dim I As Long

   Set D = New CMilitary
   Set Rs = New ADODB.Recordset
   
   D.MILITARY_ID = -1
   D.MILITARY_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMilitary
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.MILITARY_NAME)
         C.ItemData(I) = TempData.MILITARY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBankAccount(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBankAccount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankAccount
Dim I As Long

   Set D = New CBankAccount
   Set Rs = New ADODB.Recordset
   
   D.BANK_ID = -1
   D.BANK_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBankAccount
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BANK_NAME)
         C.ItemData(I) = TempData.BANK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadResignReason(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CResignReason
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CResignReason
Dim I As Long

   Set D = New CResignReason
   Set Rs = New ADODB.Recordset
   
   D.RSGRESON_ID = -1
   D.RSGRESON_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CResignReason
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RSGRESON_NAME)
         C.ItemData(I) = TempData.RSGRESON_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDocumentType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDocumentType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDocumentType
Dim I As Long

   Set D = New CDocumentType
   Set Rs = New ADODB.Recordset
   
   D.DOCTYPE_ID = -1
   D.DOCTYPE_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDocumentType
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DOCTYPE_NAME)
         C.ItemData(I) = TempData.DOCTYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDependencyType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDependencyType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDependencyType
Dim I As Long

   Set D = New CDependencyType
   Set Rs = New ADODB.Recordset
   
   D.DEPENDENCY_TYPE_ID = -1
   D.DEPENDENCY_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDependencyType
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DEPENDENCY_NAME)
         C.ItemData(I) = TempData.DEPENDENCY_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadQualificationType(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CQualificationType
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CQualificationType
Dim I As Long

   Set D = New CQualificationType
   Set Rs = New ADODB.Recordset
   
   D.QUALIFICATION_TYPE_ID = -1
   D.QUALIFICATION_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CQualificationType
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.QUALIFICATION_NAME)
         C.ItemData(I) = TempData.QUALIFICATION_TYPE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitThaiMonth(C As ComboBox)
   Dim I As Long
   
   C.Clear

   For I = 0 To 12
      C.AddItem (IntToThaiMonth(I))
      C.ItemData(I) = I
   Next
End Sub
Public Sub InitThaiYear(C As ComboBox)
   Dim I As Long
   
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   C.AddItem ("2543")
   C.AddItem ("2544")
   C.AddItem ("2545")
   C.AddItem ("2546")
   C.AddItem ("2547")
   C.AddItem ("2548")
   C.AddItem ("2549")
   C.AddItem ("2550")
   C.AddItem ("2551")
   C.AddItem ("2552")
End Sub
Public Sub InitEmpReceivableOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขใบยืม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่ยืม"))
   C.ItemData(2) = 2

   C.AddItem (MapText("ผู้ยืม"))
   C.ItemData(3) = 3
   
   
   C.AddItem (MapText("จำนวนที่ยืม"))
   C.ItemData(4) = 4
End Sub

Public Sub LoadSliptAdd(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CMonthlyAdd
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMonthlyAdd
Dim I As Long

   Set D = New CMonthlyAdd
   Set Rs = New ADODB.Recordset
   
   D.MONTHLY_ADD_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMonthlyAdd
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.MONTHLY_ADD_NAME)
         C.ItemData(I) = TempData.MONTHLY_ADD_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSliptSub(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CMonthlySub
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMonthlySub
Dim I As Long

   Set D = New CMonthlySub
   Set Rs = New ADODB.Recordset
   
   D.MONTHLY_SUB_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMonthlySub
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.MONTHLY_SUB_NAME)
         C.ItemData(I) = TempData.MONTHLY_SUB_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadProcess(C As ComboBox, Optional Cl As Collection = Nothing, Optional ProcessID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CProcess
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CProcess
Dim I As Long

   Set D = New CProcess
   Set Rs = New ADODB.Recordset
   
   D.PROCESS_ID = ProcessID
   D.PROCESS_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CProcess
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PROCESS_NAME)
         C.ItemData(I) = TempData.PROCESS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.PROCESS_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMachine(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CMachine
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMachine
Dim I As Long

   Set D = New CMachine
   Set Rs = New ADODB.Recordset
   
   D.MACHINE_ID = -1
   D.MACHINE_NAME = ""
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMachine
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.MACHINE_NAME)
         C.ItemData(I) = TempData.MACHINE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.MACHINE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadRefDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CInventoryDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDoc
Dim I As Long

   Set D = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   D.COMMIT_FLAG = "Y"
   D.INVENTORY_DOC_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DOCUMENT_NO)
         C.ItemData(I) = TempData.INVENTORY_DOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.INVENTORY_DOC_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Public Sub InitJobOrderBy(C As ComboBox)
'   C.Clear
'
'   C.AddItem ("")
'   C.ItemData(0) = 0
'
'   C.AddItem (MapText("เลขที่ใบสั่งผลิต"))
'   C.ItemData(1) = 1
'
'   C.AddItem (MapText("รายละเอียดงาน"))
'   C.ItemData(2) = 2
'
'   C.AddItem (MapText("วันที่งาน"))
'   C.ItemData(3) = 3
'
'   C.AddItem (MapText("หมายเลขแบท"))
'   C.ItemData(4) = 4
'
'   C.AddItem (MapText("วันที่เริ่มงาน"))
'   C.ItemData(1) = 5
'
'   C.AddItem (MapText("วันที่เสร็จงาน"))
'   C.ItemData(2) = 6
'
'   C.AddItem (MapText("ผู้อนุมัติ"))
'   C.ItemData(3) = 7
'
'   C.AddItem (MapText("ผู้รับผิดชอบ"))
'   C.ItemData(4) = 8
'
'   C.AddItem (MapText("โปรเซส"))
'   C.ItemData(1) = 9
'
'   C.AddItem (MapText("หมายเลขเอกสารอ้างอิง"))
'   C.ItemData(2) = 10
'
'End Sub

Public Sub InitFormulaOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสูตร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รายละเอียดสูตร"))
   C.ItemData(2) = 2

   C.AddItem (MapText("วันที่สูตร"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ประเภทสูตร"))
   C.ItemData(4) = 4
End Sub
'===

Public Sub LoadFormula(C As ComboBox, Optional Cl As Collection = Nothing, Optional FormulaType As Long = -1, Optional PartItemID As Long = -1, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1, Optional CancelFlag As String = "N")
On Error GoTo ErrorHandler
Dim D As CFormula
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFormula
Dim I As Long

   Set D = New CFormula
   Set Rs = New ADODB.Recordset
   
   D.FORMULA_ID = -1
   D.QueryFlag = -1
   D.FORMULA_TYPE = FormulaType
   D.PART_ITEM_ID = PartItemID
   D.OrderBy = 1
   D.OrderType = 1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.CANCEL_FLAG = CancelFlag
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFormula
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.FORMULA_NO)
         C.ItemData(I) = TempData.FORMULA_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.FORMULA_ID)))
         Else
            Call Cl.add(TempData, TempData.FORMULA_NO)
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctFormulaNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error Resume Next 'เนื่องจากของเก่านั้นมีมี ซ้ำกัน แบบ ใช้ Adj กับ adj
Dim D As CFormula
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFormula
Dim I As Long

   Set D = New CFormula
   Set Rs = New ADODB.Recordset
   
   D.FORMULA_ID = -1
   D.QueryFlag = -1
   D.OrderBy = 1
   D.OrderType = 1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFormula
      Call TempData.PopulateFromRS(2, Rs)
   
'      If TempData.FORMULA_NO = "PM-138adj(04/04/55)" Or TempData.FORMULA_NO = "PM-138Adj(04/04/55)" Then
'         'Debug.Print
'      End If
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.FORMULA_NO)
         C.ItemData(I) = TempData.FORMULA_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.FORMULA_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
'   Exit Sub
   
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & "-" & TempData.FORMULA_NO
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadParameterProcess(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CParameterProcess
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CParameterProcess
Dim I As Long

   Set D = New CParameterProcess
   Set Rs = New ADODB.Recordset
   
   D.PARAMETER_PROCESS_ID = -1
   D.PARAMETER_PROCESS_NAME = ""
   D.OrderBy = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CParameterProcess
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PARAMETER_PROCESS_NAME)
         C.ItemData(I) = TempData.PARAMETER_PROCESS_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PARAMETER_PROCESS_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadParameterItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional Process As Long)
On Error GoTo ErrorHandler
Dim D As CParameterItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CParameterItem
Dim I As Long

   Set D = New CParameterItem
   Set Rs = New ADODB.Recordset
   
   D.PARAMETER_ITEM_ID = -1
   D.PROCESS_ID = Process
   D.OrderBy = 1
   D.SELECT_FLAG = "Y"
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CParameterItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PARAMETER_PROCESS_NAME)
         C.ItemData(I) = TempData.PARAMETER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PARAMETER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMoneyFamily(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CMoneyFamily
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMoneyFamily
Dim I As Long

   Set D = New CMoneyFamily
   Set Rs = New ADODB.Recordset
   
   D.MONEY_FAMILY_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMoneyFamily
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.MONEY_FAMILY_NAME)
         C.ItemData(I) = TempData.MONEY_FAMILY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.MONEY_FAMILY_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitCurrencyOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลข"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("อัตราแลกเปลี่ยน"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("สกุลเงินต้นทาง"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("สกุลเงินปลายทาง"))
   C.ItemData(5) = 5
End Sub
'===

Public Sub InitFormulaItemOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลขวัตถุดิบ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เปอร์เซ็นต์วัตถุดิบ"))
   C.ItemData(2) = 2

End Sub

Public Sub InitJobStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เสร็จแล้ว"))
   C.ItemData(1) = 1
   
   

   C.AddItem (MapText("ยังไม่เสร็จ"))
   C.ItemData(2) = 2

End Sub

Public Sub LoadJob(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long

   Set D = New CJob
   Set Rs = New ADODB.Recordset
   
   D.JOB_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.JOB_NO)
         C.ItemData(I) = TempData.JOB_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, str(TempData.JOB_ID))
         Else
            Call Cl.add(TempData, TempData.JOB_NO)
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadJobByJobNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1, Optional JobNo As String = "", Optional ProcessID As Long = -1)
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim m_Rs2 As ADODB.Recordset
Dim m_Rs3 As ADODB.Recordset
Dim m_Rs4 As ADODB.Recordset
Dim IsOK As Boolean

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

Dim Ua As CJob

   Set Ua = New CJob
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
   Set m_Rs4 = New ADODB.Recordset

   Ua.JOB_ID = -1
'   Ua.JOB_NO = JobNo
   Ua.JOB_NO_LIKE = JobNo
   Ua.PROCESS_ID = ProcessID
   Ua.FROM_DATE = FromDate
   Ua.TO_DATE = ToDate
   Ua.QueryFlag = 1
   Call Ua.QueryData(1, Rs, ItemCount)
   While Not Rs.EOF
'      I = I + 1
      Set Ua = New CJob
      Call Ua.PopulateFromRS(1, Rs)
      
      If Not (Ua Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(Ua, str(Ua.JOB_ID))
         Else
            Call Cl.add(Ua, Trim(Ua.JOB_NO))
         End If
      End If
      
    'INPUT ++++++++++++++++++++++++++
      Dim Inp As CJobInput
      Set Inp = New CJobInput
      Inp.JOB_INOUT_ID = -1
      Inp.JOB_ID = Ua.JOB_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
      
      Set Ua.Inputs = Nothing
      Set Ua.Inputs = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CJobInput
         Call Inp.PopulateFromRS(1, m_Rs2)
         Inp.Flag = "I"
         If Inp.TX_TYPE = "E" Then
            Call Ua.Inputs.add(Inp, Trim(str(Inp.PART_ITEM_ID)))
         End If
         Set Inp = Nothing
         m_Rs2.MoveNext
      Wend
'INPUT ++++++++++++++++++++++++++

'OUTPUT ++++++++++++++++++++++++++
      Dim Op As CJobInput
      Set Op = New CJobInput
      Op.JOB_INOUT_ID = -1
      Op.JOB_ID = Ua.JOB_ID
      
      Call Op.QueryData(1, m_Rs2, iCount)
      Set Op = Nothing
      
      Set Ua.Outputs = Nothing
      Set Ua.Outputs = New Collection
      While Not m_Rs2.EOF
       Set Op = New CJobInput
         Call Op.PopulateFromRS(1, m_Rs2)
         Op.Flag = "I"
         
         If Op.TX_TYPE = "I" Then
            Call Ua.Outputs.add(Op, Trim(str(Op.PART_ITEM_ID)))
         End If
         Set Op = Nothing
         m_Rs2.MoveNext
      Wend
'OUTPUT ++++++++++++++++++++++++++

'MACHINE USED TIME ++++++++++++++++++++++++++
      Dim EH As CJobResource
     Set EH = New CJobResource
      EH.JOB_ID = Ua.JOB_ID
      Call EH.QueryData(m_Rs1, iCount)
      Set EH = Nothing
      
      Set Ua.Machines = Nothing
      Set Ua.Machines = New Collection
      While Not m_Rs1.EOF
         Set EH = New CJobResource
         Call EH.PopulateFromRS(1, m_Rs1)
      
         EH.Flag = "I"
         If EH.MACHINE_NO <> "" Then
         Call Ua.Machines.add(EH)
         End If
         Set EH = Nothing
         m_Rs1.MoveNext
      Wend
      'MACHINE TIME USED ++++++++++++++++++++++++++
        
'PERSON USED TIME ++++++++++++++++++++++++++
      Dim Ep As CJobResource
     Set Ep = New CJobResource
      Ep.JOB_ID = Ua.JOB_ID
      Call Ep.QueryData(m_Rs1, iCount)
      Set Ep = Nothing
      
      Set Ua.Peoples = Nothing
      Set Ua.Peoples = New Collection
      While Not m_Rs1.EOF
         Set Ep = New CJobResource
         Call Ep.PopulateFromRS(2, m_Rs1)
      
         Ep.Flag = "I"
         If Ep.EMP_ID > 0 Then
         Call Ua.Peoples.add(Ep)
         End If
         Set Ep = Nothing
         m_Rs1.MoveNext
      Wend
      'PERSON TIME USED ++++++++++++++++++++++++++
            
        
     'PARAMETER TIME ++++++++++++++++++++++++++
      Dim PP As CJobParameter
     Set PP = New CJobParameter
      PP.JOB_ID = Ua.JOB_ID
      Call PP.QueryData(1, m_Rs1, iCount)
      Set PP = Nothing
      
      Set Ua.Parameters = Nothing
      Set Ua.Parameters = New Collection
      While Not m_Rs1.EOF
         Set PP = New CJobParameter
         Call PP.PopulateFromRS(1, m_Rs1)
      
         PP.Flag = "I"
         Call Ua.Parameters.add(PP)
         Set PP = Nothing
         m_Rs1.MoveNext
      Wend
      'PARAMETER USED ++++++++++++++++++++++++++
            
      Dim Jv As CJobVerify
     Set Jv = New CJobVerify
      Jv.JOB_ID = Ua.JOB_ID
      Call Jv.QueryData(m_Rs1, iCount)
      Set Jv = Nothing
            
      'Job verify
      Set Ua.Verifies = Nothing
      Set Ua.Verifies = New Collection
      While Not m_Rs1.EOF
         Set Jv = New CJobVerify
         Call Jv.PopulateFromRS(1, m_Rs1)
      
         Jv.Flag = "I"
         Call Ua.Verifies.add(Jv)
         Set Jv = Nothing
         m_Rs1.MoveNext
      Wend
      'Job verify

      'InventoryWhDoc++++++++++++++++++++++++++
      Dim IWD As CInventoryWHDoc
      Dim LotItemWh As CLotItemWH
      Dim Lot As cLot
      Dim LTD As CLotDoc
      Dim PD As CPalletDoc

      Set IWD = New CInventoryWHDoc
      If Not Rs.EOF Then
         IWD.INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      Else
          IWD.INVENTORY_WH_DOC_ID = 0
      End If
      
      If IWD.INVENTORY_WH_DOC_ID < 0 Then
          IWD.QueryFlag = 0
      Else
         IWD.QueryFlag = 1
      End If
      
      If IWD.INVENTORY_WH_DOC_ID = -1 Then
         IWD.INVENTORY_WH_DOC_ID = 0
      End If
      
      Call glbDaily.QueryInventoryWhDocForImBulk(IWD, m_Rs2, iCount, IsOK, glbErrorLog)
      Call IWD.PopulateFromRS(1, m_Rs2)
      If m_Rs2.RecordCount > 0 Then
      IWD.Flag = "I"
      If Ua.InventoryWhDoc Is Nothing Then
         Set Ua.InventoryWhDoc = New Collection
      End If
      
      Call Ua.InventoryWhDoc.add(IWD, Trim(IWD.DOCUMENT_NO))
      Set IWD = Nothing
      Else
       Set Ua.InventoryWhDoc = Nothing
      End If
      'InventoryWhDoc ++++++++++++++++++++++++++
      
''      'InventoryWhDocInput++++++++++++++++++++++++++
''      Dim IWD_Input As CInventoryWHDoc
''
''      Set IWD_Input = New CInventoryWHDoc
''      If Not Rs.EOF Then
''         IWD_Input.INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID_INPUT"), -1)
''      Else
''          IWD_Input.INVENTORY_WH_DOC_ID = 0
''      End If
''
''
''
''      If IWD_Input.INVENTORY_WH_DOC_ID < 0 Then
''          IWD_Input.QueryFlag = 0
''      Else
''         IWD_Input.QueryFlag = 1
''      End If
''
''      If IWD_Input.INVENTORY_WH_DOC_ID = -1 Then
''         IWD_Input.INVENTORY_WH_DOC_ID = 0
''      End If
''
''      Call glbDaily.QueryInventoryWhDocInput(IWD_Input, m_Rs2, iCount, IsOK, glbErrorLog)
''      Call IWD_Input.PopulateFromRS(1, m_Rs2)
''      If m_Rs2.RecordCount > 0 Then
''      IWD_Input.Flag = "I"
''
''       If Ua.InventoryWhDocInput Is Nothing Then
''         Set Ua.InventoryWhDocInput = New Collection
''      End If
''      If Ua.InventoryWhDocInput.Count = 0 Then
''         Call Ua.InventoryWhDocInput.add(IWD_Input)
''      End If
''      Set IWD_Input = Nothing
''      Else
''       Set Ua.InventoryWhDocInput = Nothing
''      End If
''      'InventoryWhDocInput ++++++++++++++++++++++++++
   
'      Set TempData = Nothing
      Rs.MoveNext
   Wend
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadJobByJobNo2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1, Optional JobNo As String = "", Optional ProcessID As Long = -1)
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim m_Rs1 As ADODB.Recordset
Dim m_Rs2 As ADODB.Recordset
Dim m_Rs3 As ADODB.Recordset
Dim m_Rs4 As ADODB.Recordset
Dim IsOK As Boolean

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

Dim Ua As CJob

   Set Ua = New CJob
   Set Rs = New ADODB.Recordset
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
   Set m_Rs4 = New ADODB.Recordset

   Ua.JOB_ID = -1
   Ua.JOB_NO_LIKE = JobNo
   Ua.PROCESS_ID = ProcessID
   Ua.FROM_DATE = FromDate
   Ua.TO_DATE = ToDate
   Ua.QueryFlag = 1
   Call Ua.QueryData(1, Rs, ItemCount)
   While Not Rs.EOF
'      I = I + 1
      Set Ua = New CJob
      Call Ua.PopulateFromRS(1, Rs)
      
      If Not (Ua Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(Ua, str(Ua.JOB_ID))
         Else
            Call Cl.add(Ua, Trim(Ua.JOB_NO))
         End If
      End If
      
    'INPUT ++++++++++++++++++++++++++
      Dim Inp As CJobInput
      Set Inp = New CJobInput
      Inp.JOB_INOUT_ID = -1
      Inp.JOB_ID = Ua.JOB_ID
      Call Inp.QueryData(1, m_Rs2, iCount)
      Set Inp = Nothing
      
      Set Ua.Inputs = Nothing
      Set Ua.Inputs = New Collection
      While Not m_Rs2.EOF
       Set Inp = New CJobInput
         Call Inp.PopulateFromRS(1, m_Rs2)
         Inp.Flag = "I"
         If Inp.TX_TYPE = "E" Then
            Call Ua.Inputs.add(Inp)
         End If
         Set Inp = Nothing
         m_Rs2.MoveNext
      Wend
'INPUT ++++++++++++++++++++++++++

'OUTPUT ++++++++++++++++++++++++++
      Dim Op As CJobInput
      Set Op = New CJobInput
      Op.JOB_INOUT_ID = -1
      Op.JOB_ID = Ua.JOB_ID
      
      Call Op.QueryData(1, m_Rs2, iCount)
      Set Op = Nothing
      
      Set Ua.Outputs = Nothing
      Set Ua.Outputs = New Collection
      While Not m_Rs2.EOF
       Set Op = New CJobInput
         Call Op.PopulateFromRS(1, m_Rs2)
         Op.Flag = "I"
         
         If Op.TX_TYPE = "I" Then
            Call Ua.Outputs.add(Op)
         End If
         Set Op = Nothing
         m_Rs2.MoveNext
      Wend
'OUTPUT ++++++++++++++++++++++++++

'MACHINE USED TIME ++++++++++++++++++++++++++
      Dim EH As CJobResource
     Set EH = New CJobResource
      EH.JOB_ID = Ua.JOB_ID
      Call EH.QueryData(m_Rs1, iCount)
      Set EH = Nothing
      
      Set Ua.Machines = Nothing
      Set Ua.Machines = New Collection
      While Not m_Rs1.EOF
         Set EH = New CJobResource
         Call EH.PopulateFromRS(1, m_Rs1)
      
         EH.Flag = "I"
         If EH.MACHINE_NO <> "" Then
         Call Ua.Machines.add(EH)
         End If
         Set EH = Nothing
         m_Rs1.MoveNext
      Wend
      'MACHINE TIME USED ++++++++++++++++++++++++++
        
'PERSON USED TIME ++++++++++++++++++++++++++
      Dim Ep As CJobResource
     Set Ep = New CJobResource
      Ep.JOB_ID = Ua.JOB_ID
      Call Ep.QueryData(m_Rs1, iCount)
      Set Ep = Nothing
      
      Set Ua.Peoples = Nothing
      Set Ua.Peoples = New Collection
      While Not m_Rs1.EOF
         Set Ep = New CJobResource
         Call Ep.PopulateFromRS(2, m_Rs1)
      
         Ep.Flag = "I"
         If Ep.EMP_ID > 0 Then
         Call Ua.Peoples.add(Ep)
         End If
         Set Ep = Nothing
         m_Rs1.MoveNext
      Wend
      'PERSON TIME USED ++++++++++++++++++++++++++
            
        
     'PARAMETER TIME ++++++++++++++++++++++++++
      Dim PP As CJobParameter
     Set PP = New CJobParameter
      PP.JOB_ID = Ua.JOB_ID
      Call PP.QueryData(1, m_Rs1, iCount)
      Set PP = Nothing
      
      Set Ua.Parameters = Nothing
      Set Ua.Parameters = New Collection
      While Not m_Rs1.EOF
         Set PP = New CJobParameter
         Call PP.PopulateFromRS(1, m_Rs1)
      
         PP.Flag = "I"
         Call Ua.Parameters.add(PP)
         Set PP = Nothing
         m_Rs1.MoveNext
      Wend
      'PARAMETER USED ++++++++++++++++++++++++++
            
      Dim Jv As CJobVerify
     Set Jv = New CJobVerify
      Jv.JOB_ID = Ua.JOB_ID
      Call Jv.QueryData(m_Rs1, iCount)
      Set Jv = Nothing
            
      'Job verify
      Set Ua.Verifies = Nothing
      Set Ua.Verifies = New Collection
      While Not m_Rs1.EOF
         Set Jv = New CJobVerify
         Call Jv.PopulateFromRS(1, m_Rs1)
      
         Jv.Flag = "I"
         Call Ua.Verifies.add(Jv)
         Set Jv = Nothing
         m_Rs1.MoveNext
      Wend
      'Job verify
      Rs.MoveNext
   Wend
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCashDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CCashDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashDoc
Dim I As Long

   Set D = New CCashDoc
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("CASH_DOC_ID", -1)
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, str(TempData.GetFieldValue("CASH_DOC_ID")))
         Else
            Call Cl.add(TempData, str(TempData.GetFieldValue("DOCUMENT_NO")))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSerialNo(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   D.JOB_INOUT_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(1, Rs)
   If Len(TempData.SERIAL_NUMBER) > 0 Then
   I = I + 1
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SERIAL_NUMBER)
         C.ItemData(I) = TempData.JOB_INOUT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.JOB_INOUT_ID))
      End If
   End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBank(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CBank
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBank
Dim I As Long

   Set D = New CBank
   Set Rs = New ADODB.Recordset
   
   D.BANK_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBank
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BANK_NAME)
         C.ItemData(I) = TempData.BANK_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.BANK_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadBankBranch(C As ComboBox, Optional Cl As Collection = Nothing, Optional BankID As Long)
On Error GoTo ErrorHandler
Dim D As CBankBranch
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBankBranch
Dim I As Long

   Set D = New CBankBranch
   Set Rs = New ADODB.Recordset
   
   D.BBRANCH_ID = -1
   D.BANK_ID = BankID
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBankBranch
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.BBRANCH_NAME)
         C.ItemData(I) = TempData.BBRANCH_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.BBRANCH_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitPaymentOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่บัญชี"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitPaymentType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentType2Text(1))
   C.ItemData(1) = 1
   
   C.AddItem (PaymentType2Text(2))
   C.ItemData(2) = 2

   C.AddItem (PaymentType2Text(3))
   C.ItemData(3) = 3
End Sub

Public Function PaymentType2Text(ID As Long) As String
   If ID = 1 Then
      PaymentType2Text = "เงินสด"
   ElseIf ID = 2 Then
      PaymentType2Text = "เงินโอน"
   ElseIf ID = 3 Then
      PaymentType2Text = "เช็ค"
   Else
      PaymentType2Text = ""
   End If
End Function
Public Sub InitPaymentTypeEx(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentTypeToText(CASH_PMT))
   C.ItemData(1) = CASH_PMT
   
   C.AddItem (PaymentTypeToText(CHECK_PMT))
   C.ItemData(2) = CHECK_PMT

   C.AddItem (PaymentTypeToText(BANKTRF_PMT))
   C.ItemData(3) = BANKTRF_PMT

   C.AddItem (PaymentTypeToText(CASHRET_PMT))
   C.ItemData(4) = CASHRET_PMT
End Sub

Public Sub InitCurrencyExOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่"))
   C.ItemData(1) = 1

   C.AddItem (MapText("US$"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("EURO"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("YEN"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("S$"))
   C.ItemData(5) = 5
End Sub
'===
Public Sub LoadMoneyFamilyEx(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("US$"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("EURO"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("YEN"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("S$"))
   C.ItemData(4) = 4
End Sub
'===

Public Sub LoadInventoryBalanceByLocation(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)
   
   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.LOCATION_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryBalanceByPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)
   
   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryBalanceEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)
   
   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(6, Rs)
      
      If Not (C Is Nothing) Then
      End If
'If TempData.LOCATION_ID = 106 And TempData.PART_ITEM_ID = 123 Then
''Debug.Print
'End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryPartBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBalanceAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceAccum
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Dim NewDate As Date

   Set D = New CBalanceAccum
   Set Rs = New ADODB.Recordset

   NewDate = DateAdd("D", -1, FromDate)
   
   D.FROM_DATE = -1
   D.TO_DATE1 = InternalDateToDate(DateToStringIntHi(NewDate))
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(10, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBalanceAccum
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplierSpec(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CSupplierSpec
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplierSpec
Dim I As Long

   Set D = New CSupplierSpec
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_SPEC_ID = -1
   D.SUPPLIER_ID = SupplierID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierSpec
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_NO)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.SUPPLIER_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartItemSpec(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPartItemSpec
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItemSpec
Dim I As Long

   Set D = New CPartItemSpec
   Set Rs = New ADODB.Recordset

   D.PARTITEM_SPEC_ID = -1
   D.PART_ITEM_ID = PartItemID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPartItemSpec
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_NO)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPackaging(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPackaging
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackaging
Dim I As Long

   Set D = New CPackaging
   Set Rs = New ADODB.Recordset
   
   D.PACKAGING_ID = -1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackaging
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PACKAGING_NAME)
         C.ItemData(I) = TempData.PACKAGING_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PACKAGING_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPurchaseExpense(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CPurchaseExpense
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPurchaseExpense
Dim I As Long

   Set D = New CPurchaseExpense
   Set Rs = New ADODB.Recordset
   
   D.PUREXP_ID = -1
   D.OrderType = 1
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPurchaseExpense
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PUREXP_NAME)
         C.ItemData(I) = TempData.PUREXP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PUREXP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitBillSubType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ขายเชื่อ")
   C.ItemData(1) = 10

   C.AddItem ("ขายสด")
   C.ItemData(2) = 21
End Sub

Public Sub InitBillingBillSubType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ขายเชื่อ")
   C.ItemData(1) = 1

   C.AddItem ("ขายสด")
   C.ItemData(2) = 2
End Sub

Public Sub initDataMode(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("น้ำหนัก")
   C.ItemData(1) = 1

   C.AddItem ("มูลค่า")
   C.ItemData(2) = 2
End Sub

Public Function ID2PackageType(TempID As PACKAGE_TYPE) As String
   If TempID = PACKAGE_BAG Then
      ID2PackageType = "BAG"
   ElseIf TempID = PACKAGE_BULK Then
      ID2PackageType = "BULK"
   ElseIf TempID = PACKAGE_OTH Then
      ID2PackageType = "OTHER"
   Else
      ID2PackageType = "OTHER N/A"
   End If
End Function

Public Function ID2RateType(TempID As DO_RATE_TYPE) As String
   If TempID = RATE_CUSTOM Then
      ID2RateType = "ราคาเฉพาะลูกค้า"
   ElseIf TempID = RATE_MASTER Then
      ID2RateType = "ราคากลาง"
   Else
      ID2RateType = "ราคาอื่น ๆ"
   End If
End Function

Public Sub InitDoRateType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2RateType(RATE_CUSTOM))
   C.ItemData(1) = RATE_CUSTOM

   C.AddItem (ID2RateType(RATE_MASTER))
   C.ItemData(2) = RATE_MASTER
End Sub
Public Sub InitDoRateType2(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("มารับเอง (ราคากลาง)")
   C.ItemData(1) = 1
   
   C.AddItem ("รวมค่าขนส่ง (ราคากลาง+ค่าขนส่ง+อื่นๆ)")
   C.ItemData(2) = 2
   
   C.AddItem ("แยกค่าขนส่ง (ราคากลาง+อื่นๆ)")
   C.ItemData(3) = 3
   
   C.ListIndex = 0
End Sub


Public Sub InitPackageType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2PackageType(PACKAGE_BAG))
   C.ItemData(1) = PACKAGE_BAG

   C.AddItem (ID2PackageType(PACKAGE_BULK))
   C.ItemData(2) = PACKAGE_BULK

   C.AddItem (ID2PackageType(PACKAGE_OTH))
   C.ItemData(3) = PACKAGE_OTH
End Sub
Public Function DeliveryType(Ind As Long) As String
   If Ind = 1 Then
      DeliveryType = "รถ BAG"
   ElseIf Ind = 2 Then
      DeliveryType = "รถ BULK"
   ElseIf Ind = 3 Then
      DeliveryType = "เหมาเที่ยว"
   Else
      DeliveryType = ""
   End If
End Function
Public Function DeliveryUnit(Ind As Long) As String
   If Ind = 1 Then
      DeliveryUnit = "ถุง"
   ElseIf Ind = 2 Then
      DeliveryUnit = "กก."
   ElseIf Ind = 3 Then
      DeliveryUnit = "เที่ยว"
   ElseIf Ind = 10 Then
      DeliveryUnit = "ถุง"
   ElseIf Ind = 21 Then
      DeliveryUnit = "กก."
   Else
      DeliveryUnit = ""
   End If
End Function
Public Sub InitDeliveryType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ถุง")
   C.ItemData(1) = 1

   C.AddItem ("กิโลกรัม")
   C.ItemData(2) = 2

   C.AddItem ("เที่ยว")
   C.ItemData(3) = 3
End Sub
Public Sub InitParcelType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ParcelTypeToText(PARCEL_BAG))
   C.ItemData(1) = PARCEL_BAG

   C.AddItem (ParcelTypeToText(PARCEL_BULK))
   C.ItemData(2) = PARCEL_BULK

   C.AddItem (ParcelTypeToText(PARCEL_ALL))
   C.ItemData(3) = PARCEL_ALL
End Sub

Public Sub InitParcelTypeEx(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ParcelTypeToText(PARCEL_BAG))
   C.ItemData(1) = PARCEL_BAG

   C.AddItem (ParcelTypeToText(PARCEL_BULK))
   C.ItemData(2) = PARCEL_BULK
End Sub

Public Sub InitRatioType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (RatioTypeToText(RATIO_COST))
   C.ItemData(1) = RATIO_COST

   C.AddItem (RatioTypeToText(RATIO_QUANTITY))
   C.ItemData(2) = RATIO_QUANTITY

   C.AddItem (RatioTypeToText(RATIO_RAW))
   C.ItemData(3) = RATIO_RAW
   
   C.AddItem (RatioTypeToText(RATIO_VARY))
   C.ItemData(4) = RATIO_VARY
   
   C.AddItem (RatioTypeToText(RATIO_PERCENT))
   C.ItemData(5) = RATIO_PERCENT
End Sub

Public Sub LoadPartLocationTxTypeAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(15, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartTxTypeAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PART_ITEM_ID = PartItemID
   D.OrderBy = 1
   Call D.QueryData(16, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(16, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartTxTypeAmountWH(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PART_ITEM_ID = PartItemID
   D.OrderBy = 1
   Call D.QueryData(2, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.LOT_DOC_ID & "-" & TempData.BIN_NAME & "-" & TempData.LOCK_NAME)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartTxTypeDateAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(29, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(29, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.DOCUMENT_DATE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPartTxTypeDocTypeAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1, Optional PartGroup As Long = -1, Optional DocType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   D.DOCUMENT_TYPE = DocType
   D.PART_GROUP_ID = PartGroup
   Call D.QueryData(18, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.DOCUMENT_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartTxTypeDocTypeAmountYYYYMM(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long, Optional PartType As Long = -1, Optional PartGroup As Long = -1, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   D.DOCUMENT_TYPE_SET = DocTypeSet
   D.PART_GROUP_ID = PartGroup
   Call D.QueryData(27, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(27, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.YYYYMM))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartTxTypeDocTypeCusTypeAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1, Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   D.PART_GROUP_ID = PartGroup
   Call D.QueryData(24, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(24, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.CUSTOMER_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctCusType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1, Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.PART_GROUP_ID = PartGroup
   D.OrderBy = 1
   Call D.QueryData(25, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(25, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadLocationPartTxTypeDocTypeAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(20, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(20, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.DOCUMENT_TYPE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaidAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAPPaidAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional EffectiveFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.EFFECTIVE_FLAG = EffectiveFlag
   Call D.QueryData(100, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(100, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional SourceType As Long = 1, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If SourceType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf SourceType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBillingDiscountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional SourceType As Long = 1, Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CBillingDiscount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDiscount
Dim I As Long

   Set D = New CBillingDiscount
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DISCOUNT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDiscount
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBillingFreeAmountByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(18, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(18, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("BILLING_DOC_ID"))))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPaidAmountByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPaidAmountByCustomerDocDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim TempData2 As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   Call D.QueryData(122, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(122, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)) & "-" & Trim(str(TempData.DOCUMENT_DATE)))
      End If
'      If Not (Cl Is Nothing) Then
'         Set TempData2 = GetObject("CReceiptItem", Cl, Trim(str(TempData.CUSTOMER_ID)) & "-" & Trim(str(TempData.DOCUMENT_DATE)), False)
'         If TempData2 Is Nothing Then
'            Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)) & "-" & Trim(str(TempData.DOCUMENT_DATE)))
'         Else
'            TempData2.PAID_AMOUNT = TempData2.PAID_AMOUNT + TempData.PAID_AMOUNT
'            TempData2.DISCOUNT_AMOUNT = TempData2.DISCOUNT_AMOUNT + TempData.DISCOUNT_AMOUNT
'         End If
'      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCashTrnAmountByCustomerDocDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim TempData2 As CCashTran
Dim I As Long

   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DOC_DATE", FromDate)
   Call D.SetFieldValue("TO_DOC_DATE", ToDate)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(18, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(18, Rs)
   

      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("CUSTOMER_ID"))) & "-" & Trim(str(TempData.GetFieldValue("DOCUMENT_DATE"))))
      End If

      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDiscountAmountByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDiscount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDiscount
Dim I As Long

   Set D = New CBillingDiscount
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DISCOUNT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDiscount
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDiscountAmountByAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDiscount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDiscount
Dim I As Long

   Set D = New CBillingDiscount
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DISCOUNT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDiscount
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.ACCOUNT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional DateType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If DateType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf DateType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadTotalPriceByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional Ind As Long = 8)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(Ind, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBillingDocDistinctAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional CustomerCode As String = "", Optional FromCustomerCode As String, Optional ToCustomerCode As String)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set D = New CBillingDoc
   D.BILLING_DOC_ID = -1
   D.CUSTOMER_CODE = CustomerCode
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(2, Rs, ItemCount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCOUNT_NO)
         C.ItemData(I) = TempData.ACCOUNT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.ACCOUNT_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountByAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional DateType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If DateType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf DateType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.ACCOUNT_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSellPriceByDocTypeSubTypeReceiptTypeAcc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.ACCOUNT_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadReceiptByDocTypeSubTypeReceiptTypeAcc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.ACCOUNT_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadReceiptByCustomerDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, DateToStringInt(TempData.DOCUMENT_DATE) & "-" & TempData.CUSTOMER_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDueDateInterval(C As ComboBox, Optional Cl As Collection, Optional SumInDue As Long = 0)
Dim MM As CMaxMin
   
   If SumInDue = 0 Then
   '===
      Set MM = New CMaxMin
      MM.MIN = -30
      MM.MAX = 0
      Call Cl.add(MM)
      Set MM = Nothing
      
      Set MM = New CMaxMin
      MM.MIN = -60
      MM.MAX = -30
      Call Cl.add(MM)
      Set MM = Nothing
   
      Set MM = New CMaxMin
      MM.MIN = -999999
      MM.MAX = -60
      Call Cl.add(MM)
      Set MM = Nothing
   Else
      Set MM = New CMaxMin
      MM.MIN = -999999
      MM.MAX = 0
      Call Cl.add(MM)
      Set MM = Nothing
   End If
   
   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 15
   Call Cl.add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 15
   MM.MAX = 30
   Call Cl.add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 30
   MM.MAX = 60
   Call Cl.add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 60
   MM.MAX = 9999999
   Call Cl.add(MM)
   Set MM = Nothing
End Sub

Public Sub LoadDueDateInterval2(C As ComboBox, Optional Cl As Collection, Optional SumInDue As Long = 0)
Dim MM As CMaxMin
   '===
   If SumInDue = 0 Then
      Set MM = New CMaxMin
      MM.MIN = -30
      MM.MAX = 0
      Call Cl.add(MM)
      Set MM = Nothing
      
      Set MM = New CMaxMin
      MM.MIN = -60
      MM.MAX = -30
      Call Cl.add(MM)
      Set MM = Nothing
   
      Set MM = New CMaxMin
      MM.MIN = -999999
      MM.MAX = -60
      Call Cl.add(MM)
      Set MM = Nothing
   Else
      Set MM = New CMaxMin
      MM.MIN = -999999
      MM.MAX = 0
      Call Cl.add(MM)
      Set MM = Nothing
   End If
   
   '===
   Set MM = New CMaxMin
   MM.MIN = 0
   MM.MAX = 90
   Call Cl.add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 90
   MM.MAX = 180
   Call Cl.add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 180
   MM.MAX = 365
   Call Cl.add(MM)
   Set MM = Nothing

   Set MM = New CMaxMin
   MM.MIN = 365
   MM.MAX = 9999999
   Call Cl.add(MM)
   Set MM = Nothing
End Sub

Public Sub LoadTotalPriceByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(10, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaymentByType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PaymentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPaymentItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPaymentItem
Dim I As Long

   Set D = New CPaymentItem
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PAYMENT_TYPE = PaymentType
   D.OrderBy = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPaymentItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PAYMENT_TYPE & "-" & TempData.TX_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPaymentByDocTypeSubType(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PaymentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPaymentItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPaymentItem
Dim I As Long

   Set D = New CPaymentItem
   Set Rs = New ADODB.Recordset

   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PAYMENT_TYPE = PaymentType
   D.OrderBy = 1
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPaymentItem
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.DOCUMENT_TYPE & "-" & TempData.RECEIPT_TYPE & "-" & TempData.PAYMENT_TYPE & "-" & TempData.TX_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMaster(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterType As MASTER_TYPE, Optional TempID1 As Long = -1, Optional TempID2 As Long = -1, Optional KeyType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMasterRef
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long

   Set D = New CMasterRef
   Set Rs = New ADODB.Recordset
   
   D.KEY_ID = -1
   D.MASTER_AREA = MasterType
   D.TEMP_ID1 = TempID1
   D.TEMP_ID2 = TempID2
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.KEY_NAME)
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.KEY_CODE))
         Else
            Call Cl.add(TempData, Trim(str(TempData.KEY_ID)))
         End If
         
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctAccountInCashTrn(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("BANK_ACCOUNT"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCashTranAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(5, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.GetFieldValue("CUSTOMER_ID") & "-" & TempData.GetFieldValue("BANK_ACCOUNT") & "-" & DateToStringInt(TempData.GetFieldValue("TX_DATE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCashTrnAmount(Ct As CCashTran, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ct.SetFieldValue("CASH_TRAN_ID", -1)
   Call Ct.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.GetFieldValue("KEY_NAME"))
'         C.ItemData(I) = TempData.GetFieldValue("KEY_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.GetFieldValue("BANK_ACCOUNT") & "-" & TempData.GetFieldValue("TX_TYPE"))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBankAccountInCashTrn(Ct As CCashTran, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ct.SetFieldValue("CASH_TRAN_ID", -1)
   Call Ct.SetFieldValue("ORDER_TYPE", 1)
   Call Ct.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
'         C.AddItem (TempData.GetFieldValue("KEY_NAME"))
'         C.ItemData(I) = TempData.GetFieldValue("KEY_ID")
      End If
      
      If Not (Cl Is Nothing) Then
'''Debug.Print TempData.GetFieldValue("BANK_ACCOUNT") & "-" & TempData.GetFieldValue("BANK_ID") & "-" & TempData.GetFieldValue("BANK_BRANCH")
         Call Cl.add(TempData, TempData.GetFieldValue("BANK_ACCOUNT") & "-" & TempData.GetFieldValue("BANK_ID") & "-" & TempData.GetFieldValue("BANK_BRANCH"))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitReportCashTx(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
End Sub
Public Sub LoadCostRaw(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCostRaw
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCostRaw
Dim I As Long

   Set D = New CCostRaw
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CCostRaw
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.GetFieldValue("PART_ITEM_ID"))
         C.ItemData(I) = TempData.GetFieldValue("PART_ITEM_ID")
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("PART_ITEM_ID"))))
      End If
            
      Set TempData = Nothing
      
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function GeneratePartItemSet(Col As Collection) As String
Dim CR As CCostRaw
Dim I As Long
Dim TempStr As String
   
   I = 0
   TempStr = "("
   For Each CR In Col
      If CR.Flag <> "D" Then
         I = I + 1
         TempStr = TempStr & CR.GetFieldValue("PART_ITEM_ID")
         If I < Col.Count Then
            TempStr = TempStr & ", "
         End If
      End If
   Next CR
   TempStr = TempStr & ")"
   
   GeneratePartItemSet = TempStr
End Function

Public Sub LoadPrtItemSet(C As ComboBox, Optional Cl As Collection = Nothing, Optional PrtItemSetID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPrtItemSet
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPrtItemSet
Dim I As Long

   Set D = New CPrtItemSet
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("PRTITEM_SET_ID", -1)
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPrtItemSet
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("PART_ITEM_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctSellPartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CustType As Long = -1, Optional DocTypeSet As String = "")
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.CUSTOMER_TYPE = CustType
   D.DocTypeSet = DocTypeSet
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSaleByProductSale(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CustType As Long = -1, Optional DocTypeSet As String = "")
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.CUSTOMER_TYPE = CustType
   D.DocTypeSet = DocTypeSet
   Call D.QueryData(12, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.ACCEPT_BY & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleProductSale(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CustType As Long = -1, Optional DocTypeSet As String = "")
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.CUSTOMER_TYPE = CustType
   D.DocTypeSet = DocTypeSet
   Call D.QueryData(13, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctJobProduct(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional ProcessType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long

   Set D = New CJob
   Set Rs = New ADODB.Recordset
   
   D.JOB_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessType
   Call D.QueryData(4, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(4, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadProductPartUsed(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional ProcessID As Long = -1, Optional PartGroup As Long = -1, Optional PartType As Long = -1, Optional TxType As String, Optional PrtItemSet As String)
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   
   D.JOB_INOUT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_GROUP_ID = PartGroup
   D.PART_TYPE_ID = PartType
   D.PROCESS_ID = ProcessID
   D.TX_TYPE = TxType
   D.PART_ITEM_SET = PrtItemSet
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.JOB_PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadConfigDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CConfigDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CConfigDoc
Dim I As Long

   Set D = New CConfigDoc
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CConfigDoc
      Call TempData.PopulateFromRS(1, Rs)
      
      TempData.Flag = "I"
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.GetFieldValue("CONFIG_DOC_CODE"))
         C.ItemData(I) = TempData.GetFieldValue("CONFIG_DOC_TYPE")
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("CONFIG_DOC_TYPE"))))
      End If
            
      Set TempData = Nothing
      
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctCustomerPicture(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCustomerPicture
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomerPicture
Dim I As Long
   
   Set D = New CCustomerPicture
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("CUSTOMER_PICTURE_TYPE", HEAD_ACCOUNT)
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomerPicture
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("CUSTOMER_ID"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCashTranAmountByCustDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(7, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.GetFieldValue("CUSTOMER_ID") & "-" & TempData.GetFieldValue("PAYMENT_TYPE") & "-" & DateToStringInt(TempData.GetFieldValue("TX_DATE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSumCashTranAmountByCustDate2(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(8, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.GetFieldValue("CUSTOMER_ID") & "-" & TempData.GetFieldValue("PAYMENT_TYPE") & "-" & DateToStringInt(TempData.GetFieldValue("TX_DATE")))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadReceiptByBillingDocID(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(12, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function getReceiptByBillingDocIDRef(Optional BillingdocID As Long) As String
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

getReceiptByBillingDocIDRef = ""
   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.DO_ID = BillingdocID
   Call D.QueryData(123, Rs, ItemCount)
   
'   getReceiptByBillingDocIDRef = Rs.RecordCount
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(123, Rs)
      If getReceiptByBillingDocIDRef = "" Then
         getReceiptByBillingDocIDRef = TempData.DOCUMENT_NO
      Else
         getReceiptByBillingDocIDRef = getReceiptByBillingDocIDRef & "," & TempData.DOCUMENT_NO
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function
Public Sub LoadSumCashTranAmountByBillingDocID(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(9, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(9, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("BILLING_DOC_ID"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctAccountInCashTran(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(10, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("BANK_ACCOUNT"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDoIDFromReceiptItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long
Dim PrevID As Long
Dim Seq As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_ITEM_DUE_DATE = FromDate
   D.TO_ITEM_DUE_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(13, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not Rs.EOF Then
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(13, Rs)
      PrevID = TempData.DO_ID
      Seq = 1
      Set TempData = Nothing
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(13, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If PrevID <> TempData.DO_ID Then
         Seq = 1
         PrevID = TempData.DO_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.DO_ID & "-" & Seq)
'''Debug.Print TempData.DO_ID & "-" & Seq
         Seq = Seq + 1
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMaxMinReceiptDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long
Dim PrevID As Long
Dim Seq As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_ITEM_DUE_DATE = FromDate
   D.TO_ITEM_DUE_DATE = ToDate
   D.DOCUMENT_TYPE = 2
   Call D.QueryData(14, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
      
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
            
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadChequeFromReceipt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("PAYMENT_TYPE", 3)
   Call D.SetFieldValue("TX_TYPE", "I")
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(11, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(11, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
'''Debug.Print TempData.GetFieldValue("BILLING_DOC_ID")
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function ID2ExportDocTypeSet(ID As Long) As String
   If ID = 0 Then
      ID2ExportDocTypeSet = "(2, 20)"
   ElseIf ID = 1 Then
      ID2ExportDocTypeSet = "(2)"
   ElseIf ID = 2 Then
      ID2ExportDocTypeSet = "(20)"
   End If
End Function

Public Sub InitExportDocTypeSet(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใบเบิกวัสดุอุปกรณ์")
   C.ItemData(1) = 1

   C.AddItem ("รับเข้า-จ่ายออก")
   C.ItemData(2) = 2
End Sub

Public Function ID2ImportDocTypeSet(ID As Long) As String
   If ID = 0 Then
      ID2ImportDocTypeSet = "(19, 23,20)"
   ElseIf ID = 1 Then
      ID2ImportDocTypeSet = "(23)"
   ElseIf ID = 2 Then
      ID2ImportDocTypeSet = "(19)"
   ElseIf ID = 3 Then
      ID2ImportDocTypeSet = "(20)"
   End If
End Function

Public Sub InitImportDocTypeSet(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รับเข้าทั่วไป")
   C.ItemData(1) = 1

   C.AddItem ("รับเข้าพัสดุ-อุปกรณ์")
   C.ItemData(2) = 2

   C.AddItem ("รับเข้า-จ่ายออก")
   C.ItemData(3) = 3
End Sub

Public Sub InitMemoNoteOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่สร้าง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เสร็จ"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันที่เสร็จจริง"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("ผู้สร้าง"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("ผู้ได้รับมอบหมาย"))
   C.ItemData(5) = 5
   
   C.AddItem (MapText("ประเภท"))
   C.ItemData(6) = 6
   
   C.AddItem (MapText("สถานะ"))
   C.ItemData(7) = 7
End Sub
'===
Public Sub LoadUserAccount(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserAccount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserAccount
Dim I As Long

   Set D = New CUserAccount
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.USER_NAME)
         C.ItemData(I) = TempData.USER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.USER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
'===
Public Sub LoadUserAccountByRealName(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserAccount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserAccount
Dim I As Long

   Set D = New CUserAccount
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
            C.AddItem (TempData.REAL_NAME)
            C.ItemData(I) = TempData.USER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.USER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadUserAccountByName(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserAccount
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserAccount
Dim I As Long

   Set D = New CUserAccount
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserAccount
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.USER_NAME)
         C.ItemData(I) = TempData.USER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.USER_NAME))
      End If
      
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadJobProductStdAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional ProcessID As Long, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   D.TX_TYPE = TxType
   D.OrderBy = 1
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.JOB_PART_ITEM_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE)
''Debug.Print TempData.JOB_PART_ITEM_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCashTranAmountByCust(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional BillingDocType As Long = -1, Optional ReceiptType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("BILLING_DOC_TYPE", BillingDocType)
   Call D.SetFieldValue("RECEIPT_TYPE", ReceiptType)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(12, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(12, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("CUSTOMER_ID"))))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBalanceCashCheque(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(13, Rs, ItemCount)
   
   Set D = Nothing
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl Is Nothing) Then
         Set D = GetObject("CCashTran", Cl, "1", False)
         If D Is Nothing Then
            Set D = New CCashTran
            Call Cl.add(D, "1")
         End If
         If TempData.GetFieldValue("TX_TYPE") = "I" Then
            Call D.SetFieldValue("AMOUNT", D.GetFieldValue("AMOUNT") + TempData.GetFieldValue("AMOUNT"))
            Call D.SetFieldValue("FEE_AMOUNT", D.GetFieldValue("FEE_AMOUNT") + TempData.GetFieldValue("FEE_AMOUNT"))
         Else
            Call D.SetFieldValue("AMOUNT", D.GetFieldValue("AMOUNT") - TempData.GetFieldValue("AMOUNT"))
            Call D.SetFieldValue("FEE_AMOUNT", D.GetFieldValue("FEE_AMOUNT") - TempData.GetFieldValue("FEE_AMOUNT"))
         End If
      End If
      ''Debug.Print (TempData.GetFieldValue("TX_TYPE"))
      ''Debug.Print (TempData.GetFieldValue("AMOUNT"))
      ''Debug.Print (TempData.GetFieldValue("FEE_AMOUNT"))
      ''Debug.Print (D.GetFieldValue("AMOUNT"))
      ''Debug.Print (D.GetFieldValue("FEE_AMOUNT"))
      
      Set D = Nothing
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPayinByCustDateAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(14, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistictCstItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional Parameter As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCostItemRaw
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCostItemRaw
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCostItemRaw
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("PARAM_PROCESS_ID", Parameter)
   Call D.QueryData(3, Rs, ItemCount)
   
   Set D = Nothing
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCostItemRaw
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set D = Nothing
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCstItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional Parameter As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCostItemRaw
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCostItemRaw
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCostItemRaw
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("PARAM_PROCESS_ID", Parameter)
   Call D.QueryData(4, Rs, ItemCount)
   
   Set D = Nothing
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCostItemRaw
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("FEED_ID") & "-" & TempData.GetFieldValue("PART_ITEM_ID")))
      End If
      
      Set D = Nothing
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumReceiptPayByChequeDoc(Cl As Collection, CustomerCode As String, FromCustomerCode As String, ToCustomerCode As String, CustType As Long, CustGrade As Long) '
On Error GoTo ErrorHandler
Dim Cd As CChequeDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CChequeDoc
Dim TempData2 As CChequeDoc

Dim I As Long

   
   Set Cd = New CChequeDoc
   Set Rs = New ADODB.Recordset
   
'    AF.ACC_FOL_ID = -1
'    AF.CANCEL_FLAG = "N"
   Cd.CHEQUE_DOC_ID = -1
   Cd.CUSTOMER_CODE = CustomerCode
    Cd.FROM_CUSTOMER_CODE = FromCustomerCode
   Cd.TO_CUSTOMER_CODE = ToCustomerCode
   Cd.CUSTOMER_TYPE = CustType
   Cd.CUSTOMER_TYPE = CustGrade
    
   Call Cd.QueryData(7, Rs, ItemCount)
   

   If Not (Cl Is Nothing) Then
   Set Cl = Nothing
   Set Cl = New Collection
   End If

   While Not Rs.EOF

      Set TempData2 = New CChequeDoc
      Call TempData2.PopulateFromRS(7, Rs)
      Set TempData = GetObject("CChequeDoc", Cl, TempData2.CUSTOMER_ID & "-" & TempData2.DOCUMENT_NO, False)
      If TempData Is Nothing Then
         Set TempData = New CChequeDoc
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.SUM_PAID_AMOUNT = TempData2.SUM_PAID_AMOUNT
         TempData.DOCUMENT_NO = TempData2.DOCUMENT_NO
         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.DOCUMENT_NO)
     End If

      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set Cd = Nothing
   
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub InitReport8_19Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   C.AddItem (MapText("รหัสอาหาร"))
   C.ItemData(1) = 1

End Sub
Public Sub InitReport9_1_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เบอร์สินค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("เวลาเริ่มบรรจุ"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReport9_2_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReport9_2_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(3) = 3
End Sub
Public Sub InitExpenseOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub
Public Sub LoadSumExpenseDetail(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CExpenseDetail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CExpenseDetail
Dim I As Long

   Set D = New CExpenseDetail
   Set Rs = New ADODB.Recordset

   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CExpenseDetail
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("EXPENSE_DETAIL_TYPE"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumJobInputPartParam(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long
   
   Set D = New CJobInput
   Set Rs = New ADODB.Recordset

   D.JOB_INOUT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.TX_TYPE = "E"
   Call D.QueryData(9, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(9, Rs)
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PART_NO)
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_PART_ITEM_ID & "-" & TempData.PARAM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadJobProductActualAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional ProcessID As Long)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long

   Set D = New CJob
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   D.OrderBy = 1
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleAmountByPart(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CommitFlag As String = "", Optional LocationID As Long, Optional PartType As Long = -1, Optional PartGroup As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.COMMIT_FLAG = CommitFlag
   D.LOCATION_ID = LocationID
   D.PART_TYPE = PartType
   D.OrderBy = 1
   Call D.QueryData(24, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(24, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumCashTranAccountCusAccList(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.QueryData(15, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(15, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitAPChequeStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เช็คในมือ"))
   C.ItemData(1) = 102

   C.AddItem (MapText("รับเช็คแล้ว"))
   C.ItemData(2) = 100

   C.AddItem (MapText("CLEARING แล้ว"))
   C.ItemData(3) = 101
End Sub
Public Sub InitPoType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("PO สั่งซื้อวัตถุดิบ"))
   C.ItemData(1) = 1000

   C.AddItem (MapText("PO สั่งซื้อวัสดุอุปกรณ์"))
   C.ItemData(2) = 1001

   C.AddItem (MapText("PO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์"))
   C.ItemData(3) = 1002

   C.AddItem (MapText("PO สั่งซื้อทั่วไป"))
   C.ItemData(4) = 1003
End Sub
Public Sub InitPrType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("RO สั่งซื้อวัตถุดิบ"))
   C.ItemData(1) = 100

   C.AddItem (MapText("RO สั่งซื้อวัสดุอุปกรณ์"))
   C.ItemData(2) = 101

   C.AddItem (MapText("RO สั่งซื้อ รับเข้าจ่ายออกวัสดุอุปกรณ์"))
   C.ItemData(3) = 102

   C.AddItem (MapText("RO สั่งซื้อทั่วไป"))
   C.ItemData(4) = 103
End Sub

Public Sub InitCheckLayout(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("MP-UOB")
   C.ItemData(1) = 1

   C.AddItem ("MP-KTB")
   C.ItemData(2) = 2

   C.AddItem ("QMC-BKK")
   C.ItemData(3) = 3

   C.AddItem ("MGP-SCB")
   C.ItemData(4) = 4

   C.AddItem ("DTS-TFB")
   C.ItemData(5) = 5

   C.AddItem ("MH-TFB")
   C.ItemData(6) = 6
End Sub
Public Sub InitDocumentTypeSup(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ใบรับเข้าวัตถุดิบ"))
   C.ItemData(1) = 100

   C.AddItem (MapText("ใบรับเข้าวัสดุอุปกรณ์"))
   C.ItemData(2) = 101

   C.AddItem (MapText("ใบรับเข้าจ่ายออกวัสดุอุปกรณ์"))
   C.ItemData(3) = 102
   
   C.AddItem (MapText("ใบรับเข้าทั่วไป"))
   C.ItemData(4) = 103
End Sub

Public Sub LoadAPTotalPriceByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CSupItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupItem
Dim I As Long

   Set D = New CSupItem
   Set Rs = New ADODB.Recordset
   
   D.SUP_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(100, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupItem
      Call TempData.PopulateFromRS(100, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAPTotalPriceByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional Ind As Long = 8)
On Error GoTo ErrorHandler
Dim D As CSupItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupItem
Dim I As Long

   Set D = New CSupItem
   Set Rs = New ADODB.Recordset
   
   D.SUP_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(101, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupItem
      Call TempData.PopulateFromRS(101, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAPPaidAmountByCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PigID As Long = -1, Optional LocationID As Long = -1, Optional DocumentCat As Long = -1, Optional EffectiveFlag As String = "")
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.EFFECTIVE_FLAG = EffectiveFlag
   Call D.QueryData(101, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(101, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadAPSumCashTranBySupDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional TxType As String = "")
On Error GoTo ErrorHandler
Dim D As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CCashTran
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("FROM_DATE", FromDate)
   Call D.SetFieldValue("TO_DATE", ToDate)
   Call D.SetFieldValue("TX_TYPE", TxType)
   Call D.SetFieldValue("BILLING_DOC_TYPE", 8)
   Call D.SetFieldValue("ORDER_BY", 1)
   Call D.QueryData(16, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(16, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.GetFieldValue("SUPPLIER_ID") & "-" & TempData.GetFieldValue("BILLING_DOC_ID"))
      End If
         
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBillingDocDistinctDocumentNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional Area As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set D = New CBillingDoc
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(7, Rs, ItemCount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DOCUMENT_NO)
         C.ItemData(I) = TempData.BILLING_DOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistiinctSupTypeInCheque(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional ChequeStatus As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCheque
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCheque
Dim I As Long

   Set D = New CCheque
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("CHEQUE_ID", -1)
   Call D.SetFieldValue("FROM_DATE2", FromDate)
   Call D.SetFieldValue("TO_DATE2", ToDate)
   Call D.SetFieldValue("DIRECTION", 2)
   Call D.SetFieldValue("CHEQUE_STATUS", ChequeStatus)
   Call D.APCheckStatus2Flag
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCheque
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.GetFieldValue("SUPPLIER_TYPE"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitARIntervalType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 1
   
   C.AddItem (MapText("แบบที่ 1"))
   C.ItemData(1) = 1

   C.AddItem (MapText("แบบที่ 2"))
   C.ItemData(2) = 2
End Sub

Public Sub LoadBillingDocDistinctSup(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupCode As String = "", Optional SupplierType As Long = -1, Optional SupplierGrade As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Set D = New CBillingDoc
   D.BILLING_DOC_ID = -1
   D.SUPPLIER_CODE = SupCode
   D.SUPPLIER_TYPE = SupplierType
   D.SUPPLIER_GRADE = SupplierGrade
   Call D.QueryData(101, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(101, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_CODE)
         C.ItemData(I) = TempData.SUPPLIER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDnCnAmountBySupplier(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional DateType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If DateType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf DateType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(103, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(103, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadReceiptByDocTypeSubTypeReceiptTypeSup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   
   Call D.QueryData(104, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(104, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.SUPPLIER_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReceiptByDocTypeSubTypeReceiptTypeSupEx(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   
   Call D.QueryData(106, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(106, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.SUPPLIER_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSellPriceByDocTypeSubTypeReceiptTypeSup(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CSupItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupItem
Dim I As Long

   Set D = New CSupItem
   Set Rs = New ADODB.Recordset
   
   D.SUP_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(103, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupItem
      Call TempData.PopulateFromRS(103, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.SUPPLIER_ID & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_SUBTYPE & "-" & TempData.RECEIPT_TYPE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitEvaluatePayOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสซัพพลายเออร์"))
   C.ItemData(2) = 2
   
End Sub
Public Sub InitExportType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ข้อมูลซัพพลายเออร์"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ข้อมูลสินค้า/วัตถุดิบ"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("ข้อมูล PO ซัพพลายเออร์"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("ข้อมูลการซื้อจากซัพพลายเออร์"))
   C.ItemData(4) = 4
   
   C.AddItem (MapText("ข้อมูลเพิ่มหนี้/ลดหนี้จากซัพพลายเออร์"))
   C.ItemData(5) = 5
   
End Sub
Public Sub InitExportPostType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ข้อมูลเช็คที่ซัพพลายเออร์รับแล้ว"))
   C.ItemData(1) = 1
   
   
End Sub

Public Sub LoadFeatureTotalPriceByBill(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional Ind As Long = 8)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(30, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(30, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInventoryDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CInventoryDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDoc
Dim I As Long

   Set D = New CInventoryDoc
   Set Rs = New ADODB.Recordset
   
   D.INVENTORY_DOC_ID = -1
   D.DOCUMENT_NO = ""
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.DOCUMENT_NO)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadInventoryDocMaxDateBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional DocumentType As Long = 0, Optional ByRef MaxDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CInventoryWHDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryWHDoc
Dim I As Long

   Set D = New CInventoryWHDoc
   Set Rs = New ADODB.Recordset
   
   D.INVENTORY_WH_DOC_ID = -1
   D.DOCUMENT_TYPE = DocumentType
   D.COMMIT_FLAG = ""
   Call D.QueryData(8, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryWHDoc
      Call TempData.PopulateFromRS(8, Rs)
      
'      If Not (Cl Is Nothing) Then
'
'         Call Cl.add(TempData)
'      End If
      
      MaxDate = TempData.MAX_BALANCE_DATE
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBillTransportByTypePrice(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional GroupType As Long = -1, Optional Ind As Long = 2)
On Error GoTo ErrorHandler
Dim D As CBillTransport
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillTransport
Dim I As Long

   Set D = New CBillTransport
   Set Rs = New ADODB.Recordset
   
   D.BILL_TRANSPORT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   If Ind = 3 Then
      D.DOCUMENT_TYPE = 1
      D.FROM_DATE = -1
      D.TO_DATE = -1
      D.FROM_DATE_BD = FromDate
      D.TO_DATE_BD = ToDate
   End If
   Call D.QueryData(Ind, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillTransport
      Call TempData.PopulateFromRS(Ind, Rs)
   
      If Not (C Is Nothing) Then
      End If

      If Not (Cl Is Nothing) Then
         If GroupType = 1 Then 'by KEY_CODE
           If Not TempData.KEY_CODE = "" Then
            Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)) & "-" & Trim(str(TempData.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim(TempData.KEY_CODE))
          End If
         ElseIf GroupType = 2 Then 'by WEIGHT_PER_UNIT
            Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)) & "-" & Trim(str(TempData.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim(str(TempData.WEIGHT_PER_UNIT)))
          ElseIf GroupType = 3 Then 'by bill
            Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadBillTransportForShow(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional GroupType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillTransport
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillTransport
Dim I As Long

   Set D = New CBillTransport
   Set Rs = New ADODB.Recordset
   
   D.BILL_TRANSPORT_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillTransport
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If GroupType = 1 Then 'by KEY_CODE
           If Not TempData.KEY_CODE = "" Then
            Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)) & "-" & Trim(str(TempData.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim(TempData.KEY_CODE))
          End If
         ElseIf GroupType = 2 Then 'by WEIGHT_PER_UNIT
            Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID)) & "-" & Trim(str(TempData.BILL_TRANSPORT_ITEM_ID)) & "-" & Trim(str(TempData.WEIGHT_PER_UNIT)))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDoItemByLastBilling(C As ComboBox, Optional Cl As Collection = Nothing, Optional CusID As Long, Optional PartID As Long)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long
Dim mainBill As CDoItem

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.CUSTOMER_ID = CusID
   D.PART_ITEM_ID = PartID
   D.OrderType = 2
   Call D.QueryData(48, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(48, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInventoryWhDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional BillingdocID As Long, Optional DocumentType As Long, Optional EditPriceFlag As String)
On Error GoTo ErrorHandler
Dim D As CInventoryWHDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryWHDoc
Dim I As Long

   Set D = New CInventoryWHDoc
   Set Rs = New ADODB.Recordset
   
   D.INVENTORY_WH_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.BILLING_DOC_ID = BillingdocID
   D.DOCUMENT_TYPE = DocumentType
   If EditPriceFlag <> "Y" Then 'กรณีมีการแก้ไขราคาที่ SO ก็ให้ดึงเงือนไขเดิมขึ้นมาด้วย
      D.LOAD_FLAG = "I" 'เอาเฉพาะใบที่ขึ้นอาหารเรียบร้อยแล้วและรอออกใบส่งของ
      D.SUCCESS_FLAG = "N" 'เอาเฉพาะใบที่ยังไม่ออกใบส่งของ
   End If
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryWHDoc
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, str(TempData.BILLING_DOC_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPackAmountFromLWH(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional BillingdocID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset

   D.BILLING_DOC_ID = BillingdocID
'   D.FROM_DATE = FromDate
'   D.TO_DATE = ToDate
   Call D.QueryData(11, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(11, Rs)

      If Not (C Is Nothing) Then
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.FEATURE_ID) & "-" & str(TempData.BILLING_DOC_ID)))
'        If TempData.PART_ITEM_ID > 0 Then
'         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.BILLING_DOC_ID)))
'       Else
'          Call Cl.add(TempData, Trim(str(TempData.FEATURE_ID) & "-" & str(TempData.BILLING_DOC_ID)))
'       End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSumAmountWhByAnimal(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional BillingdocID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CLotItemWH
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItemWH
Dim I As Long

   Set D = New CLotItemWH
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(11, Rs, ItemCount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItemWH
      Call TempData.PopulateFromRS(11, Rs)

      If Not (C Is Nothing) Then
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID) & "-" & str(TempData.BILLING_DOC_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPartItemExt(C As ComboBox, Optional Cl As Collection = Nothing, Optional PartTypeID As Long = -1, Optional PigFlag As String = "", Optional PigType As String = "", Optional KeyType As Byte = 1)
'On Error GoTo ErrorHandler
On Error Resume Next
Dim D As CPartItemExt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPartItemExt
Dim I As Long

   Set D = New CPartItemExt
   Set Rs = New ADODB.Recordset
   
   D.PART_ITEM_ID = -1
   D.PART_TYPE = PartTypeID
   D.PIG_FLAG = PigFlag
   D.PIG_TYPE = PigType
   D.OrderBy = 5
   D.OrderType = 1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CPartItemExt
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.PART_DESC & "  (" & TempData.PART_NO & ")")
         C.ItemData(I) = TempData.PART_ITEM_ID
      End If

      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.PART_NO))
         End If
      End If
            
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplierExt(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Byte = 1)
On Error GoTo ErrorHandler
Dim D As CSupplierExt
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CSupplierExt
Dim I As Long

   Set D = New CSupplierExt
   D.SUPPLIER_ID = -1
   D.OrderBy = 1
   D.OrderType = 1
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(Rs, ItemCount)
   Else
      If Not Rs.EOF Then
         Rs.MoveFirst
      End If
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplierExt
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_NAME)
         C.ItemData(I) = TempData.SUPPLIER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
         ElseIf KeyType = 2 Then
            Call Cl.add(TempData, Trim(TempData.SUPPLIER_CODE))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupItemByRo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional SupplierCode As String)
On Error GoTo ErrorHandler
Dim D As CSupItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupItem
Dim I As Long
   Set D = New CSupItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.SUPPLIER_CODE = SupplierCode
   D.OrderBy = 1
   Call D.QueryData(106, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupItem
      Call TempData.PopulateFromRS(106, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SUPPLIER_ID & "-" & TempData.PO_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadSupItemComeIn(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional SupplierCode As String, Optional KeySel As Long, Optional FuncName As String)
On Error GoTo ErrorHandler
Dim D As CSupItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupItem
Dim I As Long
   Set D = New CSupItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.SUPPLIER_CODE = SupplierCode
   D.OrderBy = 1


   If KeySel = 1 Then
      D.DOCUMENT_TYPE_SET = "(100,101,102,103)"
      Call D.QueryData(116, Rs, ItemCount)
   Else
      Call D.QueryData(110, Rs, ItemCount)
   End If

   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupItem
      If KeySel = 1 Then
         Call TempData.PopulateFromRS(116, Rs)
      Else
         Call TempData.PopulateFromRS(110, Rs)
      End If
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If KeySel = 1 Then
            Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE))
         Else
            Call Cl.add(TempData, Trim(TempData.SUPPLIER_ID & "-" & TempData.PO_ID & "-" & TempData.PART_ITEM_ID))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & "   error LoadSupItemComeIn : " & TempData.PART_ITEM_ID & "  " & TempData.DOCUMENT_DATE '& "  " & TempData.DOCUMENT_NO
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupItemPartItemByRo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional SupplierCode As String, Optional DocTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CSupItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupItem
Dim I As Long
   Set D = New CSupItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.SUPPLIER_CODE = SupplierCode
   D.OrderBy = 1
   Call D.QueryData(111, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupItem
      Call TempData.PopulateFromRS(111, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PO_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Public Sub LoadCheque(C As ComboBox, Optional Cl As Collection = Nothing)
'On Error GoTo ErrorHandler
On Error Resume Next
Dim D As CCheque
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCheque
Dim I As Long

   Set D = New CCheque
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("CHEQUE_ID", -1)
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      Set TempData = New CCheque
      Call TempData.PopulateFromRS(1, Rs)

      If Not (C Is Nothing) Then
         I = I + 1
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.GetFieldValue("CHEQUE_NO")))
      End If
            
      Set TempData = Nothing
      If Not Rs.EOF Then Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadBillingDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.DOCUMENT_NO = ""
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.DOCUMENT_NO)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLotItemFindCostByBilling(C As ComboBox, Optional Cl As Collection = Nothing, Optional BillingdocID As Long)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim TempData2 As CLotItem
Dim I As Long
Dim Key As String

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.LOT_ITEM_ID = -1
   D.BILLING_DOC_ID = BillingdocID
   Call D.QueryData(36, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(36, Rs)
     If TempData.TX_TYPE = "E" Then
         If Not (C Is Nothing) Then
         End If
         
         If TempData.PART_TYPE = 21 Then 'ถ้าเป็น Bulk
            Key = "21"
         Else 'ถ้าเป็นถุง
           Key = "B"
         End If
         
         If Not (Cl Is Nothing) Then
             Set TempData2 = GetObject("CLotItem", Cl, Trim(str(TempData.BILLING_DOC_ID) & "-" & str(TempData.PART_ITEM_ID) & "-" & Key), False)
            If Not TempData2 Is Nothing Then
              TempData2.TOTAL_INCLUDE_PRICE = TempData2.TOTAL_INCLUDE_PRICE + TempData.TOTAL_INCLUDE_PRICE
              TempData2.TX_AMOUNT = TempData2.TX_AMOUNT + TempData.TX_AMOUNT
            Else
               Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_ID) & "-" & str(TempData.PART_ITEM_ID) & "-" & Key))
            End If
         End If
         
      End If
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleOrder(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional IdWh As Long = -1, Optional KeyType As Long = 1)
On Error GoTo ErrorHandler
Dim D As CSaleOrder
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSaleOrder
Dim I As Long

   Set D = New CSaleOrder
   Set Rs = New ADODB.Recordset
   
   D.SALE_ORDER_ID = -1
   D.INVENTORY_WH_DOC_ID = IdWh
   D.DOCUMENT_TYPE = 3
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(14, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSaleOrder
      Call TempData.PopulateFromRS(14, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
        If KeyType = 1 Then
         Call Cl.add(TempData)
       End If
      
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupDoItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DoID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.BILLING_DOC_SO_ID = DoID
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(47, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(47, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.BILLING_DOC_SO_ID)) & "-" & Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCredit(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CustomerCode As String = "", Optional FromCustomerCode As String = "", Optional ToCustomerCode As String = "")
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.CUSTOMER_CODE = CustomerCode
   D.FROM_CUSTOMER_CODE = FromCustomerCode
   D.TO_CUSTOMER_CODE = ToCustomerCode
   D.OrderType = 1
   D.DOCUMENT_TYPE = -1
   D.DocTypeSet = " (3, 4) "
   D.RECEIPT_TYPE = 3
   Call D.QueryData(10, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(10, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInventoryDocExt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CInventoryDocExt
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryDocExt
Dim I As Long

   Set D = New CInventoryDocExt
   Set Rs = New ADODB.Recordset
   
   D.INVENTORY_DOC_ID = -1
   D.DOCUMENT_NO = ""
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   Call D.QueryData(Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryDocExt
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.INVENTORY_DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReceiptWhenRecCheque(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocType
   Call D.QueryData(8, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(8, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.SUPPLIER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSaleBySetProductFeature(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional CustType As Long = -1, Optional DocTypeSet As String = "")
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.CUSTOMER_TYPE = CustType
   D.DocTypeSet = DocTypeSet
   Call D.QueryData(31, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(31, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.CUSTOMER_ID & "-" & TempData.SET_PRODUCT_ID & "-" & TempData.FEATURE_GROUP_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCashDocPostDistinctDocumentNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCashDocPost
Dim ItemCount As Long
Static Rs As ADODB.Recordset
Dim TempData As CCashDocPost
Dim I As Long

   Set D = New CCashDocPost
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   If Rs Is Nothing Then
      Set Rs = New ADODB.Recordset
      Call D.QueryData(3, Rs, ItemCount)
   Else
      Rs.MoveFirst
   End If
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCashDocPost
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.DOCUMENT_NO)
         C.ItemData(I) = TempData.CASH_DOC_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadJobInputByPartItemYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional ProcessID As Long, Optional PartGroupID As Long, Optional PartTypeID As Long, Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   D.PART_GROUP_ID = PartGroupID
   D.PART_TYPE_ID = PartTypeID
   D.LOCATION_ID = LocationID
   D.TX_TYPE = "E"
   Call D.QueryData(10, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJobInput
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadJobInputByDate(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional ProcessID As Long, Optional PartGroupID As Long, Optional PartTypeID As Long, Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CJobInput
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJobInput
Dim TempJob As CJob
Dim TempJobIn As CJobInput
Dim I As Long

   Set D = New CJobInput
   Set Rs = New ADODB.Recordset
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PROCESS_ID = ProcessID
   Call D.QueryData(12, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
While Not Rs.EOF
    I = I + 1
    Set TempData = New CJobInput
    Call TempData.PopulateFromRS(12, Rs)
    If TempData.TX_TYPE = "I" Then
      Set TempJob = GetObject("CJob", Cl, Trim(str(TempData.PART_ITEM_ID)), False)
      If TempJob Is Nothing Then
         Set TempJob = New CJob
         TempJob.PART_ITEM_ID = TempData.PART_ITEM_ID
         TempJob.PART_NO = TempData.PART_NO
         TempJob.PART_DESC = TempData.PART_DESC
'         TempJob.TX_TYPE = TempData.TX_TYPE
'         TempJob.PART_TYPE_ID = TempData.PART_TYPE_ID
'         TempJob.PART_TYPE_NAME = TempData.PART_TYPE_NAME
         Call Cl.add(TempJob, Trim(str(TempData.PART_ITEM_ID)))
      Else
         Set TempJob = Nothing
      End If
    Else
      If Not (TempJob Is Nothing) Then
         Set TempJobIn = GetObject("CJobinput", TempJob.Inputs, Trim(str(TempData.PART_ITEM_ID)), False)
         If TempJobIn Is Nothing Then
            Set TempJobIn = New CJobInput
            TempJobIn.PART_ITEM_ID = TempData.PART_ITEM_ID
            TempJobIn.TX_TYPE = TempData.TX_TYPE
            TempJobIn.PART_TYPE_ID = TempData.PART_TYPE_ID
            TempJobIn.PART_TYPE_NAME = TempData.PART_TYPE_NAME
            Call TempJob.Inputs.add(TempJobIn, Trim(str(TempData.PART_ITEM_ID)))
         End If
      End If
    End If
   
   Rs.MoveNext
   Wend
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadJobByInventoryWhDoc(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional KeyType As Long = 1, Optional InventoryWhDoc As Long = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long

   Set D = New CJob
   Set Rs = New ADODB.Recordset
   
   D.JOB_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.VERIFY_FLAG = "Y"
   D.INVENTORY_WH_DOC_ID = InventoryWhDoc
   If DocumentType = 13 Or DocumentType = 2001 Then
      D.PROCESS_ID_IN = "(4,7)"
   ElseIf DocumentType = 14 Or DocumentType = 2000 Then
      D.PROCESS_ID_IN = "(2,6)"
   End If
   Call D.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(7, Rs)
      
      If Not (Cl Is Nothing) Then
         If KeyType = 1 Then
            Call Cl.add(TempData, str(TempData.INVENTORY_WH_DOC_ID))
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAllBillingDoc(Cl As Collection, Optional ToDate As Date = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAllDocItem(Cl As Collection, Optional ToDate As Date = -1, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetAccountFollowCancelFlag_N(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim AF As CAccFol
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAccFol
Dim I As Long
   
   Set AF = New CAccFol
   Set Rs = New ADODB.Recordset
   
    AF.ACC_FOL_ID = -1
    AF.CANCEL_FLAG = "N"

   Call AF.QueryData2(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAccFol
      Call TempData.PopulateFromRS2(2, Rs)
      
      If Not (Cl Is Nothing) Then
         
         
         
         
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set AF = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPayByChequeDocPass(Cl As Collection, CustomerCode As String, FromCustomerCode As String, ToCustomerCode As String, CustType As Long, CustGrade As Long) ', CustomerID As Long
On Error GoTo ErrorHandler
Dim Cd As CChequeDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CChequeDoc
Dim TempData2 As CChequeDoc

Dim I As Long

   
   Set Cd = New CChequeDoc
   Set Rs = New ADODB.Recordset
   
'    AF.ACC_FOL_ID = -1
'    AF.CANCEL_FLAG = "N"
   Cd.CHEQUE_DOC_ID = -1
   Cd.CUSTOMER_CODE = CustomerCode
   Cd.FROM_CUSTOMER_CODE = FromCustomerCode
   Cd.TO_CUSTOMER_CODE = ToCustomerCode
   Cd.CUSTOMER_TYPE = CustType
   Cd.CUSTOMER_TYPE = CustGrade
    
   Call Cd.QueryData(4, Rs, ItemCount)
   

   If Not (Cl Is Nothing) Then
   Set Cl = Nothing
   Set Cl = New Collection
   End If

'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CChequeDoc
'      Call TempData.PopulateFromRS(4, Rs)
'
'
'      If Not (Cl Is Nothing) Then
'      Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO)
''       Call Cl.add(TempData, Trim(Str(TempData.CUSTOMER_ID)) & "-" & Trim(Str(TempData.RECEIPT_CHEQUE_DOC_NO)))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'    Set CD = Nothing
'   Set Rs = Nothing
'
'
'
'
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   While Not Rs.EOF
'      I = I + 1
      Set TempData2 = New CChequeDoc
      Call TempData2.PopulateFromRS(4, Rs)

'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.ACC_FOL_ID)))
'      End If

      Set TempData = GetObject("CChequeDoc", Cl, TempData2.CUSTOMER_ID & "-" & TempData2.RECEIPT_CHEQUE_DOC_NO & "-" & TempData2.RECEIPT_CHEQUE_DOC_ID, False)

      If TempData Is Nothing Then
         Set TempData = New CChequeDoc
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.CHEQUE_DOC_DATE = DateToStringExtEx2(TempData2.CHEQUE_DOC_DATE)
         TempData.PAID_AMOUNT = TempData2.PAID_AMOUNT
         TempData.CHEQUE_DOC_NO = TempData2.CHEQUE_DOC_NO
         TempData.RECEIPT_CHEQUE_DOC_NO = TempData2.RECEIPT_CHEQUE_DOC_NO
         TempData.RECEIPT_CHEQUE_DOC_ID = TempData2.RECEIPT_CHEQUE_DOC_ID
         TempData.PASSCHEQUE_FLAG = TempData2.PASSCHEQUE_FLAG
         TempData.BADCHEQUE_FLAG = TempData2.BADCHEQUE_FLAG
         TempData.PASSCHEQUE_DATE = TempData2.PASSCHEQUE_DATE
'         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO)
         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO & "-" & TempData.RECEIPT_CHEQUE_DOC_ID)
'      Else
'         TempData.FOL_NOTE = TempData.FOL_NOTE & vbCrLf & "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
      End If

      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set Cd = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadChequeDocWaitForBill(Cl As Collection, CustomerCode As String, FromCustomerCode As String, ToCustomerCode As String, CustType As Long, CustGrade As Long) ', CustomerID As Long
On Error GoTo ErrorHandler
Dim Cd As CChequeDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CChequeDoc
Dim TempData2 As CChequeDoc

Dim I As Long

   
   Set Cd = New CChequeDoc
   Set Rs = New ADODB.Recordset
   
'    AF.ACC_FOL_ID = -1
'    AF.CANCEL_FLAG = "N"
   Cd.CHEQUE_DOC_ID = -1
   Cd.CUSTOMER_CODE = CustomerCode
   Cd.FROM_CUSTOMER_CODE = FromCustomerCode
   Cd.TO_CUSTOMER_CODE = ToCustomerCode
   Cd.CUSTOMER_TYPE = CustType
   Cd.CUSTOMER_TYPE = CustGrade
    
   Call Cd.QueryData(9, Rs, ItemCount)
   

   If Not (Cl Is Nothing) Then
   Set Cl = Nothing
   Set Cl = New Collection
   End If

'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CChequeDoc
'      Call TempData.PopulateFromRS(4, Rs)
'
'
'      If Not (Cl Is Nothing) Then
'      Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO)
''       Call Cl.add(TempData, Trim(Str(TempData.CUSTOMER_ID)) & "-" & Trim(Str(TempData.RECEIPT_CHEQUE_DOC_NO)))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'    Set CD = Nothing
'   Set Rs = Nothing
'
'
'
'
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   While Not Rs.EOF
'      I = I + 1
      Set TempData2 = New CChequeDoc
      Call TempData2.PopulateFromRS(9, Rs)

'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.ACC_FOL_ID)))
'      End If

      Set TempData = GetObject("CChequeDoc", Cl, TempData2.CUSTOMER_ID & "-" & TempData2.RECEIPT_CHEQUE_DOC_NO & "-" & TempData2.RECEIPT_CHEQUE_DOC_ID, False)

      If TempData Is Nothing Then
         Set TempData = New CChequeDoc
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.CHEQUE_DOC_DATE = DateToStringExtEx2(TempData2.CHEQUE_DOC_DATE)
         TempData.PAID_AMOUNT = TempData2.PAID_AMOUNT
         TempData.CHEQUE_DOC_NO = TempData2.CHEQUE_DOC_NO
         TempData.RECEIPT_CHEQUE_DOC_NO = TempData2.RECEIPT_CHEQUE_DOC_NO
         TempData.RECEIPT_CHEQUE_DOC_ID = TempData2.RECEIPT_CHEQUE_DOC_ID
         TempData.PASSCHEQUE_FLAG = TempData2.PASSCHEQUE_FLAG
         TempData.BADCHEQUE_FLAG = TempData2.BADCHEQUE_FLAG
         TempData.PASSCHEQUE_DATE = TempData2.PASSCHEQUE_DATE
          TempData.BANK_NAME = TempData2.BANK_NAME
         TempData.BANK_BRANCH_NAME = TempData2.BANK_BRANCH_NAME
'         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO)
         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO & "-" & TempData.RECEIPT_CHEQUE_DOC_ID)
'      Else
'         TempData.FOL_NOTE = TempData.FOL_NOTE & vbCrLf & "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
      End If

      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set Cd = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadReceiptPayByChequeDoc(Cl As Collection, CustomerCode As String, FromCustomerCode As String, ToCustomerCode As String, CustType As Long, CustGrade As Long) ', CustomerID As Long
On Error GoTo ErrorHandler
Dim Cd As CChequeDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CChequeDoc
Dim TempData2 As CChequeDoc

Dim I As Long

   
   Set Cd = New CChequeDoc
   Set Rs = New ADODB.Recordset
   
'    AF.ACC_FOL_ID = -1
'    AF.CANCEL_FLAG = "N"
   Cd.CHEQUE_DOC_ID = -1
   Cd.CUSTOMER_CODE = CustomerCode
   Cd.FROM_CUSTOMER_CODE = FromCustomerCode
   Cd.TO_CUSTOMER_CODE = ToCustomerCode
   Cd.CUSTOMER_TYPE = CustType
   Cd.CUSTOMER_TYPE = CustGrade
    
    
   Call Cd.QueryData(6, Rs, ItemCount)
   

   If Not (Cl Is Nothing) Then
   Set Cl = Nothing
   Set Cl = New Collection
   End If

   While Not Rs.EOF

      Set TempData2 = New CChequeDoc
      Call TempData2.PopulateFromRS(6, Rs)
      Set TempData = GetObject("CChequeDoc", Cl, TempData2.CUSTOMER_ID & "-" & TempData2.DOCUMENT_NO & "-" & TempData2.CHEQUE_DOC_NO, False)
      If TempData Is Nothing Then
         Set TempData = New CChequeDoc
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.CHEQUE_DOC_DATE = DateToStringExtEx2(TempData2.CHEQUE_DOC_DATE)
         TempData.PAID_AMOUNT = TempData2.PAID_AMOUNT
         TempData.CHEQUE_DOC_NO = TempData2.CHEQUE_DOC_NO
         TempData.PASSCHEQUE_FLAG = TempData2.PASSCHEQUE_FLAG
         TempData.BADCHEQUE_FLAG = TempData2.BADCHEQUE_FLAG
         TempData.DOCUMENT_NO = TempData2.DOCUMENT_NO
         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.DOCUMENT_NO & "-" & TempData.CHEQUE_DOC_NO)
     End If

      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set Cd = Nothing
   
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadPayByChequeDocNotPass(Cl As Collection, CustomerCode As String, FromCustomerCode As String, ToCustomerCode As String, CustType As Long, CustGrade As Long)
On Error GoTo ErrorHandler
Dim Cd As CChequeDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CChequeDoc
Dim TempData2 As CChequeDoc

Dim I As Long

   
   Set Cd = New CChequeDoc
   Set Rs = New ADODB.Recordset
   
'    AF.ACC_FOL_ID = -1
'    AF.CANCEL_FLAG = "N"
   Cd.CHEQUE_DOC_ID = -1
   Cd.CUSTOMER_CODE = CustomerCode
   Cd.FROM_CUSTOMER_CODE = FromCustomerCode
   Cd.TO_CUSTOMER_CODE = ToCustomerCode
   Cd.CUSTOMER_TYPE = CustType
   Cd.CUSTOMER_TYPE = CustGrade
    
   Call Cd.QueryData(5, Rs, ItemCount)
   

   If Not (Cl Is Nothing) Then
   Set Cl = Nothing
   Set Cl = New Collection
   End If

'   While Not Rs.EOF
'      I = I + 1
'      Set TempData = New CChequeDoc
'      Call TempData.PopulateFromRS(4, Rs)
'
'
'      If Not (Cl Is Nothing) Then
'      Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO)
''       Call Cl.add(TempData, Trim(Str(TempData.CUSTOMER_ID)) & "-" & Trim(Str(TempData.RECEIPT_CHEQUE_DOC_NO)))
'      End If
'
'      Set TempData = Nothing
'      Rs.MoveNext
'   Wend
'    Set CD = Nothing
'   Set Rs = Nothing
'
'
'
'
'
'   If Not (Cl Is Nothing) Then
'      Set Cl = Nothing
'      Set Cl = New Collection
'   End If
   While Not Rs.EOF
'      I = I + 1
      Set TempData2 = New CChequeDoc
      Call TempData2.PopulateFromRS(5, Rs)

'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.ACC_FOL_ID)))
'      End If

      Set TempData = GetObject("CChequeDoc", Cl, TempData2.CUSTOMER_ID & "-" & TempData2.RECEIPT_CHEQUE_DOC_NO & "-" & TempData2.RECEIPT_CHEQUE_DOC_ID, False)

      If TempData Is Nothing Then
         Set TempData = New CChequeDoc
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.CHEQUE_DOC_DATE = DateToStringExtEx2(TempData2.CHEQUE_DOC_DATE)
         TempData.PAID_AMOUNT = TempData2.PAID_AMOUNT
         TempData.CHEQUE_DOC_NO = TempData2.CHEQUE_DOC_NO
         TempData.RECEIPT_CHEQUE_DOC_NO = TempData2.RECEIPT_CHEQUE_DOC_NO
         TempData.RECEIPT_CHEQUE_DOC_ID = TempData2.RECEIPT_CHEQUE_DOC_ID
         TempData.BADCHEQUE_DATE = TempData2.BADCHEQUE_DATE
'         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO)
         Call Cl.add(TempData, TempData.CUSTOMER_ID & "-" & TempData.RECEIPT_CHEQUE_DOC_NO & "-" & TempData.RECEIPT_CHEQUE_DOC_ID)
'      Else
'         TempData.FOL_NOTE = TempData.FOL_NOTE & vbCrLf & "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
      End If

      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set Cd = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetMKTFollow(Cl As Collection)
On Error GoTo ErrorHandler
Dim MKT As CMKTFol
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMKTFol
Dim TempData2 As CMKTFol
Dim I As Long
   
   Set MKT = New CMKTFol
   Set Rs = New ADODB.Recordset
   
    MKT.MKT_FOL_ID = -1
    MKT.CANCEL_FLAG = "N"
    
   Call MKT.QueryData2(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
'      I = I + 1
      Set TempData2 = New CMKTFol
      Call TempData2.PopulateFromRS2(1, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.ACC_FOL_ID)))
'      End If
      
      Set TempData = GetObject("CMKTFol", Cl, Trim(str(TempData2.CUSTOMER_ID)), False)
      If TempData Is Nothing Then
         Set TempData = New CMKTFol
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.FOL_NOTE = "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      Else
         TempData.FOL_NOTE = TempData.FOL_NOTE & vbCrLf & "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
      End If
         
      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set MKT = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetAccountFollow(Cl As Collection)
On Error GoTo ErrorHandler
Dim AF As CAccFol
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAccFol
Dim TempData2 As CAccFol
Dim I As Long
   
   Set AF = New CAccFol
   Set Rs = New ADODB.Recordset
   
    AF.ACC_FOL_ID = -1
    AF.CANCEL_FLAG = "N"
    
   Call AF.QueryData2(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
'      I = I + 1
      Set TempData2 = New CAccFol
      Call TempData2.PopulateFromRS2(1, Rs)
      
'      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.ACC_FOL_ID)))
'      End If
      
      Set TempData = GetObject("CAccFol", Cl, Trim(str(TempData2.CUSTOMER_ID)), False)
      If TempData Is Nothing Then
         Set TempData = New CAccFol
         TempData.CUSTOMER_ID = TempData2.CUSTOMER_ID
         TempData.FOL_NOTE = "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      Else
         TempData.FOL_NOTE = TempData.FOL_NOTE & vbCrLf & "- " & DateToStringExtEx2(TempData2.FOL_DATE) & " " & TempData2.FOL_NOTE
      End If
         
      Set TempData = Nothing
      Set TempData2 = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set AF = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetMarketingFollowCancelFlag_N(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim MF As CMKTFol
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMKTFol
Dim I As Long
   
   Set MF = New CMKTFol
   Set Rs = New ADODB.Recordset
   
    MF.MKT_FOL_ID = -1
    MF.CANCEL_FLAG = "N"
    MF.FROM_DATE = FromDate
    MF.TO_DATE = ToDate
   MF.DOCUMENT_TYPE = DocumentType

   Call MF.QueryData2(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMKTFol
      Call TempData.PopulateFromRS2(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set MF = Nothing
   
   
'   If Rs.State = adStateOpen Then
'      Rs.Close
'   End If
'   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadMonthlyBalance(C As ComboBox, Optional Cl As Collection = Nothing, Optional YYYYMM As String, Optional LocationID As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMonthlyAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMonthlyAccum
Dim I As Long

   Set D = New CMonthlyAccum
   Set Rs = New ADODB.Recordset

   D.YYYYMM = YYYYMM
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(2, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMonthlyAccum
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMonthlyBalancePartItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional YYYYMM As String, Optional LocationID As Long = -1, Optional PartItemID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CMonthlyAccum
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMonthlyAccum
Dim I As Long

   Set D = New CMonthlyAccum
   Set Rs = New ADODB.Recordset
   
   D.YYYYMM = YYYYMM
   D.PART_ITEM_ID = PartItemID
   D.LOCATION_ID = LocationID
   D.OrderBy = 1
   Call D.QueryData(5, Rs, ItemCount, False)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMonthlyAccum
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctDoIDFromReceipt(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   Call D.QueryData(107, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(107, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function FindFactoryAddress() As String
Dim m_Rs1 As ADODB.Recordset
Dim m_Rs2 As ADODB.Recordset
Dim I As Long
Dim iCount  As Long
   
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CEnterpriseAddress
      Set cEnpAddr = New CEnterpriseAddress
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
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
            
            FindFactoryAddress = cAddr.PackAddress
            
            Set cAddr = Nothing
                  
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
End Function

Public Sub InitPlanPartOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ และ รหัสวัตถุดิบ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("รหัสวัตถุดิบ และ วันที่"))
   C.ItemData(2) = 2
End Sub
Public Sub LoadPlanDateAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional Area As Long = -1)
On Error GoTo ErrorHandler
Dim D As CPlanPart
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPlanPart
Dim I As Long
   
   Set D = New CPlanning
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PLAN_AREA = Area
   D.CANCEL_FLAG = "N"
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPlanPart
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, TempData.PLAN_AREA & "-" & TempData.PART_ITEM_ID & "-" & TempData.PLAN_DATE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function LoadPlanVersion(FromDate As Date, ToDate As Date, Optional PlanArea As Integer) As Long
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Pn As CPlanning
Dim TempData As CPlanning
Dim I As Long
   
   Set Rs = New ADODB.Recordset
   Set TempData = New CPlanning
   Set Pn = New CPlanning
   
   Pn.FROM_DATE = FromDate
   Pn.TO_DATE = ToDate
   Pn.PLANNING_AREA = PlanArea
   Call Pn.QueryData(2, Rs, ItemCount)
   If Not Pn Is Nothing Then
      While Not Rs.EOF
        Call TempData.PopulateFromRS(2, Rs)
        Rs.MoveNext
      Wend
      LoadPlanVersion = TempData.PLAN_VERSION
   Else
      LoadPlanVersion = 0
   End If
   Set Rs = Nothing
   Set Pn = Nothing
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function
Public Sub LoadPlanningItemDateAmount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional PlanningArea As Long = -1, Optional PlanningSubType As Long = -1, Optional FuncName As String)
On Error GoTo ErrorHandler
Dim D As CPlanningItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPlanningItem
Dim SearchItemNo As CPlanningItem
Dim I As Long
Dim DiffDate As Long
Dim StrSQLDate As String
Dim StrSQLDate2 As String
Dim TempDate As Date

   Set D = New CPlanningItem
   Set Rs = New ADODB.Recordset
      
      'สร้าง sql ในนี้
      DiffDate = DateDiff("D", FromDate, ToDate)
      TempDate = FromDate
      StrSQLDate2 = "((PNI.PLANNING_SUB_TYPE = " & PlanningSubType & ") AND  (PN.PLANNING_AREA = " & PlanningArea & ") AND "
      StrSQLDate = StrSQLDate2 & "(PN.PLANNING_DATE ='" & DateToStringIntLow(TempDate) & "') AND (PN.PLAN_VERSION = " & LoadPlanVersion(TempDate, TempDate, Trim(str(PlanningArea))) & ")) "
      For I = 1 To DiffDate
         TempDate = DateAdd("D", 1, TempDate)
         StrSQLDate = StrSQLDate & "OR" & StrSQLDate2 & " (PN.PLANNING_DATE ='" & DateToStringIntLow(TempDate) & "') AND (PN.PLAN_VERSION = " & LoadPlanVersion(TempDate, TempDate, Trim(str(PlanningArea))) & ")) "
      Next I

   D.STR_SQL_DATE = StrSQLDate
   Call D.QueryData(2, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPlanningItem
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (Cl Is Nothing) Then
      
         Set SearchItemNo = GetObject("CPlanningItem", Cl, Trim(TempData.PART_ITEM_ID & "-" & TempData.PLANNING_DATE), False)
         If SearchItemNo Is Nothing Then
            Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PLANNING_DATE)
         ElseIf TempData.PLANNING_ITEM_ID > SearchItemNo.PLANNING_ITEM_ID Then
            Call Cl.Remove(SearchItemNo.PART_ITEM_ID & "-" & SearchItemNo.PLANNING_DATE)
            Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PLANNING_DATE)
         End If
      
         'Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.PLANNING_DATE)
      End If
   
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & " <Func: " & FuncName & "> " & TempData.PART_ITEM_ID & "-" & TempData.PLANNING_DATE
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadDistinctJobNo(C As ComboBox, Optional Cl As Collection = Nothing, Optional ProcessID As Long = 0)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CJob
   Set Rs = New ADODB.Recordset
   D.PROCESS_ID = ProcessID
   Call D.QueryData(5, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(5, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctJobNoToLot(C As ComboBox, Optional Cl As Collection = Nothing, Optional ProcessID As Long = 0, Optional JobNo As String)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CJob
   Set Rs = New ADODB.Recordset
   D.JOB_NO_LIKE = JobNo
   D.PROCESS_ID = ProcessID
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDistinctJobNoToLotGenerateBatch(C As ComboBox, Optional Cl As Collection = Nothing, Optional ProcessID As Long = 0, Optional JobNo As String)
On Error GoTo ErrorHandler
Dim D As CJob
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CJob
Dim I As Long
Dim FromDate1 As String
Dim ToDate1 As String
Static TempCol As Collection
   
   Set D = New CJob
   Set Rs = New ADODB.Recordset
   D.JOB_NO_LIKE = JobNo
   D.PROCESS_ID = ProcessID
   Call D.QueryData(6, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CJob
      Call TempData.PopulateFromRS(6, Rs)
   
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.JOB_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Set TempCol = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPaidAmountByCustomer2(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional CustomerID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.CUSTOMER_ID = CustomerID
   Call D.QueryData(115, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(115, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDnCnAmountByCustomer2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1, Optional SourceType As Long = 1, Optional CustomerID As Long)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   If SourceType = 1 Then
      D.FROM_ITEM_DATE = FromDate
      D.TO_ITEM_DATE = ToDate
   ElseIf SourceType = 2 Then
      D.FROM_DOC_DATE = FromDate
      D.TO_DOC_DATE = ToDate
   End If
   D.CUSTOMER_ID = CustomerID
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(116, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(116, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.DO_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitAlertBoxType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (AlertBoxType2Text(1))
   C.ItemData(1) = 1

   C.AddItem (AlertBoxType2Text(2))
   C.ItemData(2) = 2
End Sub

Public Function AlertBoxType2Text(Ind As Long) As String
   If Ind = 1 Then
      AlertBoxType2Text = "ข้อความแจ้งเตือน"
   ElseIf Ind = 2 Then
      AlertBoxType2Text = "แจ้งเตือนออกใบรับของโดยไม่มี PO"
   End If
End Function
Public Function LogonStatus2Text(Ind As Long) As String
   If Ind = 1 Then
      LogonStatus2Text = "LOG ON"
   ElseIf Ind = 2 Then
      LogonStatus2Text = "LOG OFF"
   Else
      LogonStatus2Text = "LOG OFF"
   End If
End Function
Public Sub InitLogonStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (LogonStatus2Text(1))
   C.ItemData(1) = 1
   
   C.AddItem (LogonStatus2Text(2))
   C.ItemData(2) = 2
End Sub
Public Sub LoadSumJobInputViaLotItem(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FuncName As String)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
Dim Li As CLotItem

   Set Li = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Li.LOT_ITEM_ID = -1
   Li.TX_TYPE = "E"
   Li.FROM_DATE = FromDate
   Li.TO_DATE = ToDate
   
   Call Li.QueryData(34, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(34, Rs)
   
      
      If Not (Cl Is Nothing) Then
        Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & " <Func: " & FuncName & ">"
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadInventoryActual(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional FuncName As String)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CInventoryActItem
Dim I As Long
Dim Iai As CInventoryActItem
Dim TempData2 As CInventoryActItem
Dim isError As Long
isError = 1
   Set Iai = New CInventoryActItem
   Set Rs = New ADODB.Recordset
   
   Iai.INVENTORY_ACT_ITEM_ID = -1
   Iai.INVENTORY_ACT_DATE = -1
   Iai.FROM_DATE = FromDate
   Iai.TO_DATE = ToDate
   
   Call Iai.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CInventoryActItem
      Call TempData.PopulateFromRS(1, Rs)
   
'   If TempData.PART_NO = 8032 Then
'     Debug.Print
'   End If
      
      If Not (Cl Is Nothing) Then
         isError = 2
         Set TempData2 = GetObject("CInventoryActItem", Cl, TempData.PART_ITEM_ID & "-" & TempData.INVENTORY_ACT_DATE, False)
         If TempData2 Is Nothing Then
             Call Cl.add(TempData, TempData.PART_ITEM_ID & "-" & TempData.INVENTORY_ACT_DATE)
         End If
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
    If isError = 2 Then
      glbErrorLog.SystemErrorMsg = Err.DESCRIPTION & " <Func: " & FuncName & ">" & "วัตถุดิบเบอร์ " & TempData.PART_NO & " ซ้ำกันในวันที่ " & TempData.INVENTORY_ACT_DATE
      Set TempData = Nothing
    Else
      glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   End If
    glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAuthenPO_Verify(Optional Cl As Collection = Nothing, Optional DOCUMENT_TYPE As String, Optional TOTAL_PRICE As Double)
On Error GoTo ErrorHandler
Dim D As CAuthenPO
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAuthenPO
Dim I As Long
   Set D = New CAuthenPO
   Set Rs = New ADODB.Recordset
   
   D.AUTHEN_PO_ID = -1
   D.AUTHEN_PO_GROUP = DOCUMENT_TYPE
   D.TOTAL_PRICE = TOTAL_PRICE
   D.AUTHEN_AREA = 1
   D.OrderBy = -1
   D.OrderType = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   Set D = Nothing
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAuthenPO
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadAuthenPO_Approve(Optional Cl As Collection = Nothing, Optional DOCUMENT_TYPE As String, Optional TOTAL_PRICE As Double)
On Error GoTo ErrorHandler
Dim D As CAuthenPO
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAuthenPO
Dim I As Long
   Set D = New CAuthenPO
   Set Rs = New ADODB.Recordset
   
   D.AUTHEN_PO_ID = -1
   D.AUTHEN_PO_GROUP = DOCUMENT_TYPE
   D.TOTAL_PRICE = TOTAL_PRICE
   D.AUTHEN_AREA = 2
   D.OrderBy = -1
   D.OrderType = 1
   Call D.QueryData(2, Rs, ItemCount)
   
   Set D = Nothing
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAuthenPO
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLoginTracking(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CLoginTracking
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLoginTracking
Dim I As Long

   Set D = New CLoginTracking
   Set Rs = New ADODB.Recordset
   
   D.LOGIN_TRACKING_ID = -1
   D.LOGIN_FROM_DATE = FromDate
   D.LOGIN_TO_DATE = ToDate
   D.OrderBy = 1
   D.OrderType = 1
   Call D.QueryData(Rs, ItemCount)
   
   Set D = Nothing
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLoginTracking
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Val(TempData.LOGIN_TRACKING_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub
Public Sub LoadTargetEmp(Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CTargetDetail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CTargetDetail
Dim I As Long
   
   Set D = New CTargetDetail
   Set Rs = New ADODB.Recordset
   
   D.TARGET_DETAIL_ID = -1
   Call D.QueryData(2, Rs, ItemCount)
   
   Set D = Nothing
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CTargetDetail
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.EMP_ID & "-" & TempData.YEAR_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitCommissionOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เริ่มใช้"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("วันที่สิ้นสุด"))
   C.ItemData(3) = 3

End Sub

Public Sub LoadCommissionBudgetChart(C As ComboBox, Optional Cl As Collection = Nothing, Optional FK_ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionBgChart
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionBgChart

Dim I As Long
   If FK_ID <= 0 Then
      Exit Sub
   End If
   Set D = New CCommissionBgChart
   Set Rs = New ADODB.Recordset
   
   D.COMMISSION_BUDGET_CHART_ID = -1
   D.MASTER_VALID_ID = FK_ID
   D.OrderType = 1
   Call D.QueryData(1, Rs, ItemCount)
      
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionBgChart
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.COMMISSION_BUDGET_CHART_ID & " - " & TempData.EMP_NAME & " " & TempData.EMP_LNAME)
         C.ItemData(I) = TempData.COMMISSION_BUDGET_CHART_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.OLD_PK)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub InitCommissionSaleType(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("%")
   C.ItemData(1) = 1
   
   C.AddItem ("บาท")
   C.ItemData(2) = 2
   
End Sub
Public Function GetCommissionSaleTypeName(Ind As Long) As String
   If Ind = 1 Then
      GetCommissionSaleTypeName = "%"
   ElseIf Ind = 2 Then
      GetCommissionSaleTypeName = "บาท"
   End If
End Function
Public Sub LoadCommissionChartValidFromTo(Cl As Collection, Optional FromValidDate As Date = -1, Optional ToValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionBgChart
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionBgChart
Dim I As Long
Dim TempStock As CCommissionBgChart

   Set D = New CCommissionBgChart
   Set Rs = New ADODB.Recordset
   D.VALID_FROM = FromValidDate
   D.VALID_TO = ToValidDate
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionBgChart
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.COMMISSION_BUDGET_CHART_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCommissionChartValidFromToEmp(Cl As Collection, Optional FromValidDate As Date = -1, Optional ToValidDate As Date = -1)
On Error GoTo ErrorHandler
Dim D As CCommissionBgChart
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionBgChart
Dim I As Long
Dim TempStock As CCommissionBgChart

   Set D = New CCommissionBgChart
   Set Rs = New ADODB.Recordset
   D.VALID_FROM = FromValidDate
   D.VALID_TO = ToValidDate
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionBgChart
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.EMP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadDnCnAmountBySale(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional DocumentType As Long = -1)
On Error GoTo ErrorHandler
Dim D As CReceiptItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CReceiptItem
Dim I As Long

   Set D = New CReceiptItem
   Set Rs = New ADODB.Recordset
   
   D.RECEIPT_ITEM_ID = -1
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.DOCUMENT_TYPE = DocumentType
   Call D.QueryData(117, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CReceiptItem
      Call TempData.PopulateFromRS(117, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.ACCEPT_BY & "-" & TempData.YYYYMM))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMasterValidCommissionSale(Cl As Collection)
On Error GoTo ErrorHandler
Dim D As CCommissionSale
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionSale
Dim I As Long
Dim m_Valid As CMasterValid
   
   Set D = New CCommissionSale
   Set Rs = New ADODB.Recordset
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionSale
      Call TempData.PopulateFromRS(2, Rs)
      
      Set m_Valid = GetObject("CMasterValid", Cl, Trim(str(TempData.MASTER_VALID_ID)), False)
      If m_Valid Is Nothing Then
         Set m_Valid = New CMasterValid
         m_Valid.VALID_FROM = TempData.VALID_FROM
         m_Valid.VALID_TO = TempData.VALID_TO
         Call Cl.add(m_Valid, Trim(str(TempData.MASTER_VALID_ID)))
      End If
      
      If TempData.COMMISSION_SALE_AREA = 1 Then
         Call m_Valid.CollSaleRcp.add(TempData)
      ElseIf TempData.COMMISSION_SALE_AREA = 2 Then
         Call m_Valid.CollSaleNow.add(TempData)
      ElseIf TempData.COMMISSION_SALE_AREA = 3 Then
         Call m_Valid.CollSaleManagerNow.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMasterValidCommissionCost(Cl As Collection, FromDate2 As Date, ToDate2 As Date)
On Error GoTo ErrorHandler
Dim D As CCommissionCost
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionCost
Dim I As Long
Dim m_Valid As CMasterValid
   
   Set D = New CCommissionCost
   Set Rs = New ADODB.Recordset
   D.VALID_FROM2 = FromDate2
   D.VALID_TO2 = ToDate2
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionCost
      Call TempData.PopulateFromRS(2, Rs)
      
      Set m_Valid = GetObject("CMasterValid", Cl, Trim(str(TempData.MASTER_VALID_ID)), False)
      If m_Valid Is Nothing Then
         Set m_Valid = New CMasterValid
         m_Valid.VALID_FROM = TempData.VALID_FROM
         m_Valid.VALID_TO = TempData.VALID_TO
         Call Cl.add(m_Valid, Trim(str(TempData.MASTER_VALID_ID)))
      End If
      
      Call m_Valid.CollCommissionCost.add(TempData, Trim(str(TempData.PART_ITEM_ID)))
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCommissionSubtractAmount(C As ComboBox, Optional Cl As Collection = Nothing, Optional YearNo As Long, Optional MonthID As Long)
On Error GoTo ErrorHandler
Dim D As CCommissionSubtract
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionSubtract
Dim I As Long
Dim mainSub  As CCommissionSubtract

   Set D = New CCommissionSubtract
   Set Rs = New ADODB.Recordset
   
   D.COMMISSION_SUBTRACT_ID = -1
   D.MONTH_ID = MonthID
   D.YEAR_NO = YearNo
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionSubtract
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
      End If
            
      Set mainSub = GetObject("CCommissionSubtract", Cl, Trim(str(TempData.EMP_ID)), False)
      If mainSub Is Nothing Then
         Set mainSub = New CCommissionSubtract
         mainSub.EMP_ID = TempData.EMP_ID
         Call Cl.add(mainSub, Trim(str(TempData.EMP_ID)))
      End If
      
      Call mainSub.collCommissionSubTractSub.add(TempData)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCommissionCredit(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CCommissionCredit
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionCredit
Dim I As Long

   Set D = New CCommissionCredit
   Set Rs = New ADODB.Recordset
   
   D.COMMISSION_CREDIT_ID = -1
   Call D.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionCredit
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(str(TempData.CUSTOMER_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
'
Public Sub LoadCommissionIncentive(C As ComboBox, Optional Cl As Collection = Nothing, Optional DocumentType As Long = -1, Optional FREELANCE_CODE As String = "", Optional FROM_FREELANCE_CODE As String = "", Optional TO_FREELANCE_CODE As String = "")
On Error GoTo ErrorHandler
Dim D As CCommissionIncentive
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCommissionIncentive
Dim I As Long

   Set D = New CCommissionIncentive
   Set Rs = New ADODB.Recordset
   
   D.INCENTIVE_ID = -1
   D.DOCUMENT_TYPE = DocumentType
   D.FREELANCE_CODE = FREELANCE_CODE
   D.FROM_FREELANCE_CODE = FROM_FREELANCE_CODE
   D.TO_FREELANCE_CODE = TO_FREELANCE_CODE
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCommissionIncentive
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         If DocumentType = 1 Then
            Call Cl.add(TempData, Trim(TempData.FREELANCE_ID & "-" & TempData.PART_ITEM_ID))
         ElseIf DocumentType = 2 Then
            Call Cl.add(TempData, Trim(TempData.FREELANCE_ID & "-" & TempData.CUSTOMER_ID & "-" & TempData.PART_ITEM_ID))
         ElseIf DocumentType = 3 Or DocumentType = 4 Then
            Call Cl.add(TempData)
         End If
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub CalSumSale(C As ComboBox, ByRef Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional EmpCode As String)
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim Cm As CEmployee
Dim Di As CDoItem
Dim TempDI As CDoItem
Dim Rs As ADODB.Recordset

   Set Rs = New ADODB.Recordset
'   Call LoadEmployee(Nothing, m_Employees)

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
   Next J

      I = 0

      Set Di = New CDoItem
      Di.DO_ITEM_ID = -1
      Di.EMP_CODE = EmpCode
      Di.FROM_DATE = FromDate
      Di.TO_DATE = ToDate
      Di.DocTypeSet = BillingDocType2Set(0)
      Di.OrderBy = 1
      Di.OrderType = 1
      Call Di.QueryData(4, Rs, iCount)

      I = 0
      PrevKey1 = ""
      If Not Rs.EOF Then
         Call Di.PopulateFromRS(4, Rs)
         PrevKey1 = Di.ACCEPT_BY
      End If

      While Not Rs.EOF
         Call Di.PopulateFromRS(4, Rs)
         If Di.PARCEL_TYPE = 2 Then
            Di.PACK_AMOUNT = 0
         End If

         If Di.PART_ITEM_ID > 0 Then
            If PrevKey1 <> Di.ACCEPT_BY Then
               Set TempDI = New CDoItem
               TempDI.PACK_AMOUNT = Total1(6)
               TempDI.ITEM_AMOUNT = Total1(7)
               Call Cl.add(TempDI, Trim(EmpCode))

               For J = 1 To UBound(Total1)
                  Total1(J) = 0
               Next J
            End If
            PrevKey1 = Di.ACCEPT_BY

            Total1(6) = Total1(6) + Di.PACK_AMOUNT
            Total1(7) = Total1(7) + Di.ITEM_AMOUNT
'            Total1(8) = Total1(8) + (Di.BEFORE_DISCOUNT_PRICE)
'            Total1(9) = Total1(9) + (Di.DISCOUNT_AMOUNT)
'            Total1(10) = Total1(10) + (Di.EXTRA_DISCOUNT)
'            Total1(11) = Total1(11) + (Di.TOTAL_PRICE - Di.EXTRA_DISCOUNT)
         End If

         Rs.MoveNext
      Wend

   Set TempDI = New CDoItem
   TempDI.PACK_AMOUNT = Total1(6)
   TempDI.ITEM_AMOUNT = Total1(7)
   Call Cl.add(TempDI, Trim(EmpCode))
               
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
ErrHandler:
   Set Rs = Nothing
End Sub
Public Sub LoadTotalPriceBySaleYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = 1
   Call D.QueryData(42, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(42, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.ACCEPT_BY & "-" & TempData.PART_ITEM_ID & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadTotalPriceByPartYYYYMM(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim D As CDoItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDoItem
Dim I As Long

   Set D = New CDoItem
   Set Rs = New ADODB.Recordset
   
   D.DO_ITEM_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE = 1
   Call D.QueryData(52, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDoItem
      Call TempData.PopulateFromRS(52, Rs)
   
      If Not (C Is Nothing) Then
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.YYYYMM))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GeneratePartGroupMenu(Col As Collection)
Dim G As CPartGroup
Dim D As CMenuItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   Set G = New CPartGroup
   Set TempRs = New ADODB.Recordset
   
   G.PART_GROUP_ID = -1
   Call G.QueryData(TempRs, iCount)
   
   While Not TempRs.EOF
      Call G.PopulateFromRS(1, TempRs)
      
      Set D = New CMenuItem
      D.KEYWORD = G.PART_GROUP_NAME
      D.KEY_ID = G.PART_GROUP_ID
      Call Col.add(D)
      Set D = Nothing
      
      TempRs.MoveNext
   Wend
      
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   Set G = Nothing
End Sub

Public Sub GenerateJobProcessMenu(Col As Collection)
Dim G As CProcess
Dim D As CMenuItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   Set G = New CProcess
   Set TempRs = New ADODB.Recordset
   
   G.PROCESS_ID = -1
   Call G.QueryData(TempRs, iCount)
   
   While Not TempRs.EOF
      Call G.PopulateFromRS(1, TempRs)
      
      Set D = New CMenuItem
      D.KEYWORD = G.PROCESS_NAME
      D.KEY_ID = G.PROCESS_ID
      Call Col.add(D)
      Set D = Nothing
      
      TempRs.MoveNext
   Wend
      
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   Set G = Nothing
End Sub
