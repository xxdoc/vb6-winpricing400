VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuiConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GuiConfigs As Collection
Public SellMenuItems As Collection
Public Tax2Items As Collection
Public BuyMenuItems As Collection
Public DOAddMenuItems As Collection
Public POAddMenuItems As Collection
Public SupAddMenuItems As Collection
Public SupAddNoPoMenuItems As Collection
Public SupItemPoOtherMenuItems As Collection
Public ReceiptAddMenuItems As Collection
Public ReportMenuItems As Collection
Public QuoatationPrintMenuItems As Collection
Public POPrintMenuItems As Collection
Public SOPrintMenuItems As Collection
Public DOPrintMenuItems As Collection
Public IVPrintMenuItems As Collection
Public RCPrintMenuItems As Collection
Public RCPrintMenuItemsSpacial As Collection
Public RCPrintMenuItemsSpacialBuy As Collection
Public CdnDbnPrintMenuItems As Collection
Public BSPrintMenuItems As Collection
Public PkglistPrintMenuItems As Collection

Public LoadGoodsAddMenuItems As Collection
Public LoadAddJopProcessItems As Collection



Public MasterCommandMenuItems As Collection
Public MainCommandMenuItems As Collection
Public StockCommandMenuItems As Collection
Public StockWHCommandMenuItems As Collection
'GenerateStockWHCommandMenuItems
Public ProdCommandMenuItems As Collection
Public LedgerCommandMenuItems As Collection
Public PackageCommandMenuItems As Collection
Public AdminCommandMenuItems As Collection
Public PlanCommandMenuItems As Collection

Public CommissionCommandMenuItems As Collection

Public PictureItems As Collection
Private Sub Class_Initialize()
   Set m_GuiConfigs = New Collection
   Set SellMenuItems = New Collection
   Set Tax2Items = New Collection
   Set BuyMenuItems = New Collection
   
   Set ReportMenuItems = New Collection
   Set DOAddMenuItems = New Collection
   Set POAddMenuItems = New Collection
   Set SupAddMenuItems = New Collection
   Set SupAddNoPoMenuItems = New Collection
   Set SupItemPoOtherMenuItems = New Collection
   Set SOPrintMenuItems = New Collection
   Set ReceiptAddMenuItems = New Collection
   Set QuoatationPrintMenuItems = New Collection
   Set POPrintMenuItems = New Collection
   Set DOPrintMenuItems = New Collection
   Set IVPrintMenuItems = New Collection
   Set RCPrintMenuItems = New Collection
   Set RCPrintMenuItemsSpacial = New Collection
   Set RCPrintMenuItemsSpacialBuy = New Collection
   Set CdnDbnPrintMenuItems = New Collection
   Set BSPrintMenuItems = New Collection
   Set PkglistPrintMenuItems = New Collection
   
   Set LoadGoodsAddMenuItems = New Collection
   Set LoadAddJopProcessItems = New Collection
   
   Set MasterCommandMenuItems = New Collection
   Set MainCommandMenuItems = New Collection
   Set StockCommandMenuItems = New Collection
   Set StockWHCommandMenuItems = New Collection
   Set ProdCommandMenuItems = New Collection
   Set LedgerCommandMenuItems = New Collection
   Set PackageCommandMenuItems = New Collection
   Set AdminCommandMenuItems = New Collection
   Set PlanCommandMenuItems = New Collection
   Set CommissionCommandMenuItems = New Collection
   
   Set PictureItems = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_GuiConfigs = Nothing
   Set SellMenuItems = Nothing
   Set Tax2Items = Nothing
   Set BuyMenuItems = Nothing
   
   Set ReportMenuItems = Nothing
   Set DOAddMenuItems = Nothing
   Set POAddMenuItems = Nothing
   Set SupAddMenuItems = Nothing
   Set SupAddNoPoMenuItems = Nothing
   Set SupItemPoOtherMenuItems = Nothing
   Set ReceiptAddMenuItems = Nothing
   Set QuoatationPrintMenuItems = Nothing
   Set POPrintMenuItems = Nothing
   Set SOPrintMenuItems = Nothing
   Set DOPrintMenuItems = Nothing
   Set IVPrintMenuItems = Nothing
   Set RCPrintMenuItems = Nothing
   Set CdnDbnPrintMenuItems = Nothing
   Set BSPrintMenuItems = Nothing
   
   Set MasterCommandMenuItems = Nothing
   Set MainCommandMenuItems = Nothing
   Set StockCommandMenuItems = Nothing
   Set StockWHCommandMenuItems = Nothing
   Set ProdCommandMenuItems = Nothing
   Set LedgerCommandMenuItems = Nothing
   Set PackageCommandMenuItems = Nothing
   Set AdminCommandMenuItems = Nothing
   Set PlanCommandMenuItems = Nothing
   Set RCPrintMenuItemsSpacial = Nothing
   Set RCPrintMenuItemsSpacialBuy = Nothing
   Set PkglistPrintMenuItems = Nothing
   Set PictureItems = Nothing
   Set CommissionCommandMenuItems = Nothing
   
   Set LoadGoodsAddMenuItems = Nothing
   Set LoadAddJopProcessItems = Nothing
End Sub

Private Sub CreateGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = New CGuiConfig
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
   
   Call m_GuiConfigs.add(TempC, Key)
   Set TempC = Nothing
End Sub

Private Sub SetGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(Key)
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
End Sub

Private Sub GenerateDOAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการใหม่ (ทีละรายการ)")
   TempID(1) = 1
   TempKey(2) = MapText("เพิ่มรายการใหม่ (เป็นชุด)")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("เลือกจากใบ PO")
   TempID(4) = 4
   TempKey(5) = MapText("เลือกจากใบเสนอราคา")
   TempID(5) = 5
   TempKey(6) = MapText("เลือกจาก SO")
   TempID(6) = 6
   TempKey(7) = MapText("เลือกจากใบขึ้นอาหาร BAG")
   TempID(7) = 7
   TempKey(8) = MapText("-")
   TempID(8) = 8
   TempKey(9) = MapText("เลือกจากใบขึ้นอาหาร BULK")
   TempID(9) = 9
   TempKey(10) = MapText("-")
   TempID(10) = 10
   TempKey(11) = MapText("เลือกจากใบขึ้นอาหาร อื่นๆ")
   TempID(11) = 11
   TempKey(12) = MapText("-")
   TempID(12) = 12
   TempKey(13) = MapText("เลือกจากใบฝากขาย")
   TempID(13) = 13
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub
Private Sub GenerateLoadGoodsAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการสินค้าจากใบ SALE ORDER")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("เลือกจาก INVOICE")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("เพิ่มรายละเอียด")
   TempID(5) = 5
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'lMenuChosen = oMenu.Popup("พิมพ์บนกระดาษเปล่า", "ปรับค่าหน้ากระดาษ", "-", "พิมพ์บนฟอร์มเปล่า (เหมือนจริง)", "พิมพ์บนฟอร์มเปล่า (ข้อมูล)", "ปรับค่าหน้ากระดาษ")
Private Sub GeneratePrintMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(35) As String
Dim TempID(35) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("พิมพ์บนกระดาษเปล่า")
   TempID(1) = 1
   TempKey(2) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("พิมพ์บนฟอร์มเปล่า (เหมือนจริง)")
   TempID(4) = 4
   TempKey(5) = MapText("พิมพ์บนฟอร์มเปล่า (ข้อมูล)")
   TempID(5) = 5
   TempKey(6) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(6) = 8
   TempKey(7) = MapText("พิมพ์บนฟอร์มเปล่า (เหมือนใบอ้างอิง)")
   TempID(7) = 6
   TempKey(8) = MapText("พิมพ์บนฟอร์มเปล่า (ข้อมูลใบอ้างอิง)")
   TempID(8) = 7
    TempKey(9) = MapText("-")
    TempID(9) = 9
    TempKey(10) = MapText("พิมพ์บนกระดาษเปล่า (รวมชื่อบริษัท)")
   TempID(10) = 10
   TempKey(11) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(11) = 11
   TempKey(12) = MapText("-")
    TempID(12) = 12
    TempKey(13) = MapText("พิมพ์บนกระดาษเปล่า (ไม่รวมชื่อบริษัท)")
   TempID(13) = 13
   TempKey(14) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(14) = 14
   TempKey(15) = MapText("-")
    TempID(15) = 15
   TempKey(16) = MapText("แบบฟอร์มขออนุมัติขายเกินวงเงิน")
   TempID(16) = 16
   TempKey(17) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(17) = 17
   TempKey(18) = MapText("-")
    TempID(18) = 18
   TempKey(19) = MapText("แบบฟอร์มใบ SO")
   TempID(19) = 19
   TempKey(20) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(20) = 20
   TempKey(21) = MapText("-")
    TempID(21) = 21
   TempKey(22) = MapText("แบบฟอร์มใบสำคัญรับ")
   TempID(22) = 22
   TempKey(23) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(23) = 23
   TempKey(24) = MapText("-")
    TempID(24) = 24
   TempKey(25) = MapText("แบบฟอร์มใบสำคัญจ่าย")
   TempID(25) = 25
   TempKey(26) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(26) = 26
   TempKey(27) = MapText("พิมพ์ฟอร์มสยามธรรมาภิบาล")
   TempID(27) = 27
   TempKey(28) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(28) = 28
   TempKey(29) = MapText("-")
    TempID(29) = 29
   TempKey(30) = MapText("ดูตัวอย่างเอกสาร")
   TempID(30) = 30
   TempKey(31) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(31) = 31
   TempKey(32) = MapText("-")
   TempID(32) = 32
   TempKey(33) = MapText("พิมพ์บิลโดยไม่แสดงราคา")
   TempID(33) = 33
   TempKey(34) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(34) = 34
   TempKey(35) = MapText("-")
   TempID(35) = 35
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'oMenu.Popup("เพิ่มรายการใหม่ (ทีละรายการ)", "เพิ่มรายการใหม่ (เป็นชุด)")
Private Sub GeneratePOAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการใหม่ (ทีละรายการ)")
   TempID(1) = 1
   TempKey(2) = MapText("เพิ่มรายการใหม่ (เป็นชุด)")
   TempID(2) = 2
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'oMenu.Popup("สร้างรายการใหม่", "-", "สร้างจากใบแจ้งหนี้", "-", "สร้างจากใบกำกับภาษี")
Private Sub GenerateAddReceiptMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("สร้างรายการใหม่ (ขายสด)")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   
   TempKey(3) = MapText("สร้างจากใบส่งของ (รับชำระหนี้)")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("สร้างจากใบกำกับภาษี")
   TempID(5) = 5
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

Private Sub GenerateSellMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(37) As String
Dim TempID(37) As Long
Dim D As CMenuItem
    
   TempKey(1) = MapText("ใบ Sale Order")
   TempID(1) = 29
   TempKey(2) = MapText("-")
   TempID(2) = 28
   TempKey(3) = MapText("ใบเสนอราคาขาย")
   TempID(3) = 1
   TempKey(4) = MapText("-")
   TempID(4) = 2
   TempKey(5) = MapText("ใบรับงาน/สั่งงาน (PO)")
   TempID(5) = 3
   TempKey(6) = MapText("-")
   TempID(6) = 4
   TempKey(7) = MapText("ใบส่งของ")
   TempID(7) = 5
   TempKey(8) = MapText("-")
   TempID(8) = 6
   TempKey(9) = MapText("ใบกำกับภาษี")
   TempID(9) = 7
   TempKey(10) = MapText("-")
   TempID(10) = 8
   TempKey(11) = MapText("ใบเสร็จรับเงิน (ขายสด)")
   TempID(11) = 9
   TempKey(12) = MapText("-")
   TempID(12) = 10
   TempKey(13) = MapText("ใบเสร็จรับเงิน (รับชำระ)")
   TempID(13) = 11
   TempKey(14) = MapText("-")
   TempID(14) = 12
   TempKey(15) = MapText("ใบเพิ่มหนี้")
   TempID(15) = 13
   TempKey(16) = MapText("-")
   TempID(16) = 14
   TempKey(17) = MapText("ใบลดหนี้")
   TempID(17) = 15
   TempKey(18) = MapText("-")
   TempID(18) = 16
   TempKey(19) = MapText("ใบสรุปวางบิล")
   TempID(19) = 17
   TempKey(20) = MapText("-")
   TempID(20) = 18
   TempKey(21) = MapText("ใบบรรจุหีบห่อ")
   TempID(21) = 19
   TempKey(22) = MapText("-")
   TempID(22) = 20
   TempKey(23) = MapText("ใบรับคืนสินค้า")
   TempID(23) = 21
   TempKey(24) = MapText("-")
   TempID(24) = 22
   TempKey(25) = MapText("ใบนำฝากธนาคาร")
   TempID(25) = 23
   TempKey(26) = MapText("-")
   TempID(26) = 24
   TempKey(27) = MapText("ใบนำฝากธนาคาร")
   TempID(27) = 25
   TempKey(28) = MapText("-")
   TempID(28) = 26
   TempKey(29) = MapText("ใบยืนยันการ clearing เช็ค")
   TempID(29) = 27
   TempKey(30) = MapText("-")
   TempID(30) = 30
   TempKey(31) = MapText("ใบตรวจสอบการรับ / คืนเช็ค")
   TempID(31) = 31
   TempKey(32) = MapText("-")
   TempID(32) = 32
   TempKey(33) = MapText("ใบขึ้นอาหาร BAG")
   TempID(33) = 33
   TempKey(34) = MapText("-")
   TempID(34) = 34
   TempKey(35) = MapText("ใบขึ้นอาหาร BULK")
   TempID(35) = 35
   TempKey(36) = MapText("-")
   TempID(36) = 36
   TempKey(37) = MapText("ใบขึ้นอาหาร อื่นๆ")
   TempID(37) = 37
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D, Trim(str(TempID(I))))
         Set D = Nothing
      End If
   Next I
End Sub
Private Sub GenerateTax2Items(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(32) As String
Dim TempID(32) As Long
Dim D As CMenuItem
   
   TempKey(1) = MapText("ข้อมูล ภ.ง.ด. 2")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("ข้อมูลผู้กู้ ภ.ง.ด. 2")
   TempID(3) = 3
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D, Trim(str(TempID(I))))
         Set D = Nothing
      End If
   Next I
End Sub

Public Sub GenerateAddJopProcessItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(7) As String
Dim TempID(7) As Long
Dim D As CMenuItem
   
   TempKey(1) = MapText("เพิ่มรายการใหม่ทั่วไป")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("เลือกจากสูตร")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("เพิ่มรายการใหม่ RE BAG -> BAG")
   TempID(5) = 5
   TempKey(6) = MapText("-")
   TempID(6) = 6
   TempKey(7) = MapText("เพิ่มรายการใหม่ RE BAG -> BULK")
   TempID(7) = 7
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D, Trim(str(TempID(I))))
         Set D = Nothing
      End If
   Next I
End Sub

Private Sub GenerateReportMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย")
   D.KEYWORD = ROOT_TREE & " A-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/ลูกค้า (SL001)")
   D.KEYWORD = ROOT_TREE & " 5-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงานขาย (SL001.0)")
   D.KEYWORD = ROOT_TREE & " 5-1-0"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงานขาย (SL001.0.1)")
   D.KEYWORD = ROOT_TREE & " 5-1-0-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปน้ำหนักขายตามสินค้า/ลูกค้า (SL001.1)")
   D.KEYWORD = ROOT_TREE & " 5-1-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามเซ็ตสินค้า/ลูกค้า (SL001.2)")
   D.KEYWORD = ROOT_TREE & " 5-1-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '3
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการขายรายวันตามลูกค้า (SL002)")
   D.KEYWORD = ROOT_TREE & " 5-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '3
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการขายตามลูกค้า (SL002.1)")
   D.KEYWORD = ROOT_TREE & " 5-2-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการขายรายวันตามบิล (SL002.2)")
   D.KEYWORD = ROOT_TREE & " 5-2-1-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการขายรายวันตามบิลตามดิว (SL002.3)")
   D.KEYWORD = ROOT_TREE & " 5-2-1-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปใบสั่งขายตามลูกค้า (SL002.3)")
   D.KEYWORD = ROOT_TREE & " 5-2-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '4
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการขายรายวัน (SL003)")
   D.KEYWORD = ROOT_TREE & " 5-3"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '5
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการขายรายวัน (SL004)")
   D.KEYWORD = ROOT_TREE & " 5-4"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '6
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงาน (SL005.1)")
   D.KEYWORD = ROOT_TREE & " 5-5"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงาน (SL005.2)")
   D.KEYWORD = ROOT_TREE & " 5-5-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงาน รายไตรมาส (SL005.1-2)")
   D.KEYWORD = ROOT_TREE & " 5-5-1-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า รายไตรมาส (SL005.1-3)")
   D.KEYWORD = ROOT_TREE & " 5-5-1-3"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/ลูกค้า (SL005.3)")
   D.KEYWORD = ROOT_TREE & " 5-5-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

'7
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ/ลูกค้า (SL006)")
'   D.KEYWORD = ROOT_TREE & " 5-6"
'   D.PARENT_KEY = ROOT_TREE & " A-1"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ 1 (SL006)")
   D.KEYWORD = ROOT_TREE & " A-1-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ 2 (SL007)")
   D.KEYWORD = ROOT_TREE & " A-1-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานราคาขายสินค้ารายลูกค้า (SL008)")
   D.KEYWORD = ROOT_TREE & " A-1-3"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานราคาขายสินค้ารายลูกค้า รายพนักงานขาย (SL009)")
   D.KEYWORD = ROOT_TREE & " A-1-4"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงาน ใบวางบิล การขายรายวันแบบละเอียด ตามรายลูกค้า (SL010)")
   D.KEYWORD = ROOT_TREE & " A-1-5"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย อิงราคาหน้าโรง")
   D.KEYWORD = ROOT_TREE & " A-7"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานราคาขายสินค้ารายวัน อิงราคาส่วนกลาง (SC001)")
   D.KEYWORD = ROOT_TREE & " A-7-1"
   D.PARENT_KEY = ROOT_TREE & " A-7"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานราคาขายสินค้ารายวัน อิงราคาส่วนกลาง แยกตามลูกค้า (SC002)")
   D.KEYWORD = ROOT_TREE & " A-7-2"
   D.PARENT_KEY = ROOT_TREE & " A-7"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานราคาขายสินค้ารายวัน อิงราคาส่วนกลาง แยกตามพนักงานขาย (SC003)")
   D.KEYWORD = ROOT_TREE & " A-7-3"
   D.PARENT_KEY = ROOT_TREE & " A-7"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   

   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบลูกหนี้")
   D.KEYWORD = ROOT_TREE & " A-2"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานลูกหนี้ค้างชำระรายตัว (AR001)")
   D.KEYWORD = ROOT_TREE & " A-2-1"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปลูกหนี้ค้างชำระรายตัว (AR002)")
   D.KEYWORD = ROOT_TREE & " A-2-2"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการ์ดลูกหนี้ (AR003)")
   D.KEYWORD = ROOT_TREE & " A-2-3"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการ์ดลูกหนี้ (AR003.1)")
   D.KEYWORD = ROOT_TREE & " A-2-3.1"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานวิเคราะห์อายุลูกหนี้ (AR004)")
'   D.KEYWORD = ROOT_TREE & " A-2-4"
'   D.PARENT_KEY = ROOT_TREE & " A-2"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานวิเคราะห์อายุลูกหนี้ (AR004)")
   D.KEYWORD = ROOT_TREE & " A-2-4-1"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
  Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานหนี้คงเหลือตามเช็ครับ (AR004.1)")
   D.KEYWORD = ROOT_TREE & " A-2-4-3"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเช็คคืน (AR004.2)")
   D.KEYWORD = ROOT_TREE & " A-2-4-5"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเช็ครับคงเหลือในมือ(AR004.3)")
   D.KEYWORD = ROOT_TREE & " A-2-4-6"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปวิเคราะห์อายุลูกหนี้ (AR005)")
   D.KEYWORD = ROOT_TREE & " A-2-9"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("วิเคราะห์อายุหนี้ตามเกรดลูกค้าแบบ D-E (AR006)")
'   D.KEYWORD = ROOT_TREE & " A-2-10"
'   D.PARENT_KEY = ROOT_TREE & " A-2"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing

'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานการนำฝากธนาคาร (AR005)")
'   D.KEYWORD = ROOT_TREE & " A-2-5"
'   D.PARENT_KEY = ROOT_TREE & " A-2"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานเงินสดคงเหลือ (AR006)")
'   D.KEYWORD = ROOT_TREE & " A-2-6"
'   D.PARENT_KEY = ROOT_TREE & " A-2"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการรับชำระตามประเภท/ลูกค้า (AR007)")
   D.KEYWORD = ROOT_TREE & " A-2-7"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการรับชำระตามประเภท/ลูกค้า(ใหม่ )(AR008)")
   D.KEYWORD = ROOT_TREE & " A-2-7-1"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการเคลื่อนไหวเงินสด (AR008)")
   D.KEYWORD = ROOT_TREE & " A-2-11"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการรับชำระเงินตามลูกค้า (AR009)")
   D.KEYWORD = ROOT_TREE & " A-2-8"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการรับชำระเงิน (AR009.1)")
   D.KEYWORD = ROOT_TREE & " A-2-17"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานวิเคราะห์อายุดิวเช็ค (AR010)")
   D.KEYWORD = ROOT_TREE & " A-2-12"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานส่วนต่างการรับชำระ (AR011)")
   D.KEYWORD = ROOT_TREE & " A-2-13"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการตามหนี้ระหว่างกัน (AR012)")
   D.KEYWORD = ROOT_TREE & " A-2-14"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบเพิ่มหนี้/ลดหนี้ (AR013)")
   D.KEYWORD = ROOT_TREE & " A-2-15"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานยอดซื้อเชื่อและการชำระหนี้ (AR014)")
   D.KEYWORD = ROOT_TREE & " A-2-16"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานบัญชีสอบทาน (AR018)")
   D.KEYWORD = ROOT_TREE & " A-2-18"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานติดตามงานด้านลูกหนี้เกินดิว ตามพนักงานขาย (AR019)")
   D.KEYWORD = ROOT_TREE & " A-2-4-2"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานแจกแจงรายละเอียดการติดตามงานด้านลูกหนี้เกินดิว ตามพนักงานขาย (AR020)")
   D.KEYWORD = ROOT_TREE & " A-2-4-4"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานประมาณการการรับและเปรียบเทียบหนี้คงค้าง (AR021)")
   D.KEYWORD = ROOT_TREE & " A-2-21"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานแจกแจงรายละเอียดการติดตามงานด้านลูกหนี้เกินดิว ตามระดับลูกค้า (AR022)")
   D.KEYWORD = ROOT_TREE & " A-2-4-22"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   'A-2-17 is already used
   #If LIMIT_AREA <> 1 Then
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบกำไรขาดทุน")
   D.KEYWORD = ROOT_TREE & " F-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานกำไรขาดทุน แยกตามลูกค้า/สินค้า (GP001)")
   D.KEYWORD = ROOT_TREE & " F-1-1"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
      '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานกำไรขาดทุน รายเดือน รายปี แยกตามลูกค้า/สินค้า (GP001-2)")
   D.KEYWORD = ROOT_TREE & " F-1-1-2"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานกำไรขาดทุน แยกตามสินค้า/ลูกค้า (GP002)")
   D.KEYWORD = ROOT_TREE & " F-1-2"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานกำไรขาดทุน แยกตามประเภทลูกค้า/สินค้า (GP003)")
   D.KEYWORD = ROOT_TREE & " F-1-3"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานกำไรขาดทุน แยกตามพนักงานขาย/สินค้า (GP004)")
   D.KEYWORD = ROOT_TREE & " F-1-4"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงาน ACTUAL GP ตามประเภทลูกค้า (GP005)")
   D.KEYWORD = ROOT_TREE & " F-1-5"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานกำไรขาดทุน แยกตามลูกค้า-พนักงานขาย/สินค้า (GP006)")
   D.KEYWORD = ROOT_TREE & " F-1-6"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงาน ACTUAL GP แจกแจงค่าใช้จ่าย (GP007)")
   D.KEYWORD = ROOT_TREE & " F-1-7"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงาน สรุปผลกำไรขาดทุน ขั้นต้น แบบละเอียด (GP008)")
   D.KEYWORD = ROOT_TREE & " F-1-8"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงาน สรุปผลกำไรขาดทุน ขั้นต้น เรียงตามประเภทสินค้า และเบอร์อาหาร (GP009)")
   D.KEYWORD = ROOT_TREE & " F-1-9"
   D.PARENT_KEY = ROOT_TREE & " F-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบบัญชีคุม")
   D.KEYWORD = ROOT_TREE & " J-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("สมุดรายวันขาย (SV)")
   D.KEYWORD = ROOT_TREE & " J-1-1"
   D.PARENT_KEY = ROOT_TREE & " J-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบสำคัญโอนบัญชี  (JV002)")
   D.KEYWORD = ROOT_TREE & " J-1-2"
   D.PARENT_KEY = ROOT_TREE & " J-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   
         
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบเจ้าหนี้")
   D.KEYWORD = ROOT_TREE & " AP-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลเจ้าหนี้  (AP001)")
   D.KEYWORD = ROOT_TREE & " BUY-1-1"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานวิเคราะห์อายุเจ้าหนี้ (AP002)")
   D.KEYWORD = ROOT_TREE & " AP-1-3"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปวิเคราะห์อายุเจ้าหนี้ (AP003)")
   D.KEYWORD = ROOT_TREE & " AP-1-4"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเช็คจ่ายลงวันที่ล่วงหน้า (AP004)")
   D.KEYWORD = ROOT_TREE & " AP-1-1"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("พิมพ์เช็คจ่าย (AP006)")
   D.KEYWORD = ROOT_TREE & " AP-1-2"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการ์ดเจ้าหนี้ (AP007)")
   D.KEYWORD = ROOT_TREE & " AP-1-5"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการ์ดเจ้าหนี้ (AP007-1)")
   D.KEYWORD = ROOT_TREE & " AP-1-5-1"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการ์ดเจ้าหนี้ (AP008)")
   D.KEYWORD = ROOT_TREE & " AP-1-6"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการ์ดเจ้าหนี้ (AP008-1)")
   D.KEYWORD = ROOT_TREE & " AP-1-6-1"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปประมาณการจ่ายเจ้าหนี้การค้า (AP009)")
   D.KEYWORD = ROOT_TREE & " AP-1-7"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการรับเช็คเจ้าหนี้การค้า (AP010)")
   D.KEYWORD = ROOT_TREE & " AP-1-8"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปจ่ายเจ้าหนี้ตามประมาณการ (AP011)")
   D.KEYWORD = ROOT_TREE & " AP-1-9"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานประมาณการจ่ายเจ้าหนี้ (AP012)")
   D.KEYWORD = ROOT_TREE & " AP-1-10"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเปรียบเทียบPOรับและใบรับของ(AP013)")
   D.KEYWORD = ROOT_TREE & " AP-1-11"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
    Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบเข้าค้างส่ง(AP014)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-1"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
 
   
    Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานติดตามการสั่งซื้อ(AP015)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-2"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
    Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบเข้าค้างส่งรวมตามวัตถุดิบ(AP016)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-3"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบเข้าค้างส่งรวมตามซัพพลายเออร์(AP017)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-4"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานติดตามต้นทุน(AP020)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-5"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานค่าขนส่ง(AP021)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-6"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานติดตามค่าขนส่ง ตามบิลขาย(AP022)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-7"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานติดตามค่าขนส่ง ตามใบวางบิล(AP023)")
   D.KEYWORD = ROOT_TREE & " AP-1-11-8"
   D.PARENT_KEY = ROOT_TREE & " AP-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบ(ซื้อ)")
   D.KEYWORD = ROOT_TREE & " AB-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบรับสินค้าเรียงตามเลขที่ (AB001)")
   D.KEYWORD = ROOT_TREE & " AB-1-1"
   D.PARENT_KEY = ROOT_TREE & " AB-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานใบรับสินค้าเรียงตามวันที่ (AB002)")
'   D.KEYWORD = ROOT_TREE & " AB-1-2"
'   D.PARENT_KEY = ROOT_TREE & " AB-1"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบรับสินค้าเรียงตามซัพพลายเออร์ (AB003)")
   D.KEYWORD = ROOT_TREE & " AB-1-3"
   D.PARENT_KEY = ROOT_TREE & " AB-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานใบรับสินค้าเรียงตามซัพพลายเออร์ แหล่งจำหน่าย (AB004)")
   D.KEYWORD = ROOT_TREE & " AB-1-4"
   D.PARENT_KEY = ROOT_TREE & " AB-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   #End If
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateBuyMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(30) As String
Dim TempID(30) As Long
Dim D As CMenuItem
   
   TempKey(1) = MapText("ใบเสนอราคาซื้อ")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
      
   TempKey(3) = MapText("ใบส่งของ")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("ใบกำกับภาษี")
   TempID(5) = 5
   TempKey(6) = MapText("-")
   TempID(6) = 6
   TempKey(7) = MapText("ใบเสร็จรับเงิน")
   TempID(7) = 7
   TempKey(8) = MapText("-")
   TempID(8) = 8
   TempKey(9) = MapText("ใบเพิ่มหนี้")
   TempID(9) = 9
   TempKey(10) = MapText("-")
   TempID(10) = 10
   TempKey(11) = MapText("ใบลดหนี้")
   TempID(11) = 11
   TempKey(12) = MapText("-")
   TempID(12) = 20
   TempKey(13) = MapText("ใบรับคืนสินค้า")
   TempID(13) = 21
   TempKey(14) = MapText("-")
   TempID(14) = 12
   TempKey(15) = MapText("ใบบรรจุหีบห่อ")
   TempID(15) = 13
   TempKey(16) = MapText("-")
   TempID(16) = 14
   TempKey(17) = MapText("ใบนำฝากธนาคาร")
   TempID(17) = 15
   TempKey(18) = MapText("-")
   TempID(18) = 16
   TempKey(19) = MapText("ใบยืนยันการรับเช็คของซัพพลายเออร์")
   TempID(19) = 17
   TempKey(20) = MapText("-")
   TempID(20) = 18
   TempKey(21) = MapText("ใบยืนยันการ Clearing เช็คจ่าย")
   TempID(21) = 19
   TempKey(22) = MapText("-")
   TempID(22) = 22
   TempKey(23) = MapText("ใบประมาณการจ่าย")
   TempID(23) = 23
   TempKey(24) = MapText("-")
   TempID(24) = 24
   TempKey(25) = MapText("ใบค่าใช้จ่ายโรงงาน (เงินสดย่อย)")
   TempID(25) = 25
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D, Trim(str(TempID(I))))
         Set D = Nothing
      End If
   Next I
End Sub

Private Sub GenerateMasterCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
      
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักส่วนกลาง")
   D.KEYWORD = "MASTER_MAIN"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบคลัง")
   D.KEYWORD = "MASTER_INVENTORY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบฝ่ายบุคคล")
   D.KEYWORD = "MASTER_HR"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบบริหารบัญชี")
   D.KEYWORD = "MASTER_LEDGER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักแพคเกจสินค้า")
   D.KEYWORD = "MASTER_PACKAGE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบการผลิต")
   D.KEYWORD = "MASTER_PRODUCTION"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateMainCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
         
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลบริษัท")
   D.KEYWORD = "MAIN_ENTERPRISE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลลูกค้า")
   D.KEYWORD = "MAIN_CUSTOMER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลซัพพลายเออร์")
   D.KEYWORD = "MAIN_SUPPLIER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลพนักงาน")
   D.KEYWORD = "MAIN_EMPLOYEE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลพนักงาน ฟรีแลนช์")
   D.KEYWORD = "MAIN_FREELANCE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลกลาง")
   D.KEYWORD = "MAIN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateProdCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลสูตรการผลิต")
   D.KEYWORD = "PRODUCT_FORMULA"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลใบสั่งผลิต")
   D.KEYWORD = "PRODUCT_JOB"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลใบสั่งแพ็คอาหาร")
   D.KEYWORD = "PRODUCT_PACK"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการผลิต")
   D.KEYWORD = "PRODUCT_JOB"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการปันค่าใช้จ่ายผลิต")
   D.KEYWORD = "PRODUCT_ESTIMATE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการวางแผนการผลิต")
   D.KEYWORD = "PRODUCT_PLAN"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบการผลิต")
   D.KEYWORD = "PRODUCT_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateStockCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูล MASTER")
   D.KEYWORD = "INVENTORY_PART-MASTER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลสินค้าและวัตถุดิบ/วัสดุอุปกรณ์")
   D.KEYWORD = "INVENTORY_PART"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   'เล็กปิดใช้งาน 2017-07-31
''   Set D = New CMenuItem
''   D.MENU_TEXT = MapText("ข้อมูลการรับเข้าคลัง")
''   D.KEYWORD = "INVENTORY_IMPORT"
''   D.PARENT_KEY = ""
''   D.ICON_INDEX1 = 1
''   D.ICON_INDEX2 = 2
''   Call Col.add(D)
''   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการเบิกจากคลัง")
   D.KEYWORD = "INVENTORY_EXPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการโอนย้าย")
   D.KEYWORD = "INVENTORY_TRANSFER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการปรับยอดคงคลัง")
   D.KEYWORD = "INVENTORY_ADJUST"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการนับจริงจากคลัง")
   D.KEYWORD = "INVENTORY_ACTUAL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบคลัง")
   D.KEYWORD = "INVENTORY_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateStockWHCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการรับเข้า")
   D.KEYWORD = "INVENTORY-WH_IMPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการจ่ายออก")
   D.KEYWORD = "INVENTORY-WH_EXPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการโอนย้าย")
   D.KEYWORD = "INVENTORY-WH_TRANSFER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการปรับยอดคงคลัง")
   D.KEYWORD = "INVENTORY-WH_ADJUST"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
'''   Set D = New CMenuItem
'''   D.MENU_TEXT = MapText("ข้อมูลการนับจริงจากคลังสินค้า")
'''   D.KEYWORD = "INVENTORY-WH_ACTUAL"
'''   D.PARENT_KEY = ""
'''   D.ICON_INDEX1 = 1
'''   D.ICON_INDEX2 = 2
'''   Call Col.add(D)
'''   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลสินค้าคงเหลือ")
   D.KEYWORD = "INVENTORY-WH_STOCK"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบคลังสินค้า")
   D.KEYWORD = "INVENTORY-WH_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GeneratePackageCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลสินค้า/บริการ")
   D.KEYWORD = "PACKAGE_FEATURE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลแพคเกจสินค้า/บริการ")
   D.KEYWORD = "PACKAGE_SOC"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ราคาประกาศกลาง หน้าโรงงาน")
    D.KEYWORD = "PACKAGE-CENTER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานแพคเกจสินค้า/บริการ")
   D.KEYWORD = "PACKAGE_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateAdminCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลกลุ่มผู้ใช้งาน")
   D.KEYWORD = "ADMIN_GROUP"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลผู้ใช้งาน")
   D.KEYWORD = "ADMIN_USER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลผู้ใช้งาน")
   D.KEYWORD = "ADMIN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateLedgerCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลอัตราการแลกเปลี่ยนเงินตรา")
   D.KEYWORD = "LEDGER_CURRENCY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลอัตราการแลกเปลี่ยนเงินตรา")
   D.KEYWORD = "LEDGER_CURRENCYEX"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย")
   D.KEYWORD = "LEDGER_SELL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบซื้อเข้า")
   D.KEYWORD = "LEDGER_STOCK_BUY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานซื้อ (รายจ่าย)")
   D.KEYWORD = "LEDGER_BUY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบข้อมูลเงินสด")
   D.KEYWORD = "LEDGER_CASH"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบบัญชี")
   D.KEYWORD = "LEDGER_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GeneratePlanCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ประมาณการการใช้วัตถุดิบและผลิตสินค้ารายวัน")
   D.KEYWORD = "PLANNING_1"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ประมาณการการใช้วัตถุดิบและผลิตสินค้ารายสัปดาห์")
   D.KEYWORD = "PLANNING_2"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ประมาณการการใช้วัตถุดิบและผลิตสินค้ารายเดือน")
   D.KEYWORD = "PLANNING_4"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ประมาณการรับเข้าวัตถุดิบรายวันจากซัพพลายเออร์")
   D.KEYWORD = "PLANNING_3"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบประมาณการ/วางแผน")
   D.KEYWORD = "PLAN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '===
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateCommissionCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("เป้าการขาย")
   D.KEYWORD = "COMMISSION_TARGET"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("เงื่อนไขคอมมิตชั่น")
   D.KEYWORD = "COMMISSION_CONDITION"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("แผนภูมิพนักงานขาย")
   D.KEYWORD = "COMMISSION_ORGANIZE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลส่วนหัก GP - ส่งเสริม")
   D.KEYWORD = "COMMISSION_SUBTRACT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("เงื่อนไข INCENTIVE")
   D.KEYWORD = "COMMISSION_INCENTIVE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("ข้อมูล INCENTIVE รายลูกค้า สินค้า")
'   D.KEYWORD = "COMMISSION_INCENTIVE-CUS-PD"
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลต้นทุนรายเบอร์")
   D.KEYWORD = "COMMISSION_COST"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ตารางปรับเครดิตลูกค้า")
   D.KEYWORD = "COMMISSION_CREDIT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบคอมมิตชั่น")
   D.KEYWORD = "COMMISSION_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GeneratePictures(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = PictureTypeToText(HEAD_ACCOUNT)
   TempID(1) = 1
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

Public Sub CreateGuiConfig(Programowner As String)
   '====
    Call CreateGuiConfigObj("LOGO_VIEW", False)
    Call CreateGuiConfigObj("AGENT_SERVER_CONNECT", False)
   
    Call SetGuiConfigObj("AGENT_SERVER_CONNECT", True)

   '====
   Call CreateGuiConfigObj("HR_VIEW", True)
   Call CreateGuiConfigObj("PRODUCTION_VIEW", True)
   Call CreateGuiConfigObj("PACKAGE_VIEW", True)
   Call CreateGuiConfigObj("LEDGER_VIEW", True)
   Call CreateGuiConfigObj("INVENTORY_VIEW", True)
   Call CreateGuiConfigObj("INVENTORY-WH_VIEW", True)
   Call CreateGuiConfigObj("ENTERPRISE_INFO_EDIT", True)
   Call CreateGuiConfigObj("PLAN_VIEW", True)
   Call CreateGuiConfigObj("COMMISSION_VIEW", True)
   '====

      
   If Programowner = MM_OWNER Then 'โรงงานอาหารสัตว์ MM
      Call SetGuiConfigObj("HR_VIEW", False)
      Call SetGuiConfigObj("PACKAGE_VIEW", True)
      Call SetGuiConfigObj("LEDGER_VIEW", True)
      Call SetGuiConfigObj("AGENT_SERVER_CONNECT", False)
      
      Call GenerateSellMenuItems("YYNNNNYYNNNNYYYYYNNNNYYYNNYYYYYYYYYYY", SellMenuItems, Programowner)
      Call GenerateBuyMenuItems("NNNNNNYYYYYYYNNNNNNNNNNYY", BuyMenuItems, Programowner)
      Call GenerateReportMenuItems("YYYYYYYYYYYYYYYYYYYYYYY", ReportMenuItems, Programowner)
      Call GenerateDOAddMenuItems("YNYNNNYYYYYYY", DOAddMenuItems, Programowner) 'YYYNNNYY
      Call GeneratePOAddMenuItems("YY", POAddMenuItems, Programowner)
      Call GenerateDOAddMenuItems("NNNYNN", SupAddMenuItems, Programowner)
      Call GenerateDOAddMenuItems("YNNNNN", SupAddNoPoMenuItems, Programowner)
      Call GenerateLoadGoodsAddMenuItems("YNNYY", LoadGoodsAddMenuItems, Programowner) '
      Call GenerateAddReceiptMenuItems("YYYNN", ReceiptAddMenuItems, Programowner)
      Call GenerateSupItemOtherAddMenuItems("YYYYYYYYYYY", SupItemPoOtherMenuItems, Programowner)
'      Call GenerateSupItemOtherAddMenuItems("YYYYY", SupItemPoOtherMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYNNNN", QuoatationPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYYYYY", POPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("NNNNNNNNNNNNNNNYYYYY", SOPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYYNNNNNNNNNNNNNNNNNNNNNNNYYYYYYYY", DOPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYNNNN", IVPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYYYYY", RCPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYYYYYNNNNNNNNNNNNNNYYY", RCPrintMenuItemsSpacial, Programowner)
      Call GeneratePrintMenuItems("NNNNNNNNNNNNNNNNNNNNNNNNYY", RCPrintMenuItemsSpacialBuy, Programowner)
      Call GeneratePrintMenuItems("YYNNNN", CdnDbnPrintMenuItems, Programowner)
      Call GeneratePrintMenuItems("YYNNNN", BSPrintMenuItems, Programowner)

      Call GenerateMainCommandMenuItems("YYYYY", MainCommandMenuItems, Programowner)
      Call GenerateMasterCommandMenuItems("YYNYYY", MasterCommandMenuItems, Programowner)
      Call GenerateAdminCommandMenuItems("YYY", AdminCommandMenuItems, Programowner)
      Call GenerateProdCommandMenuItems("YYYY", ProdCommandMenuItems, Programowner)
      Call GeneratePackageCommandMenuItems("YYY", PackageCommandMenuItems, Programowner)
      #If LIMIT_AREA <> 1 Then
         Call GenerateLedgerCommandMenuItems("NNYYYNY", LedgerCommandMenuItems, Programowner)
      #Else
         Call GenerateLedgerCommandMenuItems("NNNNNNY", LedgerCommandMenuItems, Programowner)
      #End If
      Call GenerateStockCommandMenuItems("YYYYYY", StockCommandMenuItems, Programowner)
      Call GenerateStockWHCommandMenuItems("YYYYYY", StockWHCommandMenuItems, Programowner)
      Call GeneratePlanCommandMenuItems("YYY", PlanCommandMenuItems, Programowner)
      
      
      Call GenerateCommissionCommandMenuItems("YYYYYY", CommissionCommandMenuItems, Programowner)
      Call GenerateAddJopProcessItems("YYYYYYY", LoadAddJopProcessItems, Programowner)
      'GenerateAddJopProcessItems
      
      Call GeneratePictures("Y", PictureItems, Programowner)
   
   End If
End Sub

Public Function VerifyGuiConfig(KeyCode As String, Optional PrintError As Boolean = False) As Boolean
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(KeyCode)
   VerifyGuiConfig = TempC.ENABLE_FLAG
   
   If Not TempC.ENABLE_FLAG Then
      If PrintError Then
         glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
         glbErrorLog.ShowUserError
      End If
   End If
End Function

Public Function ShowWindowCaption(Programowner As String) As String
   If Programowner = MM_OWNER Then
      ShowWindowCaption = MapText("ระบบบริหารงาน WIN Pricing 400")
   End If
End Function
Private Sub GenerateSupItemOtherAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("อนุมัติ PO")
      TempID(1) = 1
   TempKey(2) = MapText("-")
      TempID(2) = 2
   TempKey(3) = MapText("ยกเลิก การอนุมัติ PO")
      TempID(3) = 3
   TempKey(4) = MapText("-")
      TempID(4) = 4
   TempKey(5) = MapText("ปิด PO ใบนี้")
      TempID(5) = 5
   TempKey(6) = MapText("-")
      TempID(6) = 6
   TempKey(7) = MapText("ยกเลิก การปิด PO")
      TempID(7) = 7
   TempKey(8) = MapText("-")
      TempID(8) = 8
   TempKey(9) = MapText("ตรวจสอบ PO ใบนี้แล้ว")
      TempID(9) = 9
   TempKey(10) = MapText("-")
      TempID(10) = 10
   TempKey(11) = MapText("ยกเลิกการตรวจสอบ PO ใบนี้")
      TempID(11) = 11
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

