Attribute VB_Name = "modMain"
 Option Explicit
'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\somepath\mydb.mdb;User Id=admin;Password=;"
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const ROOT_TREE = "Root"
Public Const DUMMY_KEY = 27
Public GLB_GRID_COLOR As Long
Public GLB_NORMAL_COLOR As Long
Public GLB_ALERT_COLOR As Long
Public GLB_SHOW_COLOR As Long
Public GLB_FORM_COLOR As Long
Public GLB_HEAD_COLOR As Long
Public GLB_GRIDHD_COLOR As Long
Public GLB_MANDATORY_COLOR As Long
Public IP_ADDRESS As String
Public Const MM_OWNER = "MM_FACT"

Public Enum PACKAGE_TYPE
   PACKAGE_BAG = 1
   PACKAGE_BULK = 2
   PACKAGE_OTH = 3
End Enum

Public Enum PICTURE_TYPE
   HEAD_ACCOUNT = 1
   HEAD_PART = 2
End Enum

Public Enum RATE_TYPE
   RATE_FLAT = 1
   RATE_STEP = 2
   RATE_TIER = 3
End Enum

Public Enum DO_RATE_TYPE
   RATE_MASTER = 1
   RATE_CUSTOM = 2
End Enum

Public Enum FIELD_TYPE
   INT_TYPE = 1
   MONEY_TYPE = 2
   DATE_TYPE = 3
   STRING_TYPE = 4
   BOOLEAN_TYPE = 5
End Enum

Public Enum FIELD_CAT
   ID_CAT = 1
   MODIFY_DATE_CAT = 2
   CREATE_DATE_CAT = 3
   MODIFY_BY_CAT = 4
   CREATE_BY_CAT = 5
   DATA_CAT = 6
   TEMP_CAT = 7
End Enum

Public Enum PAYMENT_TYPE
   CASH_PMT = 1
   CREDITCRD_PMT = 2
   CHECK_PMT = 3
   BANKTRF_PMT = 4
   CASHRET_PMT = 5
End Enum

Public Enum PERIOD_TYPE
   DAILY_PERIOD = 1
   MONTHLY_PERIOD = 2
End Enum

Public Enum SHOW_MODE_TYPE
   SHOW_ADD = 1
   SHOW_EDIT = 2
   SHOW_VIEW = 3
   SHOW_VIEW_ONLY = 4
End Enum

Public Enum RESOURCE_TYPE
   HOTEL_RESOURCE = 1
End Enum

Public Enum TEXT_BOX_TYPE
   TEXT_STRING = 1
   TEXT_INTEGER = 2
   TEXT_FLOAT = 3
   TEXT_FLOAT_MONEY = 4
   TEXT_INTEGER_MONEY = 5
End Enum

Public Enum LANGUAGE_TYPE
   LANG_ENG = 1
   LANG_THAI = 2
End Enum

Public Enum PARCEL_TYPE
   PARCEL_BAG = 1
   PARCEL_BULK = 2
   PARCEL_ALL = 3
End Enum

Public Enum RATIO_TYPE
   RATIO_COST = 1
   RATIO_QUANTITY = 2
   RATIO_RAW = 3
   RATIO_VARY = 4                ' แปรผันตามจำนวนที่ ผลิต  ต่อ  กก
   RATIO_PERCENT = 5                ' คิดเป็นกี่ % ของมูลค่าวัตถุดิบที่ใช้
End Enum

Public Enum JOB_VARIABLE_TYPE
   LOSS_VAR = 1
   OVERHEAD_VAR = 2
   CAN_VAR = 3
   RCM_VAR = 4
   PMC_VAR = 5
End Enum

Public Enum UNIQUE_TYPE
   EMPCODE_UNIQUE = 1
   EMPNAME_LASTNAME_UNIQUE = 2
   TRUCK_UNIQUE = 3
   DO_PLAN_UNIQUE = 4
   DBN_UNIQUE = 5
   CUSTCODE_UNIQUE = 6
   USERGROUP_UNIQUE = 7
   USERNAME_UNIQUE = 8
   IMPORT_UNIQUE = 9
   EXPORT_UNIQUE = 10
   REPAIR_UNIQUE = 11
   REPAIR_FORMULA_UNIQUE = 12
   SUPPLIER_UNIQUE = 13
   PARTNO_UNIQUE = 14
   QUOATATION_UNIQUE = 15
   TEACHER_UNIQUE = 16
   SUBJECT_UNIQUE = 17
   FACULTY_UNIQUE = 18
   EXPENSE_UNIQUE = 19
   PO_UNIQUE = 20
   CUSTOMER_UNIQUE = 21
   REVENUE_UNIQUE = 22
   BORROW_UNIQUE = 23
   PRDFEATURE_UNIQUE = 24
   JOBPLAN_UNIQUE = 25
   
   PARTTYPE_NO = 26
   PARTTYPE_NAME = 27
   LOCATION_NO = 28
   LOCATION_NAME = 29
   PRODUCTTYPE_NO = 30
   PRODUCTTYPE_NAME = 31
   PRODUCTSTATUS_NO = 32
   PRODUCTSTATUS_NAME = 33
   HOUSE_NO = 34
   HOUSE_NAME = 35
   COUNTRY_NO = 36
   COUNTRY_NAME = 37
   CSTGRADE_NO = 38
   CSTGRADE_NAME = 39
   CSTTYPE_NO = 40
   CSTTYPE_NAME = 41
   SUPPLIERTYPE_NO = 42
   SUPPLIERYPE_NAME = 43
   SUPPLIERGRADE_NO = 44
   SUPPLIERGRADE_NAME = 45
   SUPPLIERSTATUS_NO = 46
   SUPPLIERSTATUS_NAME = 47
   POSITION_NO = 48
   UNIT_NO = 49
   UNIT_NAME = 50
   YEAR_NO = 51
   PARTGROUP_NO = 52
   PARTGROUP_NAME = 53
   FEATURENO_UNIQUE = 54
   SOCNO_UNIQUE = 55
   PERSON_CODE = 56
   RELIGIOUS_NO = 57
   RELIGIOUS_NAME = 58
   WORK_NO = 59
   WORK_NAME = 60
   RESIGN_NO = 61
   RESIGN_NAME = 62
   BANK_NO = 63
   BANK_NAME = 64
   DOC_NO = 65
   DOC_NAME = 66
   MONTHLY_ADD_NO = 67
   MONTHLY_ADD_NAME = 68
   MONTHLY_SUB_NO = 69
   MONTHLY_SUB_NAME = 70
   JOB_NO = 71
   FORMULA_NO = 72
   CURRENCY_NO = 73
   MONEY_FAMILY_NO = 74
   CHEQUEDOC_UNIQUE = 75
   PLANNING_UNIQUE = 76
   INVENTORY_ACT_UNIQUE = 77
   
   TARGET_UNIQUE = 78
   MASTER_VALID_NO = 79
   
   MASTER_COMMISSION_CREDIT = 80
   
   LOT_DOC = 81
   PALLET_DOC = 82
   INVENTORY_WH_DOC_UNIQUE = 83
   PACK_PRODUCTION_UNIQUE = 84
   LOT_UNIQUE = 85
   INVENTORY_WH_ACT_UNIQUE = 86
   INVENTORY_WH_DOC_DATE_UNIQUE = 87
   FREELANCE_UNIQUE = 88
   TRANSPORT_DETAIL_UNIQUE = 89

   INCENTIVE_PD_UNIQUE = 90
   INCENTIVE_PD_CUS_UNIQUE = 91
   
   SUPPLIER_ACCOUNT_NO_UNIQUE = 92
   BILLING_PAYMENT_NO_UNIQUE = 93
   DELIVERY_CUS_UNIQUE = 94
   WORKS_PRICE_ACTIVE_DATE_UNIQUE = 95
   
   PART_MASTER_NO_UNIQUE = 96

   
End Enum

Public Enum MASTER_TYPE
   DRCR_REASON = 1
   EXPENSE_TYPE = 2
   BANK_ACCOUNT = 3
   CHEQUE_TYPE = 4
   PRTITEM_SET = 5
   MEMO_TYPE = 6
   MEMO_STATUS = 7
   EXPORT_DESC = 8
   ACCOUNT_LIST = 9
   PAY_TO = 10
   SET_PRODUCT = 11
   FEATURE_GROUP = 12
   SET_PROJECT = 14
   CONDITION = 15
   PAID_TYPE = 16
   PRODUCT_TYPE = 17
   LOCATION_GROUP = 18
   CUSTOMER_SALE_TYPE = 19
   ANIMAL_TYPE = 20
   TRANSPORT_DETAIL = 21
    PROMOTIONAL_DETAIL = 22

End Enum

Public Enum CONFIG_DOC_TYPE
   SELL_SO = 1
   SELL_RETURN = 2
   
   IV_DO = 3
   
   BUY_PO_RAW = 11
   BUY_PO_MATERIAL = 12
   BUY_PO_EXPENSE = 13
   BUY_PO_GENERAL = 14
   
   BUY_PO_RAW_AUTO = 21
   BUY_PO_MATERIAL_AUTO = 22
   BUY_PO_EXPENSE_AUTO = 23
   BUY_PO_GENERAL_AUTO = 24
   
   BUY_RO_RAW = 31
   BUY_RO_MATERIAL = 32
   BUY_RO_EXPENSE = 33
   BUY_RO_GENERAL = 34
   
  WH_LOAD_GOODS_BAG = 29
  WH_LOAD_GOODS_BULK = 30
  WH_LOAD_GOODS_OTHER = 28
  
 PAYMENT_VOUCHER = 35
 TRANSFER_VOUCHER = 36
   
End Enum

Public Enum NUMBER_TYPE
   PO_NUMBER = 1
   OPERATE_NUMBER = 2
   BORROW_NUMBER = 3
   DEBIT_NOTE_NUMBER = 4
   'bum+
   EXPENSE_NUMBER = 5
   REPAIR_NUMBER = 6
   IMPORT_NUMBER = 7
   EXPORT_NUMBER = 8
   PLAN_NUMBER = 9
   FUEL_NUMBER1 = 10
   FUEL_NUMBER2 = 11
   BILL_NUMBER = 13
   QUOATATION_NUMBER = 14
   REVENUE_NUMBER = 15
   DO_NUMBER = 16
   RECEIPT_NUMBER = 17
   JOBPLAN_NUMBER = 18
   INVOICE_RECEIPT_NUMBER = 19
   BILLS_NUMBER = 20
   DBN_NUMBER = 21
   CDN_NUMBER = 22
   CUSTOMER_NUMBER = 23
   ESTIMATE_NUMBER = 24
   PKGLST_NUMBER = 25
   TRANSFER_NUMBER = 26
   CHEQUE_DOC_NUMBER = 27
   JOBPLAN_AUTO_NUMBER = 28
   LOT_NUMBER = 29
   LOAD_GOODS = 30
   PACK_PRODUCTION = 31
   AVG_GOODS = 32
   BALANCE_GOODS = 33 'ปรับยอดสินค้า
   LOAD_GOODS_BK = 34
   RE_BAG_JOBPLAN_NUMBER = 35
   RE_BULK_JOBPLAN_NUMBER = 36
   BULK_JOBPLAN_NUMBER = 37
   RE_BAG_RM_JOBPLAN_NUMBER = 38
   SUP_TRANSPORT_NUMBER = 39
   EX_WORKS_PRICE = 40
   TRANSFER_RQ_NUMBER = 41
   End Enum

Public Enum CASH_DOC_TYPE
   CASH_TRANSFER = 1
   CASH_DEPOSIT = 2
   CASH_WITHDRAW = 3
   CASH_PITTYCASH = 4
   CASH_WHTHDRAW2 = 5
   CASH_DEPOSIT2 = 6
   POST_CHEQUE = 7
   
   WAITING_CHEQUE = 8                                        'เป็นเอกสารที่ทำเมื่อได้จ่ายเช็คให้ซัพพลายเออร์แล้ว -> เช็ครอเรียกเก็บ
   PASSED_CHEQUE = 9                             'เป็นเอกสารที่บอกว่าเช็คใดบ้างในมือซัพพลายเออร์ที่ขึ้นเงินได้แล้ว -> เช็คผ่านแล้ว
   
End Enum
Public Enum CHEQUE_DOC_TYPE
   CHECK_CHEQUE = 1
End Enum

Public Enum POST_TYPE
   POST_CLEAR = 1                                                      ' ใบเครียร์ ของ เช็คจากลูกค้า
   WAITING_CLEAR = 2                                                      'เช็ครอเรียกเก็บ
   PASSED_CLEAR = 3                                                      'เช็คผ่านแล้ว
End Enum

Public Enum INVENTORY_DOCTYPE
   IMPORT_DOCTYPE = 1
   EXPORT_DOCTYPE = 2
   TRANSFER_DOCTYPE = 3
   ADJUST_DOCTYPE = 4
   DO2IVD_DOCTYPE = 5
   RO2IVD_DOCTYPE = 6
   RCP1_2IVD_DOCTYPE = 7
   RCP2_2IVD_DOCTYPE = 8
   RET1_2IVD_DOCTYPE = 9
   RET2_2IVD_DOCTYPE = 10
End Enum


Public Enum MASTER_COMMISSION_AREA
   COMMISSION_BUDGET_CHART = 1
   COMMISSION_CONDITION = 2
   COMMISSION_COST = 3
End Enum

'===================== For clear treeview =========================
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd _
    As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const TV_FIRST As Long = &H1100
Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Const TVGN_ROOT As Long = &H0
Const WM_SETREDRAW As Long = &HB
'===================== For clear treeview =========================

Public Const PROJECT_NAME = "WINPRICING400_MM"
Public Const GLB_FONT = "JasmineUPC"
Private Const MODULE_NAME = "modMain"

Private m_Conn As ADODB.Connection
Public glbErrorLog As clsErrorLog
Public glbDatabaseMngr As clsDatabaseMngr
Public glbDatabaseMngr2 As clsDatabaseMngr
Public glbSetting As clsGlobalSetting
Public glbParameterObj As clsParameter
Public glbUser As CUser
Public glbGroup As CGroup
Public glbAdmin As clsAdmin
Public glbMaster As clsMaster
Public glbDaily As clsDaily
Public glbLegacy As clsLegacy
Public glbEnterPrise As CEnterprise
Public glbSystemParams As Collection
Public glbProduction As clsProduction
'Public glbProductionWH As clsProductionWH
Public glbPlanning As clsPlanning
Public glbTargetCom As clsTargetCom
Public glbInventoryAct As clsInventoryAct
Public glbInventoryWhAct As clsInventoryWhAct
Public glbAuthenPO As clsAuthenPO
Public glbGuiConfigs As CGuiConfigs
Public glbLoginTracking As CLoginTracking
Public glbAccessRight As Collection
Public glbLockDate As CLockDate

Public glbLoginID As Long
Public m_LoginTracking As Collection
Public Temp_LTK As CLoginTracking

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function VerifyDate(L As Label, D As uctlDate, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not D.VerifyDate(NullAllow) Then
      VerifyDate = False
      D.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyDate = True
   End If
End Function

Public Function VerifyTime(L As Label, T As uctlTime, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not T.VerifyTime(NullAllow) Then
      VerifyTime = False
      T.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyTime = True
   End If
End Function

Public Function VerifyTextData(L As Label, T As TextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextData = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextData = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextData = False
            Else
               VerifyTextData = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextData = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextData = False
               Else
                  VerifyTextData = True
               End If
            End If
            Exit Function
         End If
      End If
      
      VerifyTextData = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextData = False
         Exit Function
      End If
      
      VerifyTextData = True
   End If
End Function

Public Function VerifyTextControl(L As Label, T As uctlTextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextControl = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextControl = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextControl = True
               Exit Function
            Else
               VerifyTextControl = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextControl = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextControl = False
               Else
                  VerifyTextControl = True
                  Exit Function
               End If
            End If
'            Exit Function
         End If
      End If
      
      VerifyTextControl = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextControl = False
         Exit Function
      End If
      
      VerifyTextControl = True
   End If
End Function

Private Function GetParentID(Acc As String) As Long
Dim Key As String
Dim TempLen As Long
Dim I As Long
Dim J As Long

   For I = 1 To Len(Acc)
      J = (Len(Acc) - I) + 1
      If Mid(Acc, J, 1) = "_" Then
         Exit For
      End If
   Next I
   
   Key = Mid(Acc, 1, J)
End Function

Private Sub GetParentItemDesc(Acc As String, Ri As CRightItem, ReportName As String)
   Ri.DEFAULT_VALUE = "N"
   
   If Acc = "ADMIN" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลผู้ใช้งาน"
   ElseIf Acc = "ADMIN_GROUP" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลกลุ่มผู้ใช้งาน"
   ElseIf Acc = "ADMIN_GROUP_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลกลุ่มผู้ใช้งาน"
   ElseIf Acc = "ADMIN_GROUP_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลกลุ่มผู้ใช้งาน"
   ElseIf Acc = "ADMIN_GROUP_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลกลุ่มผู้ใช้งาน"
   ElseIf Acc = "ADMIN_USER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลผู้ใช้งาน"
   ElseIf Acc = "ADMIN_USER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลผู้ใช้งาน"
   ElseIf Acc = "ADMIN_USER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลผู้ใช้งาน"
   ElseIf Acc = "ADMIN_USER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลผู้ใช้งาน"
   ElseIf Acc = "ADMIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลผู้ใช้งาน"
   
   
   
   
   
   ElseIf Acc = "MASTER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลัก"
   ElseIf Acc = "MASTER_MAIN" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักส่วนกลาง"
   ElseIf Acc = "MASTER_INVENTORY" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักคลัง"
   ElseIf Acc = "MASTER_PRODUCTION" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบการผลิต"
   ElseIf Acc = "MASTER_LEDGER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบบัญชี"
   ElseIf Acc = "MASTER_INVENTORY" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบคลัง"
   ElseIf Acc = "MASTER_PRODUCTION" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบการผลิต"
   ElseIf Acc = "MASTER_PRODUCTION" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบการผลิต"
   ElseIf Acc = "MASTER_PACKAGE" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบแพ็คเกจสินค้า/บริการ"
      
   ElseIf Acc = "MAIN" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลส่วนกลาง"
   ElseIf Acc = "MAIN_CUSTOMER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลลูกค้า"
   ElseIf Acc = "MAIN_CUSTOMER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลลูกค้า"
   ElseIf Acc = "MAIN_CUSTOMER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลลูกค้า"
   ElseIf Acc = "MAIN_CUSTOMER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลลูกค้า"
   ElseIf Acc = "MAIN_SUPPLIER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลซัพพลายเออร์"
   ElseIf Acc = "MAIN_SUPPLIER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลซัพพลายเออร์"
   ElseIf Acc = "MAIN_SUPPLIER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลซัพพลายเออร์"
   ElseIf Acc = "MAIN_SUPPLIER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลซัพพลายเออร์"
   ElseIf Acc = "MAIN_ENTERPRISE" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลบริษัท"
   ElseIf Acc = "MAIN_ENTERPRISE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลบริษัท"
   ElseIf Acc = "MAIN_EMPLOYEE" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลพนักงาน"
   ElseIf Acc = "MAIN_EMPLOYEE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลพนักงาน"
   ElseIf Acc = "MAIN_EMPLOYEE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลพนักงาน"
   ElseIf Acc = "MAIN_EMPLOYEE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลพนักงาน"
   ElseIf Acc = "MAIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลส่วนกลาง"
   ElseIf Acc = "MAIN_FREELANCE" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลพนักงานฟรีแลนช์"
   ElseIf Acc = "MAIN_FREELANCE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลพนักงานฟรีแลนช์"
   ElseIf Acc = "MAIN_FREELANCE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลพนักงานฟรีแลนช์"
   ElseIf Acc = "MAIN_FREELANCE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลพนักงานฟรีแลนช์"
      
   ElseIf Acc = "INVENTORY" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลคลัง"
   ElseIf Acc = "INVENTORY_PART-MASTER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูล MASTER"
   ElseIf Acc = "INVENTORY_PART-MASTER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูล MASTER"
   ElseIf Acc = "INVENTORY_PART-MASTER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูล MASTER"
   ElseIf Acc = "INVENTORY_PART-MASTER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูล MASTER"
   ElseIf Acc = "INVENTORY_PART" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลวัตถุดิบ"
   ElseIf Acc = "INVENTORY_PART_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลวัตถุดิบ"
   ElseIf Acc = "INVENTORY_PART_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลวัตถุดิบ"
   ElseIf Acc = "INVENTORY_PART_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลวัตถุดิบ"
      
   ElseIf Acc = "INVENTORY_IMPORT" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการนำเข้าวัตถุดิบ"
   ElseIf Acc = "INVENTORY_IMPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการนำเข้าวัตถุดิบ"
   ElseIf Acc = "INVENTORY_IMPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการนำเข้าวัตถุดิบ"
   ElseIf Acc = "INVENTORY_IMPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการนำเข้าวัตถุดิบ"
      
   ElseIf Acc = "INVENTORY_EXPORT" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการเบิกจ่ายวัตถุดิบ"
   ElseIf Acc = "INVENTORY_EXPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการเบิกจ่ายวัตถุดิบ"
   ElseIf Acc = "INVENTORY_EXPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการเบิกจ่ายวัตถุดิบ"
   ElseIf Acc = "INVENTORY_EXPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการเบิกจ่ายวัตถุดิบ"
   
   ElseIf Acc = "INVENTORY_TRANSFER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการโอนวัตถุดิบ"
   ElseIf Acc = "INVENTORY_TRANSFER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการโอนวัตถุดิบ"
   ElseIf Acc = "INVENTORY_TRANSFER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการโอนวัตถุดิบ"
   ElseIf Acc = "INVENTORY_TRANSFER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการโอนวัตถุดิบ"
      
   ElseIf Acc = "INVENTORY_ADJUST" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการปรับยอดวัตถุดิบ"
   ElseIf Acc = "INVENTORY_ADJUST_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการปรับยอดวัตถุดิบ"
   ElseIf Acc = "INVENTORY_ADJUST_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการปรับยอดวัตถุดิบ"
   ElseIf Acc = "INVENTORY_ADJUST_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการปรับยอดวัตถุดิบ"
      
   ElseIf Acc = "INVENTORY_ACTUAL" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการนับยอดวัตถุดิบคงเหลือ"
   ElseIf Acc = "INVENTORY_ACTUAL_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการนับยอดวัตถุดิบคงเหลือ"
   ElseIf Acc = "INVENTORY_ACTUAL_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการนับยอดวัตถุดิบคงเหลือ"
   ElseIf Acc = "INVENTORY_ACTUAL_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลข้อมูลการนับยอดวัตถุดิบคงเหลือ"
   
   ElseIf Acc = "INVENTORY-WH_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลคลังสินค้า"
  ElseIf Acc = "INVENTORY_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลคลัง"
   ElseIf Acc = "INVENTORY-WH" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลคลังสินค้า"
   ElseIf Acc = "INVENTORY-WH_IMPORT" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการรับเข้าสินค้า"
   ElseIf Acc = "INVENTORY-WH_IMPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการรับเข้าสินค้า"
   ElseIf Acc = "INVENTORY-WH_IMPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการรับเข้าสินค้า"
   ElseIf Acc = "INVENTORY-WH_IMPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการรับเข้าสินค้า"
   ElseIf Acc = "INVENTORY-WH_ACTUAL" Then
      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการนับยอดสินค้าคงเหลือ"
      
'   ElseIf Acc = "INVENTORY_IMPORT" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการนำเข้าวัตถุดิบ"
'
   ElseIf Acc = "INVENTORY-WH_EXPORT" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการเบิกจ่ายสินค้า"
   ElseIf Acc = "INVENTORY-WH_EXPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการเบิกจ่ายสินค้า"
   ElseIf Acc = "INVENTORY-WH_EXPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการเบิกจ่ายสินค้า"
   ElseIf Acc = "INVENTORY-WH_EXPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการเบิกจ่ายสินค้า"
'
   ElseIf Acc = "INVENTORY-WH_TRANSFER" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการโอนสินค้า"
   ElseIf Acc = "INVENTORY-WH_TRANSFER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการโอนสินค้า"
   ElseIf Acc = "INVENTORY-WH_TRANSFER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการโอนสินค้า"
   ElseIf Acc = "INVENTORY-WH_TRANSFER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการโอนสินค้า"
   ElseIf Acc = "PACKAGE" Then
      Ri.RIGHT_ITEM_DESC = "ระบบแพ็คเกจสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_FEATURE" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_FEATURE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_FEATURE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_FEATURE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_SOC" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลแพ็คเกจสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_SOC_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลแพ็คเกจสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_SOC_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลแพ็คเกจสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_SOC_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลแพ็คเกจสินค้า/บริการ"
   ElseIf Acc = "PACKAGE_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานแพ็คเกจสินค้า/บริการ"
  ElseIf Acc = "PACKAGE-CENTER" Then
      Ri.RIGHT_ITEM_DESC = "ราคาประกาศกลาง หน้าโรงงาน"
   ElseIf Acc = "PACKAGE-CENTER_EX-WORKS-PRICE" Then
      Ri.RIGHT_ITEM_DESC = "ราคาประกาศสินค้าหน้าโรงงาน"
   ElseIf Acc = "PACKAGE-CENTER_DELIVERY-COST" Then
      Ri.RIGHT_ITEM_DESC = "ราคาประกาศค่าขนส่ง"
   ElseIf Acc = "PACKAGE-CENTER_PROMOTION-PART" Then
      Ri.RIGHT_ITEM_DESC = "ราคาโปรโมชั่นค่าสินค้า"
   ElseIf Acc = "PACKAGE-CENTER_PROMOTION-DELIVERY" Then
      Ri.RIGHT_ITEM_DESC = "ราคาโปรโมชั่นค่าขนส่ง"
   ElseIf Acc = "PRODUCT" Then
      Ri.RIGHT_ITEM_DESC = "ระบบการผลิต"
   
   ElseIf Acc = "PRODUCT_FORMULA" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลสูตรการผลิต"
   ElseIf Acc = "PRODUCT_FORMULA_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลสูตรการผลิต"
   ElseIf Acc = "PRODUCT_FORMULA_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลสูตรการผลิต"
   ElseIf Acc = "PRODUCT_FORMULA_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลสูตรการผลิต"
   
   ElseIf Acc = "PRODUCT_JOB" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลใบสั่งผลิต"
   ElseIf Acc = "PRODUCT_JOB_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลใบสั่งผลิต"
   ElseIf Acc = "PRODUCT_JOB_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลใบสั่งผลิต"
   ElseIf Acc = "PRODUCT_JOB_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลใบสั่งผลิต"
      
   ElseIf Acc = "PRODUCT_PACK" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลใบสั่งแพ็คอาหาร"
   ElseIf Acc = "PRODUCT_PACK_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลใบสั่งแพ็คอาหาร"
   ElseIf Acc = "PRODUCT_PACK_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลใบสั่งแพ็คอาหาร"
   ElseIf Acc = "PRODUCT_PACK_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลใบสั่งแพ็คอาหาร"
      
   ElseIf Acc = "PRODUCT_JOB_WAREHOUSE" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการผลิต"
   ElseIf Acc = "PRODUCT_JOB_WAREHOUSE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการผลิต"
   ElseIf Acc = "PRODUCT_JOB_WAREHOUSE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการผลิต"
   ElseIf Acc = "PRODUCT_JOB_WAREHOUSE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการผลิต"
   
   ElseIf Acc = "PRODUCT_ESTIMATE" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการปันค่าใช้จ่ายผลิต"
   ElseIf Acc = "PRODUCT_ESTIMATE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการปันค่าใช้จ่ายผลิต"
   ElseIf Acc = "PRODUCT_ESTIMATE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการปันค่าใช้จ่ายผลิต"
   ElseIf Acc = "PRODUCT_ESTIMATE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการปันค่าใช้จ่ายผลิต"
   
   ElseIf Acc = "PRODUCT_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบการผลิต"
   ElseIf Acc = "PRODUCT_PLAN" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการวางแผนการผลิต"
      
   ElseIf Acc = "LEDGER" Then
      Ri.RIGHT_ITEM_DESC = "ระบบบัญชี"
   ElseIf Acc = "LEDGER_SELL" Then
      Ri.RIGHT_ITEM_DESC = "ระบบงานขาย"
   ElseIf Acc = "LEDGER_SELL" Then
      Ri.RIGHT_ITEM_DESC = "ระบบงานขาย"
   ElseIf Acc = "LEDGER_STOCKBUY" Then
      Ri.RIGHT_ITEM_DESC = "ระบบซื้อเข้า"
   ElseIf Acc = "LEDGER_BUY" Then
      Ri.RIGHT_ITEM_DESC = "ระบบงานซื้อ(รายจ่าย)"
   ElseIf Acc = "LEDGER_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบบัญชี"
      
   ElseIf Acc = "PLANNING" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลประมาณการ"
      
   ElseIf Acc = "PLANNING_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบประมาณการ/วางแผน"
      
   
   ElseIf Acc = "COMMISSION" Then
      Ri.RIGHT_ITEM_DESC = "ระบบคอมมิตชั่น"
   ElseIf Acc = "COMMISSION_TARGET" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลเป้าการขาย"
   ElseIf Acc = "COMMISSION_TARGET_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลเป้าการขาย"
   ElseIf Acc = "COMMISSION_TARGET_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลเป้าการขาย"
   ElseIf Acc = "COMMISSION_TARGET_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลเป้าการขาย"
      
   ElseIf Acc = "COMMISSION_ORGANIZE" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลแผนภูมิพนักงาน"
   ElseIf Acc = "COMMISSION_ORGANIZE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลแผนภูมิพนักงาน"
   ElseIf Acc = "COMMISSION_ORGANIZE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลแผนภูมิพนักงาน"
   ElseIf Acc = "COMMISSION_ORGANIZE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลแผนภูมิพนักงาน"
   
   ElseIf Acc = "COMMISSION_CONDITION" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลเงื่อนไข COMMISSION"
   ElseIf Acc = "COMMISSION_CONDITION_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลเงื่อนไข COMMISSION"
   ElseIf Acc = "COMMISSION_CONDITION_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลเงื่อนไข COMMISSION"
   ElseIf Acc = "COMMISSION_CONDITION_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลเงื่อนไข COMMISSION"
   
   ElseIf Acc = "COMMISSION_COST" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลต้นทุนรายเบอร์"
   ElseIf Acc = "COMMISSION_COST_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลต้นทุนรายเบอร์"
   ElseIf Acc = "COMMISSION_COST_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลต้นทุนรายเบอร์"
   ElseIf Acc = "COMMISSION_COST_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลต้นทุนรายเบอร์"
      
   ElseIf Acc = "COMMISSION_SUBTRACT" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลส่วนหัก GP ค่าส่งเสริม"
   ElseIf Acc = "COMMISSION_SUBTRACT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลส่วนหัก GP ค่าส่งเสริม"
   ElseIf Acc = "COMMISSION_SUBTRACT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลส่วนหัก GP ค่าส่งเสริม"
   ElseIf Acc = "COMMISSION_SUBTRACT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลส่วนหัก GP ค่าส่งเสริม"
      
   ElseIf Acc = "COMMISSION_INCENTIVE" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูล INCENTIVE"
   ElseIf Acc = "COMMISSION_INCENTIVE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูล INCENTIVE รายสินค้า"
   ElseIf Acc = "COMMISSION_INCENTIVE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูล INCENTIVE รายสินค้า"
   ElseIf Acc = "COMMISSION_INCENTIVE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูล INCENTIVE รายสินค้า"
      
   ElseIf Acc = "COMMISSION_INCENTIVE-CUS-PD" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูล INCENTIVE รายลูกค้า สินค้า"
   ElseIf Acc = "COMMISSION_INCENTIVE_ADD-CUS-PD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูล INCENTIVE รายลูกค้า สินค้า"
   ElseIf Acc = "COMMISSION_INCENTIVE_EDIT-CUS-PD" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูล INCENTIVE รายลูกค้า สินค้า"
   ElseIf Acc = "COMMISSION_INCENTIVE_DELETE-CUS-PD" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูล INCENTIVE รายลูกค้า สินค้า"
      
   ElseIf Acc = "COMMISSION_CREDIT" Then
      Ri.RIGHT_ITEM_DESC = "ข้อมูลการปรับเครดิตลูกค้า"
   ElseIf Acc = "COMMISSION_CREDIT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "เพิ่มข้อมูลการปรับเครดิตลูกค้า"
   ElseIf Acc = "COMMISSION_CREDIT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "แก้ไขข้อมูลการปรับเครดิตลูกค้า"
   ElseIf Acc = "COMMISSION_CREDIT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "ลบข้อมูลการปรับเครดิตลูกค้า"
      
   ElseIf Acc = "COMMISSION_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "รายงานระบบคอมมิตชั่น"
      
   ElseIf Acc = "PROGRAM" Then
      Ri.RIGHT_ITEM_DESC = "โปรแกรม"
   
   Else
      If Len(ReportName) > 0 Then
         Ri.RIGHT_ITEM_DESC = ReportName
      Else
         Ri.RIGHT_ITEM_DESC = ""
      End If
   End If
   
End Sub

Private Function GetParentKey(Acc As String, TopFlag As Boolean) As String
Dim I As Long
Dim J As Long

   For I = 1 To Len(Acc)
      If Mid(Acc, I, 1) = "_" Then
         J = I
      End If
   Next I
   
   If J > 1 Then
      GetParentKey = Mid(Acc, 1, J - 1)
      TopFlag = False
   Else
      GetParentKey = ""
      TopFlag = True
   End If
End Function

Public Function CreatePermissionNode(Acc As String, ParentID As Long, ReportName As String) As Boolean
Dim ParentKey As String
Dim TopFlag As Boolean
Dim TempParentID As Long
Dim CreateFlag As Boolean
Dim Ri As CRightItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long
   
   'Create node here
   Set Ri = New CRightItem
   Set TempRs = New ADODB.Recordset
   TempParentID = 0
   
   Ri.RIGHT_ID = -1
   Ri.RIGHT_ITEM_NAME = Acc
   Call Ri.QueryData(1, TempRs, iCount)
   If TempRs.EOF Then
      ParentKey = GetParentKey(Acc, TopFlag)
      If Not TopFlag Then
         Call CreatePermissionNode(ParentKey, TempParentID, ReportName)
         Ri.PARENT_ID = TempParentID
      End If
      
      Ri.AddEditMode = SHOW_ADD
      Call GetParentItemDesc(Acc, Ri, ReportName)
      Call Ri.AddEditData
      ParentID = Ri.RIGHT_ID
   Else
      Call Ri.PopulateFromRS(1, TempRs)
      ParentID = Ri.RIGHT_ID
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set Ri = Nothing
End Function

Public Function VerifyAccessRight(Acc As String, Optional ReportName As String = "", Optional msgType As Long) As Boolean
Dim R As CGroupRight
Dim iCount As Long
Dim TempParentID As Long
Dim FoundFlag As Boolean

   If glbUser.REAL_USER_ID = 0 Then
      VerifyAccessRight = True
      Exit Function
   End If
   
   Call glbDaily.StartTransaction
   Call CreatePermissionNode(Acc, TempParentID, ReportName)
   Call glbDaily.CommitTransaction
   
   
   FoundFlag = False
   For Each R In glbAccessRight
       If R.RIGHT_ITEM_NAME = Acc Then
         FoundFlag = True
         If R.RIGHT_STATUS = "Y" Then
            VerifyAccessRight = True
            Exit For
         Else
            VerifyAccessRight = False
            Exit For
         End If
      End If
   Next R
   
   'If FoundFlag And (Not VerifyAccessRight) Then    เอา FoundFlag ออกเพื่อป้องกันการเพิ่ม
   If Not VerifyAccessRight Then
      VerifyAccessRight = False
      If Not msgType = 2 Then
      glbErrorLog.LocalErrorMsg = "ไม่สามารถใช้งานโปรแกรมส่วนนี้ได้เนื่องจากมีสิทธ์ไม่พอเพียง -> " & Acc
      glbErrorLog.ShowUserError
      End If
   Else
      VerifyAccessRight = True
   End If
End Function

Public Function VerifyCombo(L As Label, C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyCombo = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled And C.Visible Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyCombo = True
End Function

Public Function VerifyComboEx(C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyComboEx = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyComboEx = True
End Function

Public Function VerifyItem(C As Collection, T As Object, Idx As Long) As Boolean
Dim I As Long
Dim Count As Long

   If C.Count <= 0 Then
      VerifyItem = True
      Exit Function
   End If
   
   For I = 1 To C.Count
      If C.Item(I).CURRENT_FLAG = "Y" Then
         Count = Count + 1
      End If
   Next I
   
   If Count <> 1 Then
      Call MsgBox("กรุณาเลือกข้อมูลให้มีค่าปัจจุบัน 1 รายการ", vbOKOnly, PROJECT_NAME)
   
      T.Tabs.Item(Idx).Selected = True
      VerifyItem = False
      Exit Function
   End If
   
   VerifyItem = True
End Function

Public Sub SetTextLenType(T As TextBox, TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      T.Alignment = 1
   End If
   
   T.Tag = TT
   T.MaxLength = L
End Sub

Public Function ChangeQuote(StrQ As String) As String
   ChangeQuote = Replace(StrQ, "'", "''")
End Function
Public Function CFLAG(Value As Variant) As String
If Value = 0 Then
   CFLAG = ""
ElseIf Value = 1 Then
   CFLAG = "N"
ElseIf Value = 2 Then
   CFLAG = "Y"
End If
End Function
Public Function NVLI(Value As Variant, I As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLI = I
   Else
      NVLI = Value
   End If
End Function

Public Function NVLD(Value As Variant, I As Double) As Double
On Error Resume Next

   If IsNull(Value) Then
      NVLD = I
   Else
      NVLD = Value
   End If
End Function

Public Function NVLS(Value As Variant, S As String) As String
On Error Resume Next

   If IsNull(Value) Then
      NVLS = S
'   ElseIf IsEmpty(Value) Then
'      NVLS = S
   Else
      NVLS = Value
   End If
End Function

Public Function EmptyToString(Value As String, S As String) As String
On Error Resume Next

   If Value = "" Then
      EmptyToString = S
   Else
      EmptyToString = Value
   End If
End Function

Public Function CryptString(strInput As String, strKey As String, action As Boolean)
Dim I As Integer, C As Integer
Dim strOutput As String

If Len(strKey) Then
    For I = 1 To Len(strInput)
        C = Asc(Mid$(strInput, I, 1))
        If action Then
            C = C + Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        Else: C = C - Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        End If
        strOutput = strOutput & Chr$(C And &HFF)
    Next I
Else
    strOutput = strInput
End If
CryptString = strOutput
End Function

Public Function EncryptText(PText As String) As String
   EncryptText = CryptString(PText, "GENETICOTHELLO", True)
End Function

Public Function DecryptText(CText As String) As String
   DecryptText = CryptString(CText, "GENETICOTHELLO", False)
End Function

Public Function EnableForm(Frm As Form, En As Boolean)
   If Frm Is Nothing Then
      Exit Function
   End If
   
   Frm.Enabled = En
   If En Then
      Screen.MousePointer = vbArrow
   Else
      Frm.Refresh
      DoEvents
      Screen.MousePointer = 11
   End If
End Function

Public Function IntToThaiMonth(M As Long) As String
   If glbParameterObj Is Nothing Then
      Exit Function
   End If
   
   If M = 1 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มกราคม"
      Else
         IntToThaiMonth = "January"
      End If
   ElseIf M = 2 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กุมภาพันธ์"
      Else
         IntToThaiMonth = "February"
      End If
      
   ElseIf M = 3 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มีนาคม"
      Else
         IntToThaiMonth = "March"
      End If
      
   ElseIf M = 4 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "เมษายน"
      Else
         IntToThaiMonth = "April"
      End If
      
   ElseIf M = 5 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "พฤษภาคม"
      Else
         IntToThaiMonth = "May"
      End If
      
   ElseIf M = 6 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มิถุนายน"
      Else
         IntToThaiMonth = "June"
      End If
      
   ElseIf M = 7 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กรกฎาคม"
      Else
         IntToThaiMonth = "July"
      End If
      
   ElseIf M = 8 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "สิงหาคม"
      Else
         IntToThaiMonth = "August"
      End If
      
   ElseIf M = 9 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กันยายน"
      Else
         IntToThaiMonth = "September"
      End If
      
   ElseIf M = 10 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "ตุลาคม"
      Else
         IntToThaiMonth = "October"
      End If
      
   ElseIf M = 11 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "พฤศจิกายน"
      Else
         IntToThaiMonth = "November"
      End If
      
   ElseIf M = 12 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "ธันวาคม"
      Else
         IntToThaiMonth = "December"
      End If
   Else
      IntToThaiMonth = ""
   End If
End Function

Public Function DateToStringMonthYearExt(D As Date) As String
   If D < 0 Then
      DateToStringMonthYearExt = ""
      Exit Function
   End If
   
   DateToStringMonthYearExt = " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
End Function

Public Function DateToStringExt(D As Date) As String
   If D < 0 Then
      DateToStringExt = "-"
      Exit Function
   Else
      DateToStringExt = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
   End If
End Function
Public Function DateToStringEng(D As Date) As String
   If D < 0 Then
      DateToStringEng = "-"
      Exit Function
   Else
      DateToStringEng = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D), "0000")
   End If
End Function

Public Function DateToStringExtEx(D As Date) As String
   If D < 0 Then
      DateToStringExtEx = ""
      Exit Function
   End If
   
   DateToStringExtEx = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000") & _
                     " " & Format(HOUR(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
End Function

Public Function DateToStringIntEx2(D As Date, Minute As Long, Second As Long) As String
   DateToStringIntEx2 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & " " & _
   Format(Minute, "00") & ":" & Format(Second, "00") & ":00"
End Function

Public Function DateToStringExtEx2(D As Date) As String
   If D > 0 Then
      DateToStringExtEx2 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      DateToStringExtEx2 = ""
   End If
End Function
Public Function DateToStringExtEx4(D As Date) As String
   If D > 0 Then
      DateToStringExtEx4 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Right(Format(Year(D) + 543, "0000"), 2)
   Else
      DateToStringExtEx4 = ""
   End If
End Function

Public Function DateToStringExtEx3(D As Date) As String
   If D > 0 Then
      DateToStringExtEx3 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
      DateToStringExtEx3 = DateToStringExtEx3 & " " & Format(HOUR(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   Else
      DateToStringExtEx3 = ""
   End If
End Function

Public Function DateToStringIntEx3(D As Date) As String
   DateToStringIntEx3 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00")
End Function

Public Function ThaiDateToDate(IntDate As String) As Date
Dim YStr As String
Dim MStr As String
Dim DStr As String
Dim Y As Long
Dim M As Long
Dim D As Long

   YStr = Mid(IntDate, 7, 4)
   MStr = Mid(IntDate, 4, 2)
   DStr = Mid(IntDate, 1, 2)
               
   Y = Val(YStr) - 543
   M = Val(MStr)
   D = Val(DStr)
   
   ThaiDateToDate = DateSerial(Y, M, D)
End Function

Public Function InternalDateToStringEx4(D As String) As String
Dim T As Date
   T = InternalDateToDate(D)
   If T > 0 Then
      InternalDateToStringEx4 = Format(Day(T), "00") & "/" & Format(Month(T), "00") & "/" & Format(Year(T) + 543, "0000")
   Else
      InternalDateToStringEx4 = ""
   End If
End Function
Public Function InternalDateToStringEx5(D As Date) As String
   If D > -1 Then
      InternalDateToStringEx5 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      InternalDateToStringEx5 = ""
   End If
End Function
Public Function DateToStringInt(D As Date) As String
   If D = -1 Then
      DateToStringInt = "9999-99-99 99:99:99"
   ElseIf D = -2 Then
      DateToStringInt = "0000-00-00 00:00:00"
   Else
      DateToStringInt = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " " & Format(HOUR(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   End If
End Function
Public Function TimeToStringHHMM(D As Date) As String
   If D = -1 Then
      TimeToStringHHMM = "99:99"
   ElseIf D = -2 Then
      TimeToStringHHMM = "00:00"
   Else
      TimeToStringHHMM = Format(HOUR(D), "00") & ":" & Format(Minute(D), "00")
   End If
End Function

Public Function DateToStringIntEndMonth(D As Date) As String
   DateToStringIntEndMonth = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-31" & _
                     " 00:00:00"
End Function

Public Function DateToStringIntEx(D As Date) As String
   DateToStringIntEx = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
End Function

Public Function DateToStringIntHi(D As Date) As String
   If D > 0 Then
      DateToStringIntHi = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
   Else
      DateToStringIntHi = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   End If
End Function

Public Function DateToStringIntLow(D As Date) As String
   If D = -1 Then
      DateToStringIntLow = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   ElseIf D = -2 Then
      DateToStringIntLow = "0000" & "-" & "00" & "-" & "00" & _
                     " 00:00:00"
   Else
      DateToStringIntLow = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                        " 00:00:00"
   End If
End Function

Public Function InternalDateToDate(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDate = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDate = -2
      Exit Function
   End If
   
   If Len(IntDate) < 19 Then
      InternalDateToDate = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = Mid(IntDate, 12, 2)
   MMStr = Mid(IntDate, 15, 2)
   SSStr = Mid(IntDate, 18, 2)
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDate = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function
Public Function InternalTimeToTime(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalTimeToTime = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalTimeToTime = -2
      Exit Function
   End If
   
   If Len(IntDate) < 5 Then
      InternalTimeToTime = Format(Now, "HH:mm:ss")
      Exit Function
   End If
   
   
   HHStr = Mid(IntDate, 1, 2)
   MMStr = Mid(IntDate, 4, 2)
   SSStr = "00"
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   InternalTimeToTime = TimeSerial(HH, MM, SS)
End Function
Public Function SplitStringToDate(D As String) As Date
 Dim data(3) As String
 Dim strDate As String
 If Len(D) = 6 Then
   data(0) = Mid(D, 1, 2)
   data(1) = Mid(D, 3, 2)
   data(2) = Mid(D, 5, 2)
   strDate = data(0) & "-" & data(1) & "-" & data(2)
   SplitStringToDate = Format(strDate, "DD-MM-YYYY")
 Else
   SplitStringToDate = -1
 End If
End Function
Public Function SplitStringToDate2(D As String) As String
 Dim data(3) As String
 Dim strDate As String
 If Len(D) = 6 Then
   data(0) = Mid(D, 1, 2)
   data(1) = Mid(D, 3, 2)
   data(2) = Mid(D, 5, 2)
   strDate = data(0) & "-" & data(1) & "-" & data(2)
   SplitStringToDate2 = Format(strDate, "DD-MM-YYYY")
 Else
   SplitStringToDate2 = "-1"
 End If
End Function
Public Function SplitStringToDate3(D As String) As String
 Dim data(3) As String
 Dim strDate As String
 If Len(D) = 6 Then
   data(0) = Mid(D, 1, 2)
   data(1) = Mid(D, 3, 2)
   data(2) = Mid(D, 5, 2)
   strDate = data(2) & "-" & data(1) & "-" & data(0)
   SplitStringToDate3 = Format(strDate, "YYYY-MM")
 Else
   SplitStringToDate3 = "-1"
 End If
End Function

Public Function InternalDateToDateEx2(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If Len(IntDate) < 10 Then
      InternalDateToDateEx2 = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
      
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
      
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
      
   InternalDateToDateEx2 = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If Len(IntDate) < 8 Then
      InternalDateToDateEx = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 5, 2)
   DStr = Mid(IntDate, 7, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr) - 543
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function ReFormatDate(DStr As String) As String
Dim YYYY As String
Dim MM As String
Dim DD As String

   YYYY = Mid(DStr, 5, 4)
   MM = Mid(DStr, 3, 2)
   DD = Mid(DStr, 1, 2)
   
   ReFormatDate = YYYY & MM & DD
End Function

Public Sub InitTextBox(T As TextBox, msg As String, Optional Password As String = "")
   T.PasswordChar = Password
   T.FontSize = 12
   T.FontName = "MS Sans Serif"
   T.Text = msg
   T.BackColor = GLB_GRID_COLOR
   'T.FontBold = True
End Sub

Public Sub ShowTotalLabel(L As Label, Value As Long)
   L.Caption = "รวม = " & Value
End Sub

Public Sub ClearTreeView(ByVal tvHwnd As Long)
Dim lNodeHandle As Long

    'Turn off redrawing on the Treeview for more speed improvements
    SendMessageLong tvHwnd, WM_SETREDRAW, False, 0

    Do
        lNodeHandle = SendMessageLong(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
         If lNodeHandle > 0 Then
            SendMessageLong tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
         Else
            Exit Do
         End If
    Loop

    'Turn on redrawing on the Treeview
    SendMessageLong tvHwnd, WM_SETREDRAW, True, 0
End Sub

Public Sub InitCombo(C As ComboBox)
   C.FontSize = 12
   C.FontName = "MS Sans Serif"
   C.BackColor = GLB_GRID_COLOR
End Sub

Public Function VerifyGrid(S As String) As Boolean
   If S = "" Then
      VerifyGrid = False
      glbErrorLog.LocalErrorMsg = "กรุณาเลือกข้อมูลที่ต้องการก่อน"
      glbErrorLog.ShowUserError
   Else
      VerifyGrid = True
   End If
End Function

Public Function ConfirmDelete(S As String) As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการจะลบข้อมูล '" & S & "' ใช่หรือไม่"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmDelete = False
      Exit Function
   Else
      ConfirmDelete = True
   End If
End Function

Public Sub InitFormHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 20
   L.FontName = GLB_FONT
   L.Alignment = 2
   L.ForeColor = RGB(0, 10, 0)
End Sub

Public Sub InitDialogHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 16
   L.FontName = GLB_FONT
   L.Alignment = 2
End Sub

Public Sub InitNormalLabel(L As Label, Caption As String, Optional Color As Long = 0)
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub
Public Sub InitNormalFrame(L As SSFrame, Caption As String, Optional Color As Long = 0)
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub

Public Sub InitOption(O As OptionButton, Caption As String)
   O.Caption = Caption
   O.FontSize = 14
   O.FontBold = True
   O.FontName = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
End Sub

Public Sub InitOptionEx(O As SSOption, Caption As String)
   O.Caption = Caption
   O.Font.Size = 14
   O.Font.Bold = True
   O.Font.NAME = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
   O.BackStyle = ssTransparent
End Sub

Public Sub InitCheckBox(C As SSCheck, Caption As String)
   C.Caption = Caption
   C.FontSize = 14
   C.FontBold = True
   C.FontName = GLB_FONT
   C.BackColor = GLB_FORM_COLOR
   C.BackStyle = ssTransparent
   C.TripleState = True
End Sub

Public Sub InitMainButton(B As SSCommand, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.NAME = GLB_FONT
   B.Font3D = ssInsetLight
   B.BackColor = RGB(255, 255, 255)
   B.ButtonStyle = ssActiveBorders
   B.MousePointer = ssCustom
   B.MouseIcon = LoadPicture(glbParameterObj.ButtonCursor)
End Sub

Public Sub InitHeaderFooter(H As SSPanel, F As SSPanel)
'   H.PICTURE = LoadPicture("D:\Picture\WINPricing100\header.gif")
   If Not (F Is Nothing) Then
'      F.PICTURE = LoadPicture("D:\Picture\WINPricing100\footer.gif")
   End If
End Sub

Public Sub InitMainButtonOld(B As CommandButton, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.NAME = GLB_FONT
   B.BackColor = GLB_FORM_COLOR
End Sub

Public Sub SetSelect(T As TextBox)
   T.SelStart = 0
   T.SelLength = Len(T.Text)
End Sub

Public Sub InitDialogButton(B As CommandButton, Caption As String)
   B.Caption = Caption
   B.FontBold = True
   B.FontSize = 14
   B.FontName = GLB_FONT
   
   B.BackColor = &HFFFFFF
End Sub

Public Sub ReleaseAll()
   Call glbDatabaseMngr.DisConnectDatabase
   
   Set glbErrorLog = Nothing
   Set glbDatabaseMngr = Nothing
   Set glbParameterObj = Nothing
   Set glbUser = Nothing
   Set glbGroup = Nothing
   Set glbGuiConfigs = Nothing
   Set glbLockDate = Nothing
End Sub

Public Sub SetEnableDisableTextBox(T As TextBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableComboBox(T As ComboBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableButton(B As SSCommand, En As Boolean)
   If En Then
      B.Enabled = True
      B.BackColor = GLB_GRID_COLOR
   Else
      B.Enabled = False
      B.BackColor = &H8000000F
   End If
End Sub

Public Function ConfirmExit(HasEdit As Boolean) As Boolean
   If Not HasEdit Then
      ConfirmExit = True
   Else
      glbErrorLog.LocalErrorMsg = "ท่านต้องการจะออกจากโปรแกรมโดยไม่มีการบันทึกข้อมูลใช่หรือไม่"
      If glbErrorLog.AskMessage = vbYes Then
         ConfirmExit = True
      Else
         ConfirmExit = False
      End If
   End If
End Function
Public Function ConfirmSave() As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการการบันทึกข้อมูลใช่หรือไม่"
   If glbErrorLog.AskMessage = vbYes Then
      ConfirmSave = True
   Else
      ConfirmSave = False
   End If
End Function

Public Function ThaiBaht(ByVal pamt As Double) As String
Dim valstr As String, vLen As Integer, vno As Integer, syslge As String
Dim I As Integer, J As Integer, v As Integer
Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
Dim vword(20) As String

 If pamt <= 0# Then
   ThaiBaht = ""
   Exit Function
 End If
 valstr = Trim(Format$(pamt, "##########0.00"))
 vLen = Len(valstr) - 3
 For I = 1 To 20
     vword(I) = ""
 Next I
wnumber(1) = "หนึ่ง": wnumber(2) = "สอง": wnumber(3) = "สาม": wnumber(4) = "สี่"
wnumber(5) = "ห้า": wnumber(6) = "หก": wnumber(7) = "เจ็ด": wnumber(8) = "แปด"
wnumber(9) = "เก้า": wdigit(1) = "บาท": wdigit(2) = "สิบ": wdigit(3) = "ร้อย": wdigit(4) = "พัน"
wdigit(5) = "หมื่น": wdigit(6) = "แสน": wdigit(7) = "ล้าน": spcdg(1) = "สตางค์": spcdg(2) = "เอ็ด"
spcdg(3) = "ยี่": spcdg(4) = "ถ้วน"
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
        vword(I) = ""
        If (vLen - I + 1) = 7 Then
            vword(I) = wdigit(7)             '--ล้าน
        End If
    Else
        If (vLen - I + 1) > 7 Then
            J = vLen - I - 5               '--เกินหลักล้าน
        Else
            J = vLen - I + 1               '--หลักแสน
        End If
        vword(I) = wnumber(vno) + wdigit(J) '-30ถึง90
        If vno = 1 And J = 2 Then
            vword(I) = wdigit(2)             '--สิบ
        End If
        If vno = 2 And J = 2 Then
            vword(I) = spcdg(3) + wdigit(J)  '--ยี่สิบ
        End If
        If J = 1 Then                       ' สิยเอ็ค -->เก้าสิบเอ็ด
            vword(I) = wnumber(vno)
            If vno = 1 And vLen > 1 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2)
                End If
            End If
        End If
        If J = 7 Then         '-แก้บักกรณี 11,111,111.00 สิบเอ็ด
            vword(I) = wnumber(vno) + wdigit(J)   '-ล้าน
            If vno = 1 And vLen > 7 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2) + wdigit(J)
                End If
            End If
        End If
    End If
Next I
    
If Int(pamt) > 0 Then
       vword(vLen) = vword(vLen) + wdigit(1)
End If
 '--------------ทศนิยม --------------
valstr = Mid$(valstr, vLen + 2, 2)
vLen = Len(valstr)
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
           vword(I + 10) = ""
    Else
           J = vLen - I + 1
           vword(I + 10) = wnumber(vno) + wdigit(J)
        If vno = 1 And J = 2 Then
              vword(I + 10) = wdigit(2)
        End If
        If vno = 2 And J = 2 Then
              vword(I + 10) = spcdg(3) + wdigit(J)
        End If
        If J = 1 Then
            If vno = 1 And Int(Val(Mid$(valstr, I - 1, 1))) <> 0 Then
                 vword(I + 10) = spcdg(2)
            Else
                 vword(I + 10) = wnumber(vno)
            End If
        End If
    End If
Next I
If pamt <> 0 Then
    If Val(valstr) = 0 Then
        vword(13) = spcdg(4)
    Else
        vword(13) = spcdg(1)
    End If
End If

 '*** เผื่อใช้กรณียาวมาก และต้องการตัดประโยค
 valstr = ""
 For I = 1 To 20
    'IF LEN(valstr) < 70 AND LEN(valstr + vword(i)) > 70 Then
    '   valstr = valstr + REPLICATE(" ",70 - LEN(valstr))
    'END IF
    valstr = valstr + vword(I)
 Next I
 'valstr='('+valstr+')'
 ThaiBaht = (valstr)
End Function

Public Function WildCard(WStr As String, SubLen As Long, NewStr As String) As Boolean
Dim Tmp As String
   Tmp = Trim(WStr)
   If Tmp = "" Then
      WildCard = False
      Exit Function
   End If
   
   If Mid(Tmp, Len(Tmp)) = "%" Then
      SubLen = Len(Tmp) - 1
      NewStr = Mid(Tmp, 1, SubLen)
      
      WildCard = True
   Else
      WildCard = False
   End If
End Function

Public Function FormatString(S As String, Patch As String, L As Long) As String
Dim Temp As String
Dim Start As Long
Dim I As Long
Dim J As Long

   Temp = Space(L)
   Call Replace(Temp, " ", Patch)
   J = 0
   Start = (L - Len(S)) \ 2
   
   For I = 1 To L
      If I < Start Then
         Mid(Temp, I) = Patch
      Else
         If I > Start + Len(S) Then
            Mid(Temp, I) = Patch
         Else
            J = J + 1
            Mid(Temp, I) = Mid(S, J)
         End If
      End If
   Next I
   
   FormatString = Temp
End Function

Public Function FormatNumber(N As Variant, Optional DecimalPoint As Long = 2, Optional ZeroString = "") As String
Dim T As Double
Dim TempStr As String
Dim I As Long

   If DecimalPoint = -1 Then
    Dim Spt() As String
      If IsNull(N) Then
         T = 0
      Else
         
         T = N
         Spt = Split(T, ".")
         If UBound(Spt) > 0 Then
            DecimalPoint = Len(Spt(1))
         Else
            DecimalPoint = 0
         End If
      End If
   End If

   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      If ZeroString = "" Then
         FormatNumber = "0" & TempStr
      Else
         FormatNumber = ZeroString
      End If
   Else
      FormatNumber = Format(T, "#,##0" & TempStr)
   End If
End Function

Public Function FormatNumberEx(N As Variant) As String
Dim T As Double

   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      FormatNumberEx = "0.00"
   Else
      FormatNumberEx = T
   End If
End Function
Public Function FormatNumberToNull(N As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "") As String
Dim T As Double
Dim TempStr As String
Dim I As Long
  If DecimalPoint = -1 Then
    Dim Spt() As String
      If IsNull(N) Then
         T = 0
      Else
         
         T = N
         Spt = Split(T, ".")
         If UBound(Spt) > 0 Then
            DecimalPoint = Len(Spt(1))
         Else
            DecimalPoint = 0
         End If
      End If
   End If
   
   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumberToNull = ZeroString & TempStr
      Else
         FormatNumberToNull = ZeroString
      End If
   ElseIf Quat Then
      FormatNumberToNull = Format(T, "#,##0" & TempStr)
   Else
      FormatNumberToNull = Format(T, "0" & TempStr)
   End If
End Function
Public Function FormatNumberToNullMinus(N As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "") As String
Dim T As Double
Dim TempStr As String
Dim I As Long
Dim CheckMinusFlag As Boolean
   
   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T <= 0 Then
      CheckMinusFlag = True
      T = Abs(T)
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumberToNullMinus = ZeroString & TempStr
      Else
         FormatNumberToNullMinus = ZeroString
      End If
   ElseIf Quat Then
      FormatNumberToNullMinus = Format(T, "#,##0" & TempStr)
   Else
      FormatNumberToNullMinus = Format(T, "0" & TempStr)
   End If
   If CheckMinusFlag Then
      FormatNumberToNullMinus = "(" & FormatNumberToNullMinus & ")"
   End If
End Function

Public Function ReverseFormatNumber(N As String) As Double
   ReverseFormatNumber = Val(Replace(N, ",", ""))
End Function

Public Function IDToListIndex(Cbo As ComboBox, id As Long) As Long
Dim I As Long
Dim Temp As String

   IDToListIndex = -1
   For I = 0 To Cbo.ListCount - 1
      If InStr(Cbo.ItemData(I), ":") <= 0 Then
         Temp = Cbo.ItemData(I)
      Else
         Temp = Mid(Cbo.ItemData(I), 1, InStr(Cbo.ItemData(I), ":") - 1)
      End If
      If Temp = id Then
         IDToListIndex = I
      End If
   Next I
End Function

Public Sub Main()
On Error GoTo ErrorHandler
Dim I As Long
Dim TempDB As String

   GLB_GRID_COLOR = RGB(255, 255, 250)
   GLB_NORMAL_COLOR = RGB(0, 0, 0)
   GLB_ALERT_COLOR = RGB(255, 0, 0)
   GLB_FORM_COLOR = RGB(180, 200, 200)
   GLB_HEAD_COLOR = GLB_FORM_COLOR
   GLB_GRIDHD_COLOR = RGB(149, 194, 240)
   GLB_SHOW_COLOR = RGB(0, 0, 240)
   GLB_MANDATORY_COLOR = RGB(0, 0, 255)

   Set glbSetting = New clsGlobalSetting
   Set glbParameterObj = New clsParameter
   Set glbUser = New CUser
   Set glbGroup = New CGroup
   Set glbSystemParams = New Collection


   Set glbErrorLog = New clsErrorLog
   glbErrorLog.DayKeepLog = 10
   glbErrorLog.LogFileMode = LOG_CURRENT_DATE
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Main"
   glbErrorLog.MsgBoxTitle = PROJECT_NAME
   
   If App.PrevInstance = True Then
      glbErrorLog.LocalErrorMsg = "โปรแกรมเดิมได้ถูกรันก่อนหน้านี้แล้ว"
      glbErrorLog.ShowUserError

      Set glbErrorLog = Nothing
      Exit Sub
   End If
   
   Load frmSplash
   frmSplash.Show 0
   frmSplash.Refresh
   
   Set glbGuiConfigs = New CGuiConfigs
   Call glbGuiConfigs.CreateGuiConfig(glbParameterObj.Programowner)
   
   If Command = "1" Or Command = "" Then
      TempDB = glbParameterObj.DBFile
   ElseIf Command = "2" Then
      TempDB = glbParameterObj.DBFileAP
   Else
      TempDB = glbParameterObj.DBFileAPX
   End If
   
'   Set glbDatabaseMngr = New clsDatabaseMngr
'   If Not glbDatabaseMngr.ConnectLegacyDatabase2(glbParameterObj.DBConfigFile, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
'      '''Debug.Print "Error"
'   End If

   
   
   
   Set glbDatabaseMngr = New clsDatabaseMngr
'   TempDB = "D:\Database\GDB\WINPRICING400_MM.GDB"
'   glbParameterObj.UserName = "SYSDBA"
'   glbParameterObj.Password = "masterkey"
   If Not glbDatabaseMngr.ConnectDatabase(TempDB, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
      frmDBSetting.UserName = glbParameterObj.UserName
      frmDBSetting.Password = glbParameterObj.Password
      frmDBSetting.FileDb = glbParameterObj.DBFile
      frmDBSetting.Header = " ไม่สามารถเชื่อต่อฐานข้อมูลได้ "

      Load frmDBSetting
      frmDBSetting.Show 1
      If frmDBSetting.OKClick Then
         glbParameterObj.UserName = frmDBSetting.UserName
         glbParameterObj.Password = frmDBSetting.Password
      
         If Command = "" Or Command = "1" Then
            glbParameterObj.DBFile = frmDBSetting.FileDb
         ElseIf Command = "2" Then
            glbParameterObj.DBFileAP = frmDBSetting.FileDb
         Else
            glbParameterObj.DBFileAPX = frmDBSetting.FileDb
         End If
         
      Else
         Unload frmDBSetting
         Set frmDBSetting = Nothing

         Unload frmSplash
         Set frmSplash = Nothing

         Call ReleaseAll
         End
      End If
      Unload frmDBSetting
      Set frmDBSetting = Nothing
   End If
      
   If Not glbDatabaseMngr.ConnectAgentServer(glbParameterObj.LicenseIP, glbParameterObj.LicensePort, glbErrorLog) Then
      frmAgentSetting.Port = glbParameterObj.LicensePort
      frmAgentSetting.IP = glbParameterObj.LicenseIP
      frmAgentSetting.Header = " ไม่สามารถเชื่อมต่อกับไลเซนส์เซิร์ฟเวอร์ได้ "

      Load frmAgentSetting
      frmAgentSetting.Show 1

      If frmAgentSetting.OKClick Then
         glbParameterObj.LicenseIP = frmAgentSetting.IP
         glbParameterObj.LicensePort = frmAgentSetting.Port
      Else
         Unload frmAgentSetting
         Set frmAgentSetting = Nothing

         Unload frmSplash
         Set frmSplash = Nothing

         Call ReleaseAll
         End
      End If
      Unload frmAgentSetting
      Set frmAgentSetting = Nothing
   End If
   Unload frmSplash
   Set frmSplash = Nothing
   
   Set glbDaily = New clsDaily
   Set glbAdmin = New clsAdmin
   Set glbMaster = New clsMaster
   Set glbLegacy = New clsLegacy
   Set glbLoginTracking = New CLoginTracking
   Set glbEnterPrise = New CEnterprise
   Set glbAccessRight = New Collection
   Set glbProduction = New clsProduction
'   Set glbProductionWH = New clsProductionWH
   Set glbPlanning = New clsPlanning
   Set glbTargetCom = New clsTargetCom
   Set glbInventoryAct = New clsInventoryAct
   Set glbInventoryWhAct = New clsInventoryWhAct
   Set glbAuthenPO = New clsAuthenPO
   Set glbLockDate = New CLockDate
   glbLoginID = 1
   
   
   Call LoadSystemParam(Nothing, glbSystemParams)
   
   Load frmWinPricingMain
   frmWinPricingMain.Show

   Exit Sub
   
ErrorHandler:
   If glbErrorLog Is Nothing Then
      MsgBox Err.DESCRIPTION
   Else
      glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   End If
End Sub

Public Sub InitOrderType(C As ComboBox)
   C.Clear
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("น้อยไปมาก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("มากไปน้อย"))
   C.ItemData(2) = 2
End Sub
Public Sub InitUnitType(C As ComboBox)
   C.Clear
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ถุง"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ตัน"))
   C.ItemData(2) = 2
End Sub

Public Function GetItem(Col As Collection, Idx As Long, RealIndex As Long) As Object
Dim I As Long
Dim Count As Long

   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
      If Count = Idx Then
         RealIndex = I
         Set GetItem = Col.Item(I)
         Exit Function
      End If
   Next I
   
   Set GetItem = Nothing
End Function

Public Function CountItem(Col As Collection) As Long
Dim I As Long
Dim Count As Long

   Count = 0
   
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
   Next I
   
   CountItem = Count
End Function

Public Function VSP_CalTable(ByVal pRaw As String, ByVal pWidth As Long, ByRef pPer() As Long) As String
On Error GoTo ErrorHandler
Dim strTemp As String
Dim I As Long
Dim Count As Long
Dim iPer As Long
Dim tPer As Long
Dim Total As Long
Dim Prefix() As String
Dim Value() As Long
Dim iTemp As Long
   
   pRaw = Trim$(pRaw)
   If Len(pRaw) <= 0 Then
      VSP_CalTable = ""
      Exit Function
   End If
   Count = 0
   iPer = 1
   Total = 0
   strTemp = ""
   While iPer <= Len(pRaw)
      If Val(Mid$(pRaw, iPer, 1)) <= 0 Then
         strTemp = strTemp & Mid$(pRaw, iPer, 1)
      Else
         Count = Count + 1
         ReDim Preserve Prefix(Count)
         ReDim Preserve Value(Count)
         Prefix(Count) = strTemp
         tPer = InStr(iPer, pRaw, "|")
         If tPer <= 0 Then tPer = InStr(iPer, pRaw, ";")

         Value(Count) = Val(Mid$(pRaw, iPer, tPer - iPer))
         Total = Total + Value(Count)
         iPer = tPer
         strTemp = ""
      End If
      iPer = iPer + 1
   Wend
   strTemp = ""
   ReDim pPer(Count)
   For I = 1 To Count - 1
      iTemp = CLng((Value(I) * pWidth) / Total)
      strTemp = strTemp & Trim$(Prefix(I)) & Trim$(str$(iTemp)) & "|"
      If I = 1 Then
         pPer(I - 1) = iTemp
      Else
         pPer(I - 1) = pPer(I - 2) + iTemp
      End If
   Next I
   strTemp = strTemp & Trim$(Prefix(I)) & CLng(((Value(I) * pWidth) / Total)) & ";"
   If I > 1 Then
      iTemp = CLng((Value(I) * pWidth) / Total)
      pPer(I - 1) = pPer(I - 2) + iTemp
   End If
   VSP_CalTable = strTemp

   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function

Public Function Check2Flag(A As Long) As String
   If A = ssCBChecked Then
      Check2Flag = "Y"
   Else
      Check2Flag = "N"
   End If
End Function
Public Function Option2Flag(A As Boolean) As String
   If A Then
      Option2Flag = "Y"
   Else
      Option2Flag = "N"
   End If
End Function
Public Function TypeUnitFlag(A As Long) As String
   If 1 Then
      TypeUnitFlag = "ถุง"
   ElseIf 2 Then
      TypeUnitFlag = "ตัน"
   Else
      TypeUnitFlag = ""
   End If
End Function
Public Function CheckUniqueNs(UnqType As UNIQUE_TYPE, Key As String, id As Long, Optional FieldNameExTendValue As String, Optional FlagSpacial As Long = 0, Optional FieldNameExTendValue2 As String, Optional ID_TYPE As Long = 0, Optional ID_TEMP As Long = 0) As Boolean
On Error GoTo ErrorHandler
Dim TableName As String
Dim FieldName1 As String
Dim FieldName2 As String
Dim Flag As Boolean
Dim Count As Long
Dim FieldNameExTend As String
Dim FieldNameExTend2 As String
   
   CheckUniqueNs = False
   
   Flag = False
   If UnqType = TEACHER_UNIQUE Then
      TableName = "TEACHER"
      FieldName1 = "TEACHER_CODE"
      FieldName2 = "TEACHER_ID"
      Flag = True
   ElseIf UnqType = USERGROUP_UNIQUE Then
      TableName = "USER_GROUP"
      FieldName1 = "GROUP_NAME"
      FieldName2 = "GROUP_ID"
      Flag = True
   ElseIf UnqType = SUBJECT_UNIQUE Then
      TableName = "SUBJECT"
      FieldName1 = "SUBJECT_CODE"
      FieldName2 = "SUBJECT_ID"
      Flag = True
   ElseIf UnqType = PRDFEATURE_UNIQUE Then
      TableName = "PRDFEATURE_NAME"
      FieldName1 = "PRODUCT_CODE"
      FieldName2 = "PRDFEATURE_NAME_ID"
      Flag = True
   ElseIf UnqType = FACULTY_UNIQUE Then
      TableName = "FACULTY"
      FieldName1 = "FACULTY_CODE"
      FieldName2 = "FACULTY_ID"
      Flag = True
   ElseIf UnqType = DBN_UNIQUE Then
      TableName = "BILL"
      FieldName1 = "BILL_NO"
      FieldName2 = "BILL_ID"
      Flag = True
   ElseIf UnqType = FEATURENO_UNIQUE Then
      TableName = "FEATURE"
      FieldName1 = "FEATURE_CODE"
      FieldName2 = "FEATURE_ID"
      Flag = True
   ElseIf UnqType = SOCNO_UNIQUE Then
      TableName = "SOC"
      FieldName1 = "SOC_CODE"
      FieldName2 = "SOC_ID"
      Flag = True
   ElseIf UnqType = EMPCODE_UNIQUE Then
      TableName = "EMPLOYEE"
      FieldName1 = "EMP_CODE"
      FieldName2 = "EMP_ID"
      Flag = True
   ElseIf UnqType = USERNAME_UNIQUE Then
      TableName = "USER_ACCOUNT"
      FieldName1 = "USER_NAME"
      FieldName2 = "USER_ID"
      Flag = True
   ElseIf UnqType = REPAIR_UNIQUE Then
      TableName = "REPAIR_DATA"
      FieldName1 = "REPAIR_NUM"
      FieldName2 = "REPAIR_ID"
      Flag = True
   ElseIf UnqType = IMPORT_UNIQUE Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = EXPORT_UNIQUE Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = REPAIR_FORMULA_UNIQUE Then
      TableName = "REPAIR_FORMULA"
      FieldName1 = "FORMULA_CODE"
      FieldName2 = "FORMULA_ID"
      Flag = True
   ElseIf UnqType = SUPPLIER_UNIQUE Then
      TableName = "SUPPLIER"
      FieldName1 = "SUPPLIER_CODE"
      FieldName2 = "SUPPLIER_ID"
      Flag = True
   ElseIf UnqType = FREELANCE_UNIQUE Then
      TableName = "FREELANCE"
      FieldName1 = "FREELANCE_CODE"
      FieldName2 = "FREELANCE_ID"
      Flag = True
   ElseIf UnqType = PARTNO_UNIQUE Then
      TableName = "PART_ITEM"
      FieldName1 = "PART_NO"
      FieldName2 = "PART_ITEM_ID"
      Flag = True
   ElseIf UnqType = QUOATATION_UNIQUE Then
      TableName = "QUOATATION"
      FieldName1 = "QUOATATION_NO"
      FieldName2 = "QUOATATION_ID"
      Flag = True
   ElseIf UnqType = EXPENSE_UNIQUE Then
      TableName = "EXPENSE_GROUP"
      FieldName1 = "GROUP_NO"
      FieldName2 = "EXPENSE_GROUP_ID"
      Flag = True
   ElseIf UnqType = REVENUE_UNIQUE Then
      TableName = "REVENUE_GROUP"
      FieldName1 = "GROUP_NO"
      FieldName2 = "REVENUE_GROUP_ID"
      Flag = True
   ElseIf UnqType = PO_UNIQUE Then
      TableName = "PURCHASE_ORDER"
      FieldName1 = "PO_NO"
      FieldName2 = "PO_ID"
      Flag = True
   ElseIf UnqType = CUSTOMER_UNIQUE Then
      TableName = "PATIENT"
      FieldName1 = "PATIENT_CODE"
      FieldName2 = "PATIENT_ID"
      Flag = True
   ElseIf UnqType = BORROW_UNIQUE Then
      TableName = "EMP_RECEIVABLE"
      FieldName1 = "BORROW_NO"
      FieldName2 = "EMP_RECEIVABLE_ID"
      Flag = True
   ElseIf UnqType = TRUCK_UNIQUE Then
      TableName = "RESOURCE"
      FieldName1 = "RESOURCE_NO"
      FieldName2 = "RESOURCE_ID"
      Flag = True
   ElseIf UnqType = JOBPLAN_UNIQUE Then
      TableName = "JOB_PLAN"
      FieldName1 = "PLAN_NO"
      FieldName2 = "JOB_PLAN_ID"
      Flag = True
   ElseIf UnqType = PARTTYPE_NO Then
      TableName = "PART_TYPE"
      FieldName1 = "PART_TYPE_NO"
      FieldName2 = "PART_TYPE_ID"
      Flag = True
   ElseIf UnqType = PARTTYPE_NAME Then
      TableName = "PART_TYPE"
      FieldName1 = "PART_TYPE_NAME"
      FieldName2 = "PART_TYPE_ID"
      Flag = True
   ElseIf UnqType = LOCATION_NO Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NO"
      FieldName2 = "LOCATION_ID"
      Flag = True
   ElseIf UnqType = LOCATION_NAME Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NAME"
      FieldName2 = "LOCATION_ID"
      Flag = True
   ElseIf UnqType = PRODUCTTYPE_NO Then
      TableName = "PRODUCT_TYPE"
      FieldName1 = "PRODUCT_TYPE_NO"
      FieldName2 = "PRODUCT_TYPE_ID"
      Flag = True
   ElseIf UnqType = PRODUCTTYPE_NAME Then
      TableName = "PRODUCT_TYPE"
      FieldName1 = "PRODUCT_TYPE_NAME"
      FieldName2 = "PRODUCT_TYPE_ID"
      Flag = True
   ElseIf UnqType = PRODUCTSTATUS_NO Then
      TableName = "PRODUCT_STATUS"
      FieldName1 = "PRODUCT_STATUS_NO"
      FieldName2 = "PRODUCT_STATUS_ID"
      Flag = True
   ElseIf UnqType = PRODUCTSTATUS_NAME Then
      TableName = "PRODUCT_STATUS"
      FieldName1 = "PRODUCT_STATUS_NAME"
      FieldName2 = "PRODUCT_STATUS_ID"
      Flag = True
   ElseIf UnqType = HOUSE_NO Then
      TableName = "HOUSE"
      FieldName1 = "HOUSE_NO"
      FieldName2 = "HOUSE_ID"
      Flag = True
   ElseIf UnqType = HOUSE_NAME Then
      TableName = "HOUSE"
      FieldName1 = "HOUSE_NAME"
      FieldName2 = "HOUSE_ID"
      Flag = True
   ElseIf UnqType = COUNTRY_NO Then
      TableName = "COUNTRY"
      FieldName1 = "COUNTRY_NO"
      FieldName2 = "COUNTRY_ID"
      Flag = True
   ElseIf UnqType = COUNTRY_NAME Then
      TableName = "COUNTRY"
      FieldName1 = "COUNTRY_NAME"
      FieldName2 = "COUNTRY_ID"
      Flag = True
   ElseIf UnqType = CSTGRADE_NO Then
      TableName = "CUSTOMER_GRADE"
      FieldName1 = "CSTGRADE_NO"
      FieldName2 = "CSTGRADE_ID"
      Flag = True
   ElseIf UnqType = CSTGRADE_NAME Then
      TableName = "CUSTOMER_GRADE"
      FieldName1 = "CSTGRADE_NAME"
      FieldName2 = "CSTGRADE_ID"
      Flag = True
   ElseIf UnqType = CSTTYPE_NO Then
      TableName = "CUSTOMER_TYPE"
      FieldName1 = "CSTTYPE_NO"
      FieldName2 = "CSTTYPE_ID"
      Flag = True
   ElseIf UnqType = CSTTYPE_NAME Then
      TableName = "CUSTOMER_TYPE"
      FieldName1 = "CSTTYPE_NAME"
      FieldName2 = "CSTTYPE_ID"
      Flag = True
   ElseIf UnqType = CUSTCODE_UNIQUE Then
      TableName = "CUSTOMER"
      FieldName1 = "CUSTOMER_CODE"
      FieldName2 = "CUSTOMER_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERGRADE_NO Then
      TableName = "SUPPLIER_GRADE"
      FieldName1 = "SUPPLIER_GRADE_NO"
      FieldName2 = "SUPPLIER_GRADE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERGRADE_NAME Then
      TableName = "SUPPLIER_GRADE"
      FieldName1 = "SUPPLIER_GRADE_NAME"
      FieldName2 = "SUPPLIER_GRADE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERTYPE_NO Then
      TableName = "SUPPLIER_TYPE"
      FieldName1 = "SUPPLIER_TYPE_NO"
      FieldName2 = "SUPPLIER_TYPE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERYPE_NAME Then
      TableName = "SUPPLIER_TYPE"
      FieldName1 = "SUPPLIER_TYPE_NAME"
      FieldName2 = "SUPPLIER_TYPE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERSTATUS_NO Then
      TableName = "SUPPLIER_STATUS"
      FieldName1 = "SUPPLIER_STATUS_NO"
      FieldName2 = "SUPPLIER_STATUS_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERSTATUS_NAME Then
      TableName = "SUPPLIER_STATUS"
      FieldName1 = "SUPPLIER_STATUS_NAME"
      FieldName2 = "SUPPLIER_STATUS_ID"
      Flag = True
   ElseIf UnqType = POSITION_NO Then
      TableName = "EMP_POSITION"
      FieldName1 = "POSITION_NAME"
      FieldName2 = "POSITION_ID"
      Flag = True
   ElseIf UnqType = UNIT_NO Then
      TableName = "UNIT"
      FieldName1 = "UNIT_NO"
      FieldName2 = "UNIT_ID"
      Flag = True
   ElseIf UnqType = UNIT_NAME Then
      TableName = "UNIT"
      FieldName1 = "UNIT_NAME"
      FieldName2 = "UNIT_ID"
      Flag = True
   ElseIf UnqType = YEAR_NO Then
      TableName = "YEAR_SEQ"
      FieldName1 = "YEAR_NO"
      FieldName2 = "YEAR_SEQ_ID"
      Flag = True
   ElseIf UnqType = PARTGROUP_NO Then
      TableName = "PART_GROUP"
      FieldName1 = "PART_GROUP_NO"
      FieldName2 = "PART_GROUP_ID"
      Flag = True
   ElseIf UnqType = PARTGROUP_NAME Then
      TableName = "PART_GROUP"
      FieldName1 = "PART_GROUP_NAME"
      FieldName2 = "PART_GROUP_ID"
      Flag = True
   ElseIf UnqType = DO_PLAN_UNIQUE Then
      TableName = "BILLING_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "BILLING_DOC_ID"
      Flag = True
ElseIf UnqType = PERSON_CODE Then
      TableName = "EMPLOYEE"
      FieldName1 = "EMP_CODE"
      FieldName2 = "EMP_ID"
      Flag = True
ElseIf UnqType = RELIGIOUS_NO Then
      TableName = "RELIGIOUS_DATA"
      FieldName1 = "RELIGIOUS_NO"
      FieldName2 = "RELIGIOUS_ID"
      Flag = True
ElseIf UnqType = RELIGIOUS_NAME Then
      TableName = "RELIGIOUS_DATA"
      FieldName1 = "RELIGIOUS_NAME"
      FieldName2 = "RELIGIOUS_ID"
      Flag = True
ElseIf UnqType = WORK_NO Then
      TableName = "WORK_STATUS"
      FieldName1 = "WORK_NO"
      FieldName2 = "WORK_ID"
      Flag = True
ElseIf UnqType = WORK_NAME Then
      TableName = "WORK_STATUS"
      FieldName1 = "WORK_NAME"
      FieldName2 = "WORK_ID"
      Flag = True
ElseIf UnqType = RESIGN_NO Then
      TableName = "RESIGN_REASON"
      FieldName1 = "RSGRESON_NO"
      FieldName2 = "RSGRESON_ID"
      Flag = True
ElseIf UnqType = RESIGN_NAME Then
      TableName = "RESIGN_REASON"
      FieldName1 = "RSGRESON_NAME"
      FieldName2 = "RSGRESON_ID"
      Flag = True
   ElseIf UnqType = BANK_NO Then
      TableName = "BANK_ACCOUNT"
      FieldName1 = "BANK_NO"
      FieldName2 = "BANK_ID"
      Flag = True
ElseIf UnqType = BANK_NAME Then
      TableName = "BANK_ACCOUNT"
      FieldName1 = "BANK_NAME"
      FieldName2 = "BANK_ID"
      Flag = True
   ElseIf UnqType = DOC_NO Then
      TableName = "DOCUMENT_TYPE"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "DOCTYPE_ID"
      Flag = True
ElseIf UnqType = DOC_NAME Then
      TableName = "DOCUMENT_TYPE"
      FieldName1 = "DOCTYPE_NAME"
      FieldName2 = "DOCTYPE_ID"
      Flag = True
   ElseIf UnqType = MONTHLY_ADD_NO Then
      TableName = "MONTHLY_ADD"
      FieldName1 = "MONTHLY_ADD_NO"
      FieldName2 = "MONTHLY_ADD_ID"
      Flag = True
ElseIf UnqType = MONTHLY_ADD_NAME Then
      TableName = "MONTHLY_ADD"
      FieldName1 = "MONTHLY_ADD_NAME"
      FieldName2 = "MONTHLY_ADD_ID"
      Flag = True
   ElseIf UnqType = MONTHLY_SUB_NO Then
      TableName = "MONTHLY_SUB"
      FieldName1 = "MONTHLY_SUB_NO"
      FieldName2 = "MONTHLY_SUB_ID"
      Flag = True
   ElseIf UnqType = MONTHLY_SUB_NAME Then
      TableName = "MONTHLY_SUB"
      FieldName1 = "MONTHLY_SUB_NAME"
      FieldName2 = "MONTHLY_SUB_ID"
      Flag = True
   ElseIf UnqType = JOB_NO Then
      TableName = "JOB"
      FieldName1 = "JOB_NO"
      FieldName2 = "JOB_ID"
      Flag = True
   ElseIf UnqType = CHEQUEDOC_UNIQUE Then
      TableName = "CHEQUE_DOC"
      FieldName1 = "CHEQUE_DOC_NO"
      FieldName2 = "CHEQUE_DOC_ID"
      Flag = True
   ElseIf UnqType = FORMULA_NO Then
      TableName = "FORMULA"
      FieldName1 = "FORMULA_NO"
      FieldName2 = "FORMULA_ID"
      Flag = True
   ElseIf UnqType = CURRENCY_NO Then
      TableName = "CURRENCY"
      FieldName1 = "CURRENCY_EXC_NO"
      FieldName2 = "CURRENCY_EXC_ID"
      Flag = True
   ElseIf UnqType = PLANNING_UNIQUE Then
      TableName = "PLANNING"
      FieldName1 = "PLANNING_FROM"
      FieldName2 = "PLANNING_ID"
      Flag = True
      FieldNameExTend = "PLANNING_AREA"
   ElseIf UnqType = INVENTORY_ACT_UNIQUE Then
      TableName = "INVENTORY_ACT"
      FieldName1 = "INVENTORY_ACT_DATE"
      FieldName2 = "INVENTORY_ACT_ID"
      Flag = True
      FieldNameExTend = "INVENTORY_ACT_AREA"
   ElseIf UnqType = TARGET_UNIQUE Then
      TableName = "TARGET"
      FieldName1 = "YEAR_NO"
      FieldName2 = "TARGET_ID"
      Flag = True
   ElseIf UnqType = MASTER_VALID_NO Then
      TableName = "MASTER_VALID"
      FieldName1 = "MASTER_VALID_NO"
      FieldName2 = "MASTER_VALID_ID"
      FieldNameExTend = "MASTER_VALID_TYPE"
      Flag = True
   ElseIf UnqType = MASTER_COMMISSION_CREDIT Then
      TableName = "COMMISSION_CREDIT"
      FieldName1 = "CUSTOMER_ID"
      FieldName2 = "COMMISSION_CREDIT_ID"
      Flag = True
   ElseIf UnqType = LOT_DOC Then
      TableName = "LOT_DOC"
      FieldName1 = "LOT_ID"
      FieldName2 = "LOT_ITEM_WH_ID"
      Flag = True
   ElseIf UnqType = PALLET_DOC Then
      TableName = "PALLET_DOC"
      FieldName1 = "PALLET_DOC_NO"
      FieldName2 = "LOT_DOC_ID"
      Flag = True
   ElseIf UnqType = INVENTORY_WH_DOC_UNIQUE Then
      TableName = "INVENTORY_WH_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_WH_DOC_ID"
      Flag = True
   ElseIf UnqType = PACK_PRODUCTION_UNIQUE Then
      TableName = "PACK_PRODUCTION"
      FieldName1 = "PACK_PRODUCTION_DATE"
      FieldName2 = "PACK_PRODUCTION_ID"
      Flag = True
   ElseIf UnqType = PACK_PRODUCTION_UNIQUE Then
      TableName = "PACK_PRODUCTION"
      FieldName1 = "PACK_PRODUCTION_DATE"
      FieldName2 = "PACK_PRODUCTION_ID"
      Flag = True
   ElseIf UnqType = LOT_UNIQUE Then
      TableName = "LOT"
      FieldName1 = "LOT_NO"
      FieldName2 = "LOT_ID"
      Flag = True
   ElseIf UnqType = INVENTORY_WH_DOC_DATE_UNIQUE Then
      TableName = "INVENTORY_WH_DOC"
      FieldName1 = "DOCUMENT_DATE"
      FieldName2 = "DOCUMENT_TYPE"
      Flag = True
   ElseIf UnqType = TRANSPORT_DETAIL_UNIQUE Then
      TableName = "SUPPLIER_TRANSPORT"
      FieldName1 = "SUPPLIER_TRANSPORT_CODE"
      FieldName2 = "SUPPLIER_TRANSPORT_ID"
      Flag = True
   ElseIf UnqType = SUPPLIER_ACCOUNT_NO_UNIQUE Then
      TableName = "SUPPLIER_ACCOUNT"
      FieldName1 = "SUPPLIER_ACCOUNT_NO"
      FieldName2 = "SUPPLIER_ACCOUNT_ID"
      Flag = True
   ElseIf UnqType = INCENTIVE_PD_UNIQUE Then
      TableName = "INCENTIVE"
      FieldName1 = "FREELANCE_ID"
      FieldName2 = "INCENTIVE_ID"
      FieldNameExTend = "PART_ITEM_ID"
      Flag = True
   ElseIf UnqType = INCENTIVE_PD_CUS_UNIQUE Then
      TableName = "INCENTIVE"
      FieldName1 = "FREELANCE_ID"
      FieldName2 = "INCENTIVE_ID"
      FieldNameExTend = "PART_ITEM_ID"
      FieldNameExTend2 = "CUSTOMER_ID"
      Flag = True
   ElseIf UnqType = BILLING_PAYMENT_NO_UNIQUE Then
      TableName = "BILLING_PAYMENT"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "BILLING_PAYMENT_ID"
      Flag = True
   ElseIf UnqType = DELIVERY_CUS_UNIQUE Then
      TableName = "DELIVERY_CUS_ITEM"
      FieldName1 = "DELIVERY_CUS_ITEM_CODE"
      FieldName2 = "DELIVERY_CUS_ITEM_ID"
      Flag = True
   ElseIf UnqType = WORKS_PRICE_ACTIVE_DATE_UNIQUE Then
      TableName = "EX_WORKS_PRICE"
      FieldName1 = "TO_VALID_DATE"
      FieldName2 = "FROM_ACTIVE_DATE"
      FieldNameExTend = "EX_WORKS_PRICE_TYPE"
      FieldNameExTend2 = "EX_WORKS_PRICE_ID"
      Flag = True
   ElseIf UnqType = PART_MASTER_NO_UNIQUE Then
      TableName = "PART_MASTER"
      FieldName1 = "PART_MASTER_NO"
      FieldName2 = "PART_MASTER_ID"
      Flag = True
   End If

   If Flag Then
      Count = glbDatabaseMngr.CountRecord(TableName, FieldName1, FieldName2, Key, id, glbErrorLog, FieldNameExTend, FieldNameExTendValue, FlagSpacial, FieldNameExTend2, FieldNameExTendValue2, ID_TYPE, ID_TEMP)
      If Count <> 0 Then
         CheckUniqueNs = False
      Else
         CheckUniqueNs = True
      End If
   End If
      
   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   CheckUniqueNs = False
End Function
Public Function Check2FlagConvert3(A As Long) As String
   If A = 1 Then
      Check2FlagConvert3 = "Y"
   ElseIf A = 2 Then
      Check2FlagConvert3 = "N"
   End If
End Function
Public Function Check2FlagConvert(A As Long) As String
   If A = 1 Then
      Check2FlagConvert = "N"
   Else
      Check2FlagConvert = "Y"
   End If
End Function
Public Function Check2FlagConvert2(A As Long) As String
   If A = 1 Then
      Check2FlagConvert2 = "N"
   Else
      Check2FlagConvert2 = ""
   End If
End Function

Public Function Check2FlagConvert4(A As Long) As String
   If A = 1 Then
      Check2FlagConvert4 = "N"
   Else
      Check2FlagConvert4 = "Y"
   End If
End Function


Public Function FlagToCheck(F As String) As Long
   If F = "Y" Then
      FlagToCheck = 1
   Else
      FlagToCheck = 0
   End If
End Function
Public Function FlagToOption(F As String) As Boolean
   If F = "Y" Then
      FlagToOption = True
   Else
      FlagToOption = False
   End If
End Function
Public Function ConvertFlag(F As String) As String
   If F = "Y" Then
      ConvertFlag = "*(ระงับการขาย)"
   Else
       ConvertFlag = ""
   End If
End Function
Public Function ConvertLoadFlag(F As String) As String
   If F = "N" Then
      ConvertLoadFlag = "(รอขึ้นอาหาร)"
   ElseIf F = "C" Then
      ConvertLoadFlag = "(รอชั่งน้ำหนักออก)"
   ElseIf F = "I" Then
       ConvertLoadFlag = "(รอออกใบส่งของ)"
   ElseIf F = "A" Then
       ConvertLoadFlag = "(รอเพิ่มรายการจาก SO)"
   ElseIf F = "O" Then
       ConvertLoadFlag = "(ออกใบส่งของโดยไม่มี SO)"
   ElseIf F = "L" Then
       ConvertLoadFlag = "(รอเพิ่มรายการอาหาร)"
   ElseIf F = "Y" Then
       ConvertLoadFlag = "ออกใบส่งของเรียบร้อย"
   Else
       ConvertLoadFlag = ""
   End If
End Function
Public Function ConvertThinkRate(Ind As Long) As String
   If Ind = 1 Then
      ConvertThinkRate = "มารับเอง"
   ElseIf Ind = 2 Then
      ConvertThinkRate = "รวมค่าขนส่ง"
   ElseIf Ind = 3 Then
      ConvertThinkRate = "แยกค่าขนส่ง"
   End If
End Function
Public Function ConvertSuccessFlagBillingSo(F As String) As String
   If F = "N" Then
       ConvertSuccessFlagBillingSo = "รอออกใบขึ้นอาหาร"
   ElseIf F = "Y" Then
       ConvertSuccessFlagBillingSo = "ออกใบขึ้นอาหารแล้ว"
   ElseIf F = "C" Then
       ConvertSuccessFlagBillingSo = "ออกใบส่งของเรียบร้อย"
   Else
       ConvertSuccessFlagBillingSo = ""
   End If
End Function
Public Function ConvertTypeProduct(id As Long) As String
   If id = 221 Then
       ConvertTypeProduct = "ผง"
   ElseIf id = 222 Then
       ConvertTypeProduct = "เม็ด"
   ElseIf id = 227 Then
       ConvertTypeProduct = "ครัม"
   Else
       ConvertTypeProduct = ""
   End If
End Function
Public Function Minus2Zero(A As Double) As Long
   If A < 0 Then
      Minus2Zero = 0
   Else
      Minus2Zero = A
   End If
End Function

Public Function Zero2One(A As Double) As Long
   If A = 0 Then
      Zero2One = 1
   Else
      Zero2One = A
   End If
End Function

Public Function Minus2Flag(A As Double) As String
   If A < 0 Then
      Minus2Flag = "Y"
   Else
      Minus2Flag = "N"
   End If
End Function

Public Function AdjustPage(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional TestFlag As Boolean = False, Optional SpaceCount As Long) As Boolean
Dim TempStr As String

   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   If (Vsp.Y1 + Offset - SpaceCount) > (Vsp.PageHeight - Vsp.MarginBottom) Then
      If Not TestFlag Then
         Vsp.NewPage
      End If
      AdjustPage = True
   Else
      AdjustPage = False
   End If
End Function

Public Function PatchTable(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional EnableFlag As Boolean = True, Optional SpaceCount As Long = 0) As Boolean
Dim TempStr As String
   
   If Not EnableFlag Then
      PatchTable = True
      Exit Function
   End If
   
   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   While Not AdjustPage(Vsp, Header, Body, Offset, True, SpaceCount)
      Call Vsp.AddTable(Header, "", Body)
   Wend
End Function

Public Sub PatchDB()
Dim p As CPatch

   Set p = New CPatch

'  If Not p.IsPatch("1_0_12_41_seub") Then
'     Call p.Patch_1_0_12_41_seub
'    End If
'
'  If Not p.IsPatch("1_0_12_42_seub") Then
'     Call p.Patch_1_0_12_42_seub
'    End If
'
'  If Not p.IsPatch("1_0_12_77_seub") Then
'     Call p.Patch_1_0_12_77_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_78_seub") Then
'     Call p.Patch_1_0_12_78_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_79_seub") Then
'     Call p.Patch_1_0_12_79_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_80_seub") Then
'     Call p.Patch_1_0_12_80_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_82_seub") Then
'     Call p.Patch_1_0_12_82_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_83_seub") Then
'     Call p.Patch_1_0_12_83_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_85_seub") Then
'     Call p.Patch_1_0_12_85_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_86_seub") Then
'     Call p.Patch_1_0_12_86_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_88_seub") Then
'     Call p.Patch_1_0_12_88_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_89_seub") Then
'     Call p.Patch_1_0_12_89_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_90_seub") Then
'     Call p.Patch_1_0_12_90_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_92_seub") Then
'     Call p.Patch_1_0_12_92_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_93_seub") Then
'     Call p.Patch_1_0_12_93_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_94_seub") Then
'     Call p.Patch_1_0_12_94_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_95_seub") Then
'     Call p.Patch_1_0_12_95_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_96_seub") Then
'     Call p.Patch_1_0_12_96_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_97_seub") Then
'     Call p.Patch_1_0_12_97_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_98_seub") Then
'     Call p.Patch_1_0_12_98_seub
'   End If
'
'  If Not p.IsPatch("1_0_12_99_seub") Then
'     Call p.Patch_1_0_12_99_seub
'   End If
'
'   If Not p.IsPatch("2006_09_21_1_jill") Then '1
'      Call p.Patch_2006_09_21_1_jill
'   End If
'
'   If Not p.IsPatch("2006_09_22_1_jill") Then '2
'      Call p.Patch_2006_09_22_1_jill
'   End If
'
'   If Not p.IsPatch("2006_09_23_1_jill") Then '3
'      Call p.Patch_2006_09_23_1_jill
'   End If
'
'   If Not p.IsPatch("2006_09_25_1_jill") Then '4
'      Call p.Patch_2006_09_25_1_jill
'   End If
'
'   If Not p.IsPatch("2006_09_25_2_jill") Then '5
'      Call p.Patch_2006_09_25_2_jill
'   End If
'
'   If Not p.IsPatch("2006_09_25_1_seub") Then '6
'      Call p.Patch_2006_09_25_1_seub
'   End If
'
'   If Not p.IsPatch("2006_09_26_1_jill") Then '7
'      Call p.Patch_2006_09_26_1_jill
'   End If
'
'   If Not p.IsPatch("2006_09_26_2_jill") Then '8
'      Call p.Patch_2006_09_26_2_jill
'   End If
'
'   If Not p.IsPatch("2006_09_26_3_jill") Then '9
'      Call p.Patch_2006_09_26_3_jill
'   End If
'
'   If Not p.IsPatch("2006_09_27_1_jill") Then '10
'      Call p.Patch_2006_09_27_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_02_1_jill") Then '11
'      Call p.Patch_2006_10_02_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_03_1_seub") Then '12
'      Call p.Patch_2006_10_03_1_seub
'   End If
'
'   If Not p.IsPatch("2006_10_04_1_jill") Then '13
'      Call p.Patch_2006_10_04_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_06_1_jill") Then '14
'      Call p.Patch_2006_10_06_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_09_1_jill") Then '15
'      Call p.Patch_2006_10_09_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_09_2_jill") Then '16
'      Call p.Patch_2006_10_09_2_jill
'   End If
'
'   If Not p.IsPatch("2006_10_12_1_jill") Then '17
'      Call p.Patch_2006_10_12_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_13_1_jill") Then '18
'      Call p.Patch_2006_10_13_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_16_1_jill") Then '19
'      Call p.Patch_2006_10_16_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_18_1_jill") Then '20
'      Call p.Patch_2006_10_18_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_18_1_jill") Then '20
'      Call p.Patch_2006_10_18_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_19_1_jill") Then '21
'      Call p.Patch_2006_10_19_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_19_2_jill") Then '22
'      Call p.Patch_2006_10_19_2_jill
'   End If
'
'   If Not p.IsPatch("2006_10_20_1_jill") Then '23
'      Call p.Patch_2006_10_20_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_25_1_jill") Then '24
'      Call p.Patch_2006_10_25_1_jill
'   End If
'
'   If Not p.IsPatch("2006_10_25_2_jill") Then '25
'      Call p.Patch_2006_10_25_2_jill
'   End If
'
'   If Not p.IsPatch("2006_11_02_1_seub") Then '25
'      Call p.Patch_2006_11_02_1_seub
'   End If
'
'   If Not p.IsPatch("2006_11_15_1_jill") Then '26
'      Call p.Patch_2006_11_15_1_jill
'   End If
'
'   If Not p.IsPatch("2006_11_19_1_jill") Then '27
'      Call p.Patch_2006_11_19_1_jill
'   End If
'
'   If Not p.IsPatch("2006_11_21_1_jill") Then '28
'      Call p.Patch_2006_11_21_1_jill
'   End If
'
''   If Not p.IsPatch("2006_11_25_1_jill") Then '29
''      Call p.Patch_2006_11_25_1_jill
''   End If
'
'   If Not p.IsPatch("2007_01_22_1_seub") Then '30
'      Call p.Patch_2007_01_22_1_seub
'   End If
'
'   If Not p.IsPatch("2007_01_23_1_seub") Then '31
'      Call p.Patch_2007_01_23_1_seub
'   End If
'
'   If Not p.IsPatch("2007_01_25_1_jill") Then '32
'      Call p.Patch_2007_01_25_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_01_1_jill") Then '33
'      Call p.Patch_2007_02_01_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_06_1_jill") Then '34
'      Call p.Patch_2007_02_06_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_06_2_jill") Then '35
'      Call p.Patch_2007_02_06_2_jill
'   End If
'
'   If Not p.IsPatch("2007_02_07_1_jill") Then '36
'      Call p.Patch_2007_02_07_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_07_2_jill") Then '37
'      Call p.Patch_2007_02_07_2_jill
'   End If
'
'   If Not p.IsPatch("2007_02_07_1_seub") Then '38
'      Call p.Patch_2007_02_07_1_seub
'   End If
'
'   If Not p.IsPatch("2007_02_12_1_jill") Then '39
'      Call p.Patch_2007_02_12_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_13_1_jill") Then '40
'      Call p.Patch_2007_02_13_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_13_2_jill") Then '41
'      Call p.Patch_2007_02_13_2_jill
'   End If
'
'   If Not p.IsPatch("2007_02_13_3_jill") Then '42
'      Call p.Patch_2007_02_13_3_jill
'   End If
'
'   If Not p.IsPatch("2007_02_19_1_jill") Then '43
'      Call p.Patch_2007_02_19_1_jill
'   End If
'
'   If Not p.IsPatch("2007_02_23_1_seub") Then '44
'      Call p.Patch_2007_02_23_1_seub
'   End If
'
'   If Not p.IsPatch("2007_03_02_1_jill") Then '45
'      Call p.Patch_2007_03_02_1_jill
'   End If
'
'   If Not p.IsPatch("2007_03_02_2_jill") Then '46
'      Call p.Patch_2007_03_02_2_jill
'   End If
'
''   If Not p.IsPatch("2007_03_13_1_jill") Then '47
''      Call p.Patch_2007_03_13_1_jill
''   End If
'
'   If Not p.IsPatch("2007_03_13_1_seub") Then '47
'      Call p.Patch_2007_03_13_1_seub
'   End If
'
'   If Not p.IsPatch("2007_03_13_2_jill") Then '48
'      Call p.Patch_2007_03_13_2_jill
'   End If
'
'   If Not p.IsPatch("2007_03_14_1_seub") Then '50
'      Call p.Patch_2007_03_14_1_seub
'   End If
'
'   If Not p.IsPatch("2007_03_15_1_jill") Then '51
'      Call p.Patch_2007_03_15_1_jill
'   End If
'
'   If Not p.IsPatch("2007_03_15_2_jill") Then '52
'      Call p.Patch_2007_03_15_2_jill
'   End If
'
'   If Not p.IsPatch("2007_03_15_3_jill") Then '53
'      Call p.Patch_2007_03_15_3_jill
'   End If
'
'   If Not p.IsPatch("2007_03_20_1_jill") Then '54
'      Call p.Patch_2007_03_20_1_jill
'   End If
'
'   If Not p.IsPatch("2007_03_20_2_jill") Then '55
'      Call p.Patch_2007_03_20_2_jill
'   End If
'
'   If Not p.IsPatch("2007_03_20_1_seub") Then '56
'      Call p.Patch_2007_03_20_1_seub
'   End If
'
'   If Not p.IsPatch("2007_03_20_2_seub") Then '57
'      Call p.Patch_2007_03_20_2_seub
'   End If
'
'   If Not p.IsPatch("2007_03_21_1_seub") Then '58
'      Call p.Patch_2007_03_21_1_seub
'   End If
'
'   If Not p.IsPatch("2007_03_22_1_seub") Then '59
'      Call p.Patch_2007_03_22_1_seub
'   End If
'
'   If Not p.IsPatch("2007_03_26_1_jill") Then '60
'      Call p.Patch_2007_03_26_1_jill
'   End If
'
'   If Not p.IsPatch("2007_03_26_2_jill") Then '61
'      Call p.Patch_2007_03_26_2_jill
'   End If
'
'   If Not p.IsPatch("2007_03_28_1_jill") Then '62
'      Call p.Patch_2007_03_28_1_jill
'   End If
'
'   If Not p.IsPatch("2007_03_28_2_jill") Then '63
'      Call p.Patch_2007_03_28_2_jill
'   End If
'
'   If Not p.IsPatch("2007_03_28_3_jill") Then '64
'      Call p.Patch_2007_03_28_3_jill
'   End If
'
'   If Not p.IsPatch("2007_03_28_4_jill") Then '65
'      Call p.Patch_2007_03_28_4_jill
'   End If
'
'   If Not p.IsPatch("2007_03_28_5_jill") Then '66
'      Call p.Patch_2007_03_28_5_jill
'   End If
'
'   If Not p.IsPatch("2007_03_28_6_jill") Then '67
'      Call p.Patch_2007_03_28_6_jill
'   End If
'
'   If Not p.IsPatch("2007_04_02_1_jill") Then '68
'      Call p.Patch_2007_04_02_1_jill
'   End If
'
'   If Not p.IsPatch("2007_04_02_2_jill") Then '69
'      Call p.Patch_2007_04_02_2_jill
'   End If
'
'   If Not p.IsPatch("2007_04_02_3_jill") Then '70
'      Call p.Patch_2007_04_02_3_jill
'   End If
'
'   If Not p.IsPatch("2007_04_02_4_jill") Then '71
'      Call p.Patch_2007_04_02_4_jill
'   End If
'
'   If Not p.IsPatch("2007_04_09_1_jill") Then '72
'      Call p.Patch_2007_04_09_1_jill
'   End If
'
'   If Not p.IsPatch("2007_04_09_2_jill") Then '73
'      Call p.Patch_2007_04_09_2_jill
'   End If
'
'   If Not p.IsPatch("2007_04_09_3_jill") Then '74
'      Call p.Patch_2007_04_09_3_jill
'   End If
'
'   If Not p.IsPatch("2007_04_09_4_jill") Then '75
'      Call p.Patch_2007_04_09_4_jill
'   End If
'
'   If Not p.IsPatch("2007_04_09_5_jill") Then '76
'      Call p.Patch_2007_04_09_5_jill
'   End If
'
'   If Not p.IsPatch("2007_04_09_6_jill") Then '77
'      Call p.Patch_2007_04_09_6_jill
'   End If
'
'   If Not p.IsPatch("2007_04_27_1_seub") Then '78
'      Call p.Patch_2007_04_27_1_seub
'   End If
'
'   If Not p.IsPatch("2007_04_27_2_seub") Then '79
'      Call p.Patch_2007_04_27_2_seub
'   End If
'
'   If Not p.IsPatch("2007_04_27_1_jill") Then '80
'      Call p.Patch_2007_04_27_1_jill
'   End If
'
'   If Not p.IsPatch("2007_05_22_1_seub") Then '81
'      Call p.Patch_2007_05_22_1_seub
'   End If
'
'   If Not p.IsPatch("2007_05_22_1_jill") Then '82
'      Call p.Patch_2007_05_22_1_jill
'   End If
'
'   If Not p.IsPatch("2007_05_28_1_seub") Then '83
'      Call p.Patch_2007_05_28_1_seub
'   End If
'
'   If Not p.IsPatch("2007_05_29_1_jill") Then '84
'      Call p.Patch_2007_05_29_1_jill
'   End If
'
'   If Not p.IsPatch("2007_05_29_2_jill") Then '85
'      Call p.Patch_2007_05_29_2_jill
'   End If
'
'   If Not p.IsPatch("2007_06_04_1_seub") Then '86
'      Call p.Patch_2007_06_04_1_seub
'   End If
'
'   If Not p.IsPatch("2007_06_06_1_seub") Then '87
'      Call p.Patch_2007_06_06_1_seub
'   End If
'
'   If Not p.IsPatch("2007_06_22_1_jill") Then '88
'      Call p.Patch_2007_06_22_1_jill
'   End If
'
'   If Not p.IsPatch("2007_06_22_2_jill") Then '89
'      Call p.Patch_2007_06_22_2_jill
'   End If
'
'   If Not p.IsPatch("2007_06_28_1_jill") Then '90
'      Call p.Patch_2007_06_28_1_jill
'   End If
'
'   If Not p.IsPatch("2007_06_21_1_seub") Then '91
'      Call p.Patch_2007_06_21_1_seub
'   End If
'
'   If Not p.IsPatch("2007_06_29_1_jill") Then '92
'      Call p.Patch_2007_06_29_1_jill
'   End If
'
'   If Not p.IsPatch("2007_06_29_2_jill") Then '93
'      Call p.Patch_2007_06_29_2_jill
'   End If
'
'   If Not p.IsPatch("2007_07_11_1_jill") Then '94
'      Call p.Patch_2007_07_11_1_jill
'   End If
'
'   If Not p.IsPatch("2007_07_18_1_jill") Then '95
'      Call p.Patch_2007_07_18_1_jill
'   End If
'
'   If Not p.IsPatch("2007_07_18_2_jill") Then '96
'      Call p.Patch_2007_07_18_2_jill
'   End If
'
'   If Not p.IsPatch("2007_07_18_3_jill") Then '97
'      Call p.Patch_2007_07_18_3_jill
'   End If
'
'   If Not p.IsPatch("2007_07_25_1_jill") Then '98
'      Call p.Patch_2007_07_25_1_jill
'   End If
'
'   If Not p.IsPatch("2007_07_26_1_jill") Then '99
'      Call p.Patch_2007_07_26_1_jill
'   End If
   
   If Not p.IsPatch("2007_08_14_1_jill") Then '100
      Call p.Patch_2007_08_14_1_jill
   End If
   
   If Not p.IsPatch("2007_10_24_1_jill") Then '101
      Call p.Patch_2007_10_24_1_jill
   End If
   
   If Not p.IsPatch("2008_02_13_1_jill") Then '102
      Call p.Patch_2008_02_13_1_jill
   End If
   
   If Not p.IsPatch("2009_05_04_1_jill") Then '103
      Call p.Patch_2009_05_04_1_jill
   End If
   
   If Not p.IsPatch("2009_05_04_2_jill") Then '104
      Call p.Patch_2009_05_04_2_jill
   End If
   
   If Not p.IsPatch("2009_05_04_3_jill") Then '105
      Call p.Patch_2009_05_04_3_jill
   End If
   
   If Not p.IsPatch("2013_05_02_1_jill") Then '106
      Call p.Patch_2013_05_02_1_jill
   End If
   
   If Not p.IsPatch("2013_05_02_2_jill") Then '107
      Call p.Patch_2013_05_02_2_jill
   End If
   
   If Not p.IsPatch("2013_05_02_3_jill") Then '108
      Call p.Patch_2013_05_02_3_jill
   End If
   
   If Not p.IsPatch("2013_05_02_4_jill") Then '109
      Call p.Patch_2013_05_02_4_jill
   End If
   
   If Not p.IsPatch("2013_05_15_1_yong") Then '110
      Call p.Patch_2013_05_15_1_yong
   End If
   
   If Not p.IsPatch("2013_06_03_1_jill") Then '111
      Call p.Patch_2013_06_03_1_jill
   End If
   
   If Not p.IsPatch("2013_06_17_1_yong") Then '112
      Call p.Patch_2013_06_17_1_yong
   End If
   
   If Not p.IsPatch("2013_06_26_1_jill") Then '113
      Call p.Patch_2013_06_26_1_jill
   End If
   
   If Not p.IsPatch("2013_06_26_2_jill") Then '114
      Call p.Patch_2013_06_26_2_jill
   End If
   
   If Not p.IsPatch("2015_01_13_1_pui") Then '115
      Call p.Patch_2015_01_13_1_pui
   End If
   'Patch_2015_01_13_1_pui()
     If Not p.IsPatch("2015_02_02_1_pui") Then '116
      Call p.Patch_2015_02_02_1_pui
   End If
   
   If Not p.IsPatch("2015_02_02_2_pui") Then '117
      Call p.Patch_2015_02_02_2_pui
   End If

     If Not p.IsPatch("2015_02_02_3_pui") Then '118
      Call p.Patch_2015_02_02_3_pui
   End If

   If Not p.IsPatch("2015_02_20_1_pui") Then '119
      Call p.Patch_2015_02_20_1_pui
   End If
   
     If Not p.IsPatch("2015_02_20_2_pui") Then '120
      Call p.Patch_2015_02_20_2_pui
   End If

     If Not p.IsPatch("2015_02_20_3_pui") Then '121
      Call p.Patch_2015_02_20_3_pui
   End If


   If Not p.IsPatch("2015_02_20_4_pui") Then '122
      Call p.Patch_2015_02_20_4_pui
   End If

   If Not p.IsPatch("2015_02_20_5_pui") Then '123
      Call p.Patch_2015_02_20_5_pui
   End If

   If Not p.IsPatch("2015_02_20_6_pui") Then '124
      Call p.Patch_2015_02_20_6_pui
   End If


   If Not p.IsPatch("2015_03_18_1_pui") Then '125
      Call p.Patch_2015_03_18_1_pui
   End If
   
   If Not p.IsPatch("2015_03_21_1_jill") Then '126
      Call p.Patch_2015_03_21_1_jill
   End If
   
   If Not p.IsPatch("2015_03_21_2_jill") Then '127
      Call p.Patch_2015_03_21_2_jill
   End If
   
    If Not p.IsPatch("2015_03_27_1_pui") Then '128
      Call p.Patch_2015_03_27_1_pui
   End If
   
     If Not p.IsPatch("2015_03_27_2_pui") Then '129
      Call p.Patch_2015_03_27_2_pui
   End If
   
  If Not p.IsPatch("2015_03_27_3_pui") Then '130
      Call p.Patch_2015_03_27_3_pui
   End If
   
    If Not p.IsPatch("2015_03_27_4_pui") Then '131
      Call p.Patch_2015_03_27_4_pui
   End If
   
    If Not p.IsPatch("2015_03_27_5_pui") Then '132
      Call p.Patch_2015_03_27_5_pui
   End If
   
    If Not p.IsPatch("2015_03_27_6_pui") Then '133
      Call p.Patch_2015_03_27_6_pui
   End If
   
     If Not p.IsPatch("2015_05_21_1_pui") Then '134
      Call p.Patch_2015_05_21_1_pui
   End If
   
     If Not p.IsPatch("2015_05_21_2_pui") Then '135
      Call p.Patch_2015_05_21_2_pui
   End If
   
     If Not p.IsPatch("2015_05_26_1_pui") Then '136
      Call p.Patch_2015_05_26_1_pui
   End If
   
   If Not p.IsPatch("2015_05_26_2_pui") Then '137
      Call p.Patch_2015_05_26_2_pui
   End If
   
   If Not p.IsPatch("2015_05_26_3_pui") Then '138
      Call p.Patch_2015_05_26_3_pui
   End If
   
   If Not p.IsPatch("2015_06_10_1_pui") Then '139
     Call p.Patch_2015_06_10_1_pui
   End If
   
    If Not p.IsPatch("2015_08_04_1_pui") Then '140
     Call p.Patch_2015_08_04_1_pui
   End If
   
   If Not p.IsPatch("2015_08_04_2_pui") Then '141
     Call p.Patch_2015_08_04_2_pui
   End If
   
   If Not p.IsPatch("2015_09_10_1_jill") Then '142
      Call p.Patch_2015_09_10_1_jill
   End If
   
   If Not p.IsPatch("2015_09_10_2_jill") Then '143
      Call p.Patch_2015_09_10_2_jill
   End If
   
   If Not p.IsPatch("2015_09_11_1_jill") Then '144
      Call p.Patch_2015_09_11_1_jill
   End If
   
   If Not p.IsPatch("2015_09_11_2_jill") Then '145
      Call p.Patch_2015_09_11_2_jill
   End If
   
   If Not p.IsPatch("2015_09_11_3_jill") Then '146
      Call p.Patch_2015_09_11_3_jill
   End If

   If Not p.IsPatch("2015_09_11_4_jill") Then '147
      Call p.Patch_2015_09_11_4_jill
   End If
   
   If Not p.IsPatch("2015_09_15_1_pui") Then '148
     Call p.Patch_2015_09_15_1_pui
   End If
   
   If Not p.IsPatch("2015_09_15_2_pui") Then '149
     Call p.Patch_2015_09_15_2_pui
   End If

           
   If Not p.IsPatch("2015_09_17_1_jill") Then '150
      Call p.Patch_2015_09_17_1_jill
   End If
   
    If Not p.IsPatch("2015_09_25_1_pui") Then '151
     Call p.Patch_2015_09_25_1_pui
   End If
   
  If Not p.IsPatch("2015_09_25_2_pui") Then '152
     Call p.Patch_2015_09_25_2_pui
   End If
   
    If Not p.IsPatch("2015_09_25_3_pui") Then '153
     Call p.Patch_2015_09_25_3_pui
   End If

   If Not p.IsPatch("2015_10_13_1_jill") Then '154
      Call p.Patch_2015_10_13_1_jill
   End If
   
   If Not p.IsPatch("2015_10_13_2_jill") Then '155
      Call p.Patch_2015_10_13_2_jill
   End If
   
   If Not p.IsPatch("2015_10_13_3_jill") Then '155
      Call p.Patch_2015_10_13_3_jill
   End If
   
   If Not p.IsPatch("2015_10_13_4_jill") Then '155
      Call p.Patch_2015_10_13_4_jill
   End If
   
   If Not p.IsPatch("2015_10_19_1_jill") Then '156
      Call p.Patch_2015_10_19_1_jill
   End If
   
   If Not p.IsPatch("2015_10_19_2_jill") Then '157
      Call p.Patch_2015_10_19_2_jill
   End If
   
   If Not p.IsPatch("2015_10_20_1_jill") Then '158
      Call p.Patch_2015_10_20_1_jill
   End If
   
   If Not p.IsPatch("2015_10_20_2_jill") Then '159
      Call p.Patch_2015_10_20_2_jill
   End If
   
   If Not p.IsPatch("2015_10_20_3_jill") Then '160
      Call p.Patch_2015_10_20_3_jill
   End If
   
   If Not p.IsPatch("2015_10_20_4_jill") Then '161
      Call p.Patch_2015_10_20_4_jill
   End If
   
   If Not p.IsPatch("2015_11_10_1_jill") Then '162
      Call p.Patch_2015_11_10_1_jill
   End If
   
   If Not p.IsPatch("2015_12_10_1_jill") Then '163
      Call p.Patch_2015_12_10_1_jill
   End If
   
   If Not p.IsPatch("2015_12_10_2_jill") Then '164
      Call p.Patch_2015_12_10_2_jill
   End If
   
   If Not p.IsPatch("2015_12_10_3_jill") Then '165
      Call p.Patch_2015_12_10_3_jill
   End If
   
   If Not p.IsPatch("2015_12_13_1_jill") Then '166
      Call p.Patch_2015_12_13_1_jill
   End If
   
   If Not p.IsPatch("2015_12_14_1_jill") Then '167
      Call p.Patch_2015_12_14_1_jill
   End If
   
   If Not p.IsPatch("2015_12_14_2_jill") Then '168
      Call p.Patch_2015_12_14_2_jill
   End If
   
     If Not p.IsPatch("2016_03_07_1_lek") Then '169
      Call p.Patch_2016_03_07_1_lek
   End If
   
   If Not p.IsPatch("2016_03_09_1_lek") Then '170
      Call p.Patch_2016_03_09_1_lek
   End If
   
   If Not p.IsPatch("2016_03_15_1_jill") Then '171
      Call p.Patch_2016_03_15_1_jill
   End If
   
   If Not p.IsPatch("2016_03_25_1_lek") Then '172
      Call p.Patch_2016_03_25_1_lek
   End If
   
    If Not p.IsPatch("2016_05_12_1_lek") Then '173
      Call p.Patch_2016_05_12_1_lek
   End If

   If Not p.IsPatch("2016_05_12_2_lek") Then '174
      Call p.Patch_2016_05_12_2_lek
   End If

   If Not p.IsPatch("2016_05_12_3_lek") Then '175
      Call p.Patch_2016_05_12_3_lek
   End If

   If Not p.IsPatch("2016_06_01_1_lek") Then '176
      Call p.Patch_2016_06_01_1_lek
   End If

   If Not p.IsPatch("2016_06_01_2_lek") Then '177
      Call p.Patch_2016_06_01_2_lek
   End If
   
   If Not p.IsPatch("2016_09_22_1_lek") Then '178
      Call p.Patch_2016_09_22_1_lek
   End If
   
   If Not p.IsPatch("2017_03_16_1_lek") Then '179
      Call p.Patch_2017_03_16_1_lek
   End If
   
   If Not p.IsPatch("2017_03_16_2_lek") Then '180
      Call p.Patch_2017_03_16_2_lek
   End If
   
   If Not p.IsPatch("2017_04_15_1_jill") Then '181
      Call p.Patch_2017_04_15_1_jill
   End If
   
   If Not p.IsPatch("2017_04_15_2_jill") Then '182
      Call p.Patch_2017_04_15_2_jill
   End If
   
   If Not p.IsPatch("2017_04_16_1_jill") Then '183
      Call p.Patch_2017_04_16_1_jill
   End If
   
   If Not p.IsPatch("2017_04_16_2_jill") Then '184
      Call p.Patch_2017_04_16_2_jill
   End If
   
   If Not p.IsPatch("2017_04_16_3_jill") Then '185
      Call p.Patch_2017_04_16_3_jill
   End If
   
   If Not p.IsPatch("2017_04_16_4_jill") Then '186
      Call p.Patch_2017_04_16_4_jill
   End If
   
   If Not p.IsPatch("2017_04_16_5_jill") Then '187
      Call p.Patch_2017_04_16_5_jill
   End If
   
   If Not p.IsPatch("2017_04_16_6_jill") Then '188
      Call p.Patch_2017_04_16_6_jill
   End If
   
   If Not p.IsPatch("2017_04_16_7_jill") Then '189
      Call p.Patch_2017_04_16_7_jill
   End If
   
   If Not p.IsPatch("2017_04_16_8_jill") Then '190
      Call p.Patch_2017_04_16_8_jill
   End If
   
   If Not p.IsPatch("2017_04_16_9_jill") Then '191
      Call p.Patch_2017_04_16_9_jill
   End If
   
   If Not p.IsPatch("2017_04_16_10_jill") Then '192
      Call p.Patch_2017_04_16_10_jill
   End If
   
   If Not p.IsPatch("2017_04_16_11_jill") Then '193
      Call p.Patch_2017_04_16_11_jill
   End If
   
   If Not p.IsPatch("2017_04_16_12_jill") Then '194
      Call p.Patch_2017_04_16_12_jill
   End If
   
   If Not p.IsPatch("2017_04_16_13_jill") Then '195
      Call p.Patch_2017_04_16_13_jill
   End If
   
   If Not p.IsPatch("2017_04_16_14_jill") Then '196
      Call p.Patch_2017_04_16_14_jill
   End If
   
   If Not p.IsPatch("2017_04_16_15_jill") Then '197
      Call p.Patch_2017_04_16_15_jill
   End If
   
   If Not p.IsPatch("2017_04_16_16_jill") Then '198
      Call p.Patch_2017_04_16_16_jill
   End If
   
   If Not p.IsPatch("2017_04_16_17_jill") Then '199
      Call p.Patch_2017_04_16_17_jill
   End If
   
   If Not p.IsPatch("2017_04_16_18_jill") Then '200
      Call p.Patch_2017_04_16_18_jill
   End If
   
   If Not p.IsPatch("2017_04_16_19_jill") Then '201
      Call p.Patch_2017_04_16_19_jill
   End If
   
   If Not p.IsPatch("2017_04_16_20_jill") Then '202
      Call p.Patch_2017_04_16_20_jill
   End If
   
   If Not p.IsPatch("2017_04_16_21_jill") Then '203
      Call p.Patch_2017_04_16_21_jill
   End If
   
   If Not p.IsPatch("2017_04_16_22_jill") Then '204
      Call p.Patch_2017_04_16_22_jill
   End If
   
   If Not p.IsPatch("2017_04_17_1_jill") Then '205
      Call p.Patch_2017_04_17_1_jill
   End If
   
   If Not p.IsPatch("2017_04_17_2_jill") Then '206
      Call p.Patch_2017_04_17_2_jill
   End If
   
   If Not p.IsPatch("2017_04_25_1_jill") Then '207
      Call p.Patch_2017_04_25_1_jill
   End If
   
   If Not p.IsPatch("2017_05_01_1_jill") Then '208
      Call p.Patch_2017_05_01_1_jill
   End If
   
   If Not p.IsPatch("2017_05_01_2_jill") Then '209
      Call p.Patch_2017_05_01_2_jill
   End If
   
'   If Not p.IsPatch("2017_05_01_3_jill") Then '210
'      Call p.Patch_2017_05_01_3_jill
'   End If
'
'   If Not p.IsPatch("2017_05_01_4_jill") Then '211
'      Call p.Patch_2017_05_01_4_jill
'   End If
   
   If Not p.IsPatch("2017_05_08_1_jill") Then '212
      Call p.Patch_2017_05_08_1_jill
   End If
      
   If Not p.IsPatch("2017_07_05_1_jill") Then '213
      Call p.Patch_2017_07_05_1_jill
   End If
      
   If Not p.IsPatch("2017_07_05_2_jill") Then '214
      Call p.Patch_2017_07_05_2_jill
   End If
   
   If Not p.IsPatch("2017_07_05_3_jill") Then '215
      Call p.Patch_2017_07_05_3_jill
   End If
   
   If Not p.IsPatch("2017_05_05_1_lek") Then '216
      Call p.Patch_2017_05_05_1_lek
   End If
   
   If Not p.IsPatch("2017_05_31_1_lek") Then '217
      Call p.Patch_2017_05_31_1_lek
   End If
   
   If Not p.IsPatch("2017_06_29_1_lek") Then '218
      Call p.Patch_2017_06_29_1_lek
   End If
   
   If Not p.IsPatch("2017_07_17_1_lek") Then '219
      Call p.Patch_2017_07_17_1_lek
   End If
   
   If Not p.IsPatch("2017_07_17_2_lek") Then '220
      Call p.Patch_2017_07_17_2_lek
   End If
   
'   If Not p.IsPatch("2017_07_21_01_lek") Then '221
'      Call p.Patch_2017_07_21_01_lek
'   End If
'
   If Not p.IsPatch("2017_08_09_1_lek") Then '222
      Call p.Patch_2017_08_09_1_lek
   End If
   
   If Not p.IsPatch("2017_08_14_1_lek") Then '223 เริ่ม version ใหม่
      Call p.Patch_2017_08_14_1_lek
   End If

   If Not p.IsPatch("2017_08_14_2_lek") Then '224
      Call p.Patch_2017_08_14_2_lek
   End If

   If Not p.IsPatch("2017_08_15_1_lek") Then '225
      Call p.Patch_2017_08_15_1_lek
   End If

   If Not p.IsPatch("2017_08_15_2_lek") Then '226
      Call p.Patch_2017_08_15_2_lek
   End If
   
   If Not p.IsPatch("2017_08_16_1_lek") Then '227
      Call p.Patch_2017_08_16_1_lek
   End If
   
   If Not p.IsPatch("2017_08_16_2_lek") Then '228
      Call p.Patch_2017_08_16_2_lek
   End If
   
   If Not p.IsPatch("2017_08_21_1_lek") Then '229
      Call p.Patch_2017_08_21_1_lek
   End If
   
   If Not p.IsPatch("2017_08_24_1_lek") Then '230
      Call p.Patch_2017_08_24_1_lek
   End If
   
   If Not p.IsPatch("2017_08_25_1_lek") Then '231
      Call p.Patch_2017_08_25_1_lek
   End If
   
   If Not p.IsPatch("2017_08_28_1_lek") Then '232
      Call p.Patch_2017_08_28_1_lek
   End If
   
   If Not p.IsPatch("2017_08_29_1_lek") Then '232
      Call p.Patch_2017_08_29_1_lek
   End If
   
   If Not p.IsPatch("2017_08_31_1_lek") Then '233
      Call p.Patch_2017_08_31_1_lek
   End If

   If Not p.IsPatch("2017_09_01_1_lek") Then '234
      Call p.Patch_2017_09_01_1_lek
   End If
   
   If Not p.IsPatch("2017_09_12_1_lek") Then '235
      Call p.Patch_2017_09_12_1_lek
   End If
   
   If Not p.IsPatch("2017_09_15_1_lek") Then '236
      Call p.Patch_2017_09_15_1_lek
   End If
   
   If Not p.IsPatch("2017_09_15_2_lek") Then '237
      Call p.Patch_2017_09_15_2_lek
   End If
   
   If Not p.IsPatch("2017_09_22_1_lek") Then '238
      Call p.Patch_2017_09_22_1_lek
   End If
   
   If Not p.IsPatch("2017_09_23_1_lek") Then '239
      Call p.Patch_2017_09_23_1_lek
   End If
   
   If Not p.IsPatch("2017_10_11_1_lek") Then '240
      Call p.Patch_2017_10_11_1_lek
   End If
   
   If Not p.IsPatch("2017_10_12_1_lek") Then '241
      Call p.Patch_2017_10_12_1_lek
   End If
   
   If Not p.IsPatch("2017_10_14_1_lek") Then '242
      Call p.Patch_2017_10_14_1_lek
   End If
   
   If Not p.IsPatch("2017_10_18_1_lek") Then '243
      Call p.Patch_2017_10_18_1_lek
   End If
   
   If Not p.IsPatch("2017_10_31_1_lek") Then '244
      Call p.Patch_2017_10_31_1_lek
   End If
   
   If Not p.IsPatch("2017_11_01_1_jill") Then '245
      Call p.Patch_2017_11_01_1_jill
   End If
   
   If Not p.IsPatch("2017_11_02_1_lek") Then '246
      Call p.Patch_2017_11_02_1_lek
   End If
   
   If Not p.IsPatch("2017_11_06_1_lek") Then '247
      Call p.Patch_2017_11_06_1_lek
   End If
   
   If Not p.IsPatch("2017_11_12_1_lek") Then '248
      Call p.Patch_2017_11_12_1_lek
   End If
   
   If Not p.IsPatch("2017_11_15_1_lek") Then '249
      Call p.Patch_2017_11_15_1_lek
   End If
   
   If Not p.IsPatch("2017_11_27_1_lek") Then '250
      Call p.Patch_2017_11_27_1_lek
   End If

   If Not p.IsPatch("2017_12_11_1_lek") Then '251
      Call p.Patch_2017_12_11_1_lek
   End If
   
   If Not p.IsPatch("2017_12_12_1_lek") Then '252
      Call p.Patch_2017_12_12_1_lek
   End If
   
   If Not p.IsPatch("2017_12_22_1_lek") Then '253
      Call p.Patch_2017_12_22_1_lek
   End If
   
   If Not p.IsPatch("2017_12_25_1_lek") Then '254
      Call p.Patch_2017_12_25_1_lek
   End If
   
   If Not p.IsPatch("2018_01_05_1_lek") Then '255
      Call p.Patch_2018_01_05_1_lek
   End If
   
'   If Not p.IsPatch("2018_01_30_1_lek") Then
'      Call p.Patch_2018_01_30_1_lek
'   End If

   If Not p.IsPatch("2018_02_15_1_lek") Then '256
      Call p.Patch_2018_02_15_1_lek
   End If
   
   If Not p.IsPatch("2018_02_20_1_lek") Then '257
      Call p.Patch_2018_02_20_1_lek
      Call UpdateDateLot(Nothing, Nothing)
   End If
   
'   If Not p.IsPatch("2018_02_28_1_jill") Then '258 ยกเลิกเพราะพี่จิวทำผิด ไม่เข้าใจโปรแกรม
'      Call p.Patch_2018_02_28_1_jill
'   End If
   
   If Not p.IsPatch("2018_03_02_1_lek") Then '259
      Call p.Patch_2018_03_02_1_lek
   End If
   
   If Not p.IsPatch("2018_03_06_1_lek") Then '260
      Call p.Patch_2018_03_06_1_lek
   End If
   
   If Not p.IsPatch("2018_03_06_2_lek") Then '261
      Call p.Patch_2018_03_06_2_lek
   End If
   
   If Not p.IsPatch("2018_03_15_1_lek") Then '262
      Call p.Patch_2018_03_15_1_lek
   End If
   
   If Not p.IsPatch("2018_03_22_1_lek") Then '263
      Call p.Patch_2018_03_22_1_lek
   End If
   
   If Not p.IsPatch("2018_03_31_1_lek") Then '264
      Call p.Patch_2018_03_31_1_lek
   End If
   
   If Not p.IsPatch("2018_04_05_1_lek") Then '265
      Call p.Patch_2018_04_05_1_lek
   End If
   
   If Not p.IsPatch("2018_04_06_1_lek") Then '266
      Call p.Patch_2018_04_06_1_lek
   End If
   
   If Not p.IsPatch("2018_04_07_1_lek") Then '267
      Call p.Patch_2018_04_07_1_lek
   End If
   
   If Not p.IsPatch("2018_04_10_1_lek") Then '268
      Call p.Patch_2018_04_10_1_lek
   End If
   
   If Not p.IsPatch("2018_04_19_1_lek") Then '269
      Call p.Patch_2018_04_19_1_lek
   End If
   
   If Not p.IsPatch("2018_04_23_1_lek") Then '269
      Call p.Patch_2018_04_23_1_lek
   End If
   
   If Not p.IsPatch("2018_04_23_2_lek") Then '270
      Call p.Patch_2018_04_23_2_lek
   End If
   
   If Not p.IsPatch("2018_04_26_1_dear") Then '271
      Call p.Patch_2018_04_26_1_dear
   End If
   
   If Not p.IsPatch("2018_04_27_1_dear") Then '272
      Call p.Patch_2018_04_27_1_dear
   End If
   
   If Not p.IsPatch("2018_04_30_1_dear") Then '273
      Call p.Patch_2018_04_30_1_dear
   End If

   If Not p.IsPatch("2018_05_03_1_dear") Then '274
      Call p.Patch_2018_05_03_1_dear
   End If
   
   If Not p.IsPatch("2018_05_14_1_lek") Then '275
      Call p.Patch_2018_05_14_1_lek
   End If
   
   If Not p.IsPatch("2018_05_16_1_lek") Then '276
      Call p.Patch_2018_05_16_1_lek
   End If

   If Not p.IsPatch("2018_05_30_1_lek") Then '277
      Call p.Patch_2018_05_30_1_lek
   End If
   
   If Not p.IsPatch("2018_06_04_1_lek") Then '278
      Call p.Patch_2018_06_04_1_lek
   End If
   
   If Not p.IsPatch("2018_06_12_1_lek") Then '279
      Call p.Patch_2018_06_12_1_lek
   End If
   
   If Not p.IsPatch("2018_06_13_1_lek") Then '280
      Call p.Patch_2018_06_13_1_lek
   End If
   
   If Not p.IsPatch("2018_06_25_1_lek") Then '281
      Call p.Patch_2018_06_25_1_lek
   End If
   
   If Not p.IsPatch("2018_06_25_2_lek") Then '282
      Call p.Patch_2018_06_25_2_lek
   End If
   
   If Not p.IsPatch("2018_06_28_1_lek") Then '283
      Call p.Patch_2018_06_28_1_lek
   End If
   
   If Not p.IsPatch("2018_06_28_2_lek") Then '284
      Call p.Patch_2018_06_28_2_lek
   End If
   
   If Not p.IsPatch("2018_07_03_1_lek") Then '285
      Call p.Patch_2018_07_03_1_lek
   End If

   If Not p.IsPatch("2018_07_09_1_lek") Then '286
      Call p.Patch_2018_07_09_1_lek
   End If
   
   If Not p.IsPatch("2018_07_09_2_lek") Then '287
      Call p.Patch_2018_07_09_2_lek
   End If
   
   If Not p.IsPatch("2018_07_16_1_lek") Then '288
      Call p.Patch_2018_07_16_1_lek
   End If
   
  If Not p.IsPatch("2018_07_18_1_lek") Then '289
      Call p.Patch_2018_07_18_1_lek
   End If
   
  If Not p.IsPatch("2018_07_18_2_lek") Then '290
      Call p.Patch_2018_07_18_2_lek
   End If
   
   If Not p.IsPatch("2018_07_24_1_lek") Then '291
      Call p.Patch_2018_07_24_1_lek
   End If
   
   If Not p.IsPatch("2018_07_26_1_lek") Then '292
      Call p.Patch_2018_07_26_1_lek
   End If
   
   If Not p.IsPatch("2018_08_15_1_lek") Then '293
      Call p.Patch_2018_08_15_1_lek
   End If
   
   If Not p.IsPatch("2018_08_27_1_lek") Then '294
      Call p.Patch_2018_08_27_1_lek
   End If
   
   If Not p.IsPatch("2018_08_28_1_lek") Then '295
      Call p.Patch_2018_08_28_1_lek
   End If
   
   If Not p.IsPatch("2018_08_30_1_lek") Then '296
      Call p.Patch_2018_08_30_1_lek
   End If

   If Not p.IsPatch("2018_08_31_1_lek") Then '297
      Call p.Patch_2018_08_31_1_lek
   End If
   
   If Not p.IsPatch("2018_08_31_2_lek") Then '298
      Call p.Patch_2018_08_31_2_lek
   End If
   
   If Not p.IsPatch("2018_09_20_1_lek") Then '299
      Call p.Patch_2018_09_20_1_lek
   End If
   
   If Not p.IsPatch("2018_09_20_2_lek") Then '300
      Call p.Patch_2018_09_20_2_lek
   End If
   
   If Not p.IsPatch("2018_09_28_1_lek") Then '301
      Call p.Patch_2018_09_28_1_lek
   End If
   
   If Not p.IsPatch("2018_10_04_1_lek") Then '302
      Call p.Patch_2018_10_04_1_lek
   End If
   
   If Not p.IsPatch("2018_10_05_1_lek") Then '303
      Call p.Patch_2018_10_05_1_lek
   End If
   
   If Not p.IsPatch("2018_10_08_1_lek") Then '304
      Call p.Patch_2018_10_08_1_lek
   End If
   
   If Not p.IsPatch("2018_10_10_1_lek") Then '305
      Call p.Patch_2018_10_10_1_lek
   End If
   
   If Not p.IsPatch("2018_10_17_1_lek") Then '306
      Call p.Patch_2018_10_17_1_lek
   End If
   
   If Not p.IsPatch("2018_11_19_1_lek") Then '307
      Call p.Patch_2018_11_19_1_lek
   End If
   
   If Not p.IsPatch("2018_11_26_1_lek") Then '308
      Call p.Patch_2018_11_26_1_lek
   End If
   
   If Not p.IsPatch("2018_11_20_28_lek_1") Then '309
      Call p.Patch_2018_11_20_28_lek_1
   End If

  If Not p.IsPatch("2018_12_03_1_lek") Then '310
      Call p.Patch_2018_12_03_1_lek
   End If
   
   If Not p.IsPatch("2018_12_11_1_lek") Then '311
      Call p.Patch_2018_12_11_1_lek
   End If
   
   If Not p.IsPatch("2018_12_18_1_lek") Then '312
      Call p.Patch_2018_12_18_1_lek
   End If
   
   If Not p.IsPatch("2018_12_25_1_lek") Then '313
      Call p.Patch_2018_12_25_1_lek
   End If
   
   If Not p.IsPatch("2019_01_08_1_lek") Then '314
      Call p.Patch_2019_01_08_1_lek
   End If
   
   If Not p.IsPatch("2019_01_08_3_lek") Then '315
      Call p.Patch_2019_01_08_3_lek
   End If
   
   If Not p.IsPatch("2019_01_08_2_lek") Then '316
      Call p.Patch_2019_01_08_2_lek
   End If
   
'   If Not p.IsPatch("2019_01_23_1_lek") Then '317
'      Call p.Patch_2019_01_23_1_lek
'   End If

   If Not p.IsPatch("2019_01_24_1_lek") Then '318
      Call p.Patch_2019_01_24_1_lek
   End If
   
   If Not p.IsPatch("2019_01_24_2_lek") Then '319
      Call p.Patch_2019_01_24_2_lek
   End If
   
   If Not p.IsPatch("2019_01_25_1_lek") Then '320
      Call p.Patch_2019_01_25_1_lek
   End If
   
  If Not p.IsPatch("2019_02_01_1_lek") Then '321
      Call p.Patch_2019_02_01_1_lek
   End If
   
   If Not p.IsPatch("2019_02_04_1_lek") Then '322
      Call p.Patch_2019_02_04_1_lek
   End If
   
   If Not p.IsPatch("2019_02_05_1_lek") Then '323
      Call p.Patch_2019_02_05_1_lek
   End If
   
   If Not p.IsPatch("2019_02_07_1_lek") Then '324
      Call p.Patch_2019_02_07_1_lek
   End If
   
   If Not p.IsPatch("2019_02_07_2_lek") Then '325
      Call p.Patch_2019_02_07_2_lek
   End If
   
   If Not p.IsPatch("2019_02_11_1_lek") Then '326
      Call p.Patch_2019_02_11_1_lek
   End If
   
   If Not p.IsPatch("2019_02_13_1_lek") Then '327
      Call p.Patch_2019_02_13_1_lek
   End If
   
   If Not p.IsPatch("2019_03_12_1_lek") Then '328
      Call p.Patch_2019_03_12_1_lek
   End If
   
   If Not p.IsPatch("2019_03_22_1_lek") Then '329
      Call p.Patch_2019_03_22_1_lek
   End If
   
   If Not p.IsPatch("2019_03_26_1_lek") Then '330
      Call p.Patch_2019_03_26_1_lek
   End If
   
   If Not p.IsPatch("2019_03_29_1_lek") Then '331
      Call p.Patch_2019_03_29_1_lek
   End If
   
   If Not p.IsPatch("2019_04_03_1_lek") Then '332
      Call p.Patch_2019_04_03_1_lek
   End If
   
   If Not p.IsPatch("2019_04_07_1_lek") Then '333
      Call p.Patch_2019_04_07_1_lek
   End If

   If Not p.IsPatch("2019_04_07_2_lek") Then '334
      Call p.Patch_2019_04_07_2_lek
   End If
   
   If Not p.IsPatch("2019_04_10_1_lek") Then '335
      Call p.Patch_2019_04_10_1_lek
   End If
   
   If Not p.IsPatch("2019_04_10_2_lek") Then '336
      Call p.Patch_2019_04_10_2_lek
   End If
   
   If Not p.IsPatch("2019_04_24_1_lek") Then '337
      Call p.Patch_2019_04_24_1_lek
   End If
   
   If Not p.IsPatch("2019_04_25_1_lek") Then '338
      Call p.Patch_2019_04_25_1_lek
   End If
   
   If Not p.IsPatch("2019_04_26_1_lek") Then '339
      Call p.Patch_2019_04_26_1_lek
   End If
   
   If Not p.IsPatch("2019_05_03_1_lek") Then '340
      Call p.Patch_2019_05_03_1_lek
   End If
   
   If Not p.IsPatch("2019_05_10_1_lek") Then '341
      Call p.Patch_2019_05_10_1_lek
   End If
   
   If Not p.IsPatch("2019_05_15_1_lek") Then '342
      Call p.Patch_2019_05_15_1_lek
   End If
   
  If Not p.IsPatch("2019_05_16_1_lek") Then '343
      Call p.Patch_2019_05_16_1_lek
   End If
   
   If Not p.IsPatch("2019_05_21_1_lek") Then '344
      Call p.Patch_2019_05_21_1_lek
   End If
   
   If Not p.IsPatch("2019_05_23_1_lek") Then '345
      Call p.Patch_2019_05_23_1_lek
   End If
   
  If Not p.IsPatch("2019_06_04_1_lek") Then '346
      Call p.Patch_2019_06_04_1_lek
   End If
   
   If Not p.IsPatch("2019_06_18_1_lek") Then '347
      Call p.Patch_2019_06_18_1_lek
   End If
   
  If Not p.IsPatch("2019_06_20_1_lek") Then '348
      Call p.Patch_2019_06_20_1_lek
   End If
   
  If Not p.IsPatch("2019_06_26_1_lek") Then '349
      Call p.Patch_2019_06_26_1_lek
   End If
   
  If Not p.IsPatch("2019_07_03_1_lek") Then '350
      Call p.Patch_2019_07_03_1_lek
   End If
   
  If Not p.IsPatch("2019_07_26_1_lek") Then '351
      Call p.Patch_2019_07_26_1_lek
   End If
   
  If Not p.IsPatch("2019_08_15_1_lek") Then '352
      Call p.Patch_2019_08_15_1_lek
   End If
   
  If Not p.IsPatch("2019_11_04_1_lek") Then '353
      Call p.Patch_2019_11_04_1_lek
   End If
   
  If Not p.IsPatch("2019_11_12_1_lek") Then '354
      Call p.Patch_2019_11_12_1_lek
   End If
   
   If Not p.IsPatch("2020_03_27_1_lek") Then '355
      Call p.Patch_2020_03_27_1_lek
   End If
   
   If Not p.IsPatch("2020_05_28_1_lek") Then '356
      Call p.Patch_2020_05_28_1_lek
   End If
'Patch_2020_05_28_1_lek
   Set p = Nothing
End Sub
Public Function MyDiff(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiff = 0
   Else
      MyDiff = CDbl(Format(D1 / D2, "0.00000000000000"))
   End If
End Function
Public Function PackAddress(Rs As ADODB.Recordset) As String
Dim AddressStr As String

   AddressStr = ""
   
   If NVLS(Rs("HOME_NO1"), "") <> "" Then
      AddressStr = AddressStr & NVLS(Rs("HOME_NO1"), "") & " "
   End If

   If NVLS(Rs("MOO1"), "") <> "" Then
      AddressStr = AddressStr & "หมู่." & NVLS(Rs("MOO1"), "") & " "
   End If

   If NVLS(Rs("SOI1"), "") <> "" Then
      AddressStr = AddressStr & "ซอย." & NVLS(Rs("SOI1"), "") & " "
   End If

   If NVLS(Rs("ROAD1"), "") <> "" Then
      AddressStr = AddressStr & "ถ." & NVLS(Rs("ROAD1"), "") & " "
   End If

   If NVLS(Rs("KWANG1"), "") <> "" Then
      AddressStr = AddressStr & "แขวง" & NVLS(Rs("KWANG1"), "") & " "
   End If

   If NVLS(Rs("KHATE1"), "") <> "" Then
      AddressStr = AddressStr & "เขต" & NVLS(Rs("KHATE1"), "") & " "
   End If

   If NVLS(Rs("PROVINCE"), "") <> "" Then
      AddressStr = AddressStr & "จ." & NVLS(Rs("PROVINCE"), "") & " "
   End If

   If NVLS(Rs("ZIPCODE1"), "") <> "" Then
      AddressStr = AddressStr & " " & NVLS(Rs("ZIPCODE1"), "") & " "
   End If

   PackAddress = AddressStr
End Function

Public Function MapText(msg As String) As String
   MapText = msg
End Function

Public Function SetReportConfig(Vsp As VSPrinter, ReportClassName As String, Optional ReportConfig As CReportConfig = Nothing, Optional Flag As Boolean = True) As Boolean
Dim I As Long
Dim Count As Long
Dim Rp As CReportConfig
Dim TempRs As ADODB.Recordset
Dim Rps As Collection
Dim iCount As Long

   If Rps Is Nothing Then
      Set TempRs = New ADODB.Recordset
      
      Set Rps = New Collection
      Set Rp = New CReportConfig
      
      Rp.REPORT_CONFIG_ID = -1
      Rp.COMPUTER_NAME = glbDatabaseMngr.GetComputerName    'จิวแก้แต่ ให้ การ ตั้ง ค่า เป็น ของ แต่ ละเครื่อง ละ จำเป็น ต้องใช้ แบบ ตาม เครื่อง
      Call Rp.QueryData(TempRs, iCount)
      Set Rp = Nothing
      
      While Not TempRs.EOF
         Set Rp = New CReportConfig
         
         Call Rp.PopulateFromRS(1, TempRs)
         Call Rps.add(Rp)
         
         Set Rp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Rp = Nothing
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   SetReportConfig = False
   For Each Rp In Rps
      If Rp.REPORT_KEY = ReportClassName Then
         If Flag Then
            Vsp.PaperSize = Rp.PAPER_SIZE
            Vsp.ORIENTATION = Rp.ORIENTATION
            Vsp.MarginBottom = Rp.MARGIN_BOTTOM * 567
            Vsp.MarginFooter = Rp.MARGIN_FOOTER * 567
            Vsp.MarginHeader = Rp.MARGIN_HEADER * 567
            Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
            Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
            Vsp.MarginTop = Rp.MARGIN_TOP * 567
   '         Vsp.FontName = Rp.FONT_NAME
            If Rp.FONT_SIZE > 0 Then
               Vsp.FontSize = Rp.FONT_SIZE
            End If
            Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
            Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
            If Rp.PAPER_HEIGHT > 0 Then
               Vsp.PaperWidth = Rp.PAPER_HEIGHT * 567
            End If
            If Rp.PAPER_WIDTH > 0 Then
               Vsp.PaperHeight = Rp.PAPER_HEIGHT * 567
            End If
         End If
         
         If Not ReportConfig Is Nothing Then
            Set ReportConfig = Rp
         End If
         
         SetReportConfig = True
         Exit Function
      End If
   Next Rp
   Set Rps = Nothing
End Function

Public Function GetBalanceItem(Col As Collection, PartItemID As Long, LocationID As Long, DocDate As Date) As Object
Dim D As Object
Dim Key As String
Dim MaxSeq As Long
Dim I As Long
Dim MaxIndex As Long
Static II As CLotItem
Dim MaxDate As Date


   For Each D In Col
'''Debug.Print D.TX_TYPE & ";" & D.PART_ITEM_ID & ";" & D.LOCATION_ID & ";" & DateToStringInt(D.DOCUMENT_DATE) & ";" & D.CURRENT_AMOUNT
      If (DateToStringInt(D.DOCUMENT_DATE) < DateToStringInt(DocDate)) And (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) Then
         If DateToStringInt(D.DOCUMENT_DATE) > DateToStringInt(MaxDate) Then
            MaxDate = DateToStringInt(D.DOCUMENT_DATE)
         End If
      End If
   Next D

'If MaxDate <= 0 Then
''Debug.Print
'End If

   I = 0
   MaxSeq = -1
   MaxIndex = -1
   For Each D In Col
      I = I + 1

      If (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) And _
         (DateToStringInt(D.DOCUMENT_DATE) = DateToStringInt(MaxDate)) Then
            If D.TRANSACTION_SEQ > MaxSeq Then
               MaxSeq = D.TRANSACTION_SEQ
               MaxIndex = I
            End If
      End If
   Next D
   
   If MaxIndex > 0 Then
      Set GetBalanceItem = Col(MaxIndex)
   Else
      If II Is Nothing Then
         Set II = New CLotItem
      End If
      Set GetBalanceItem = II
   End If
End Function

Public Function GetImportItem(m_TempCol As Collection, TempKey As String) As CLotItem
On Error Resume Next
Dim Ei As CLotItem
Static TempEi As CLotItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLotItem
      End If
      Set GetImportItem = TempEi
   Else
      Set GetImportItem = Ei
   End If
End Function

Public Function GetLotItem(m_TempCol As Collection, TempKey As String) As CLotItem
On Error Resume Next
Dim Ei As CLotItem
Static TempEi As CLotItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLotItem
      Else
         Set TempEi = Nothing
         Set TempEi = New CLotItem
      End If
      Set GetLotItem = TempEi
   Else
      Set GetLotItem = Ei
   End If
End Function
Public Function GetLotItemWh(m_TempCol As Collection, TempKey As String) As CLotItemWH
On Error Resume Next
Dim Ei As CLotItemWH
Static TempEi As CLotItemWH

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLotItemWH
      Else
         Set TempEi = Nothing
         Set TempEi = New CLotItemWH
      End If
      Set GetLotItemWh = TempEi
   Else
      Set GetLotItemWh = Ei
   End If
End Function
Public Function GetPlanPart(m_TempCol As Collection, TempKey As String) As CPlanPart
On Error Resume Next
Dim Ei As CPlanPart
Static TempEi As CPlanPart

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPlanPart
      Else
         Set TempEi = Nothing
         Set TempEi = New CPlanPart
      End If
      Set GetPlanPart = TempEi
   Else
      Set GetPlanPart = Ei
   End If
End Function

Public Function GetLotItemEx(m_TempCol As Collection, TempKey As String) As CLotItem
On Error Resume Next
Dim Ei As CLotItem
Static TempEi As CLotItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      Set GetLotItemEx = Nothing
   Else
      Set GetLotItemEx = Ei
   End If
End Function

Public Function GetTransferItem(m_TempCol As Collection, TempKey As String) As CTransferItem
On Error Resume Next
Dim Ei As CTransferItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      Set GetTransferItem = Nothing
   Else
      Set GetTransferItem = Ei
   End If
End Function

Public Function GetJobInOut(m_TempCol As Collection, TempKey As String) As CJobInput
On Error Resume Next
Dim Ei As CJobInput
Static TempEi As CJobInput

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CJobInput
      End If
      Set GetJobInOut = TempEi
   Else
      Set GetJobInOut = Ei
   End If
End Function
Public Function GetJob(m_TempCol As Collection, TempKey As String) As CJob
On Error Resume Next
Dim Ei As CJob
Static TempEi As CJob

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CJob
      End If
      Set GetJob = TempEi
   Else
      Set GetJob = Ei
   End If
End Function

Public Function GetJobParameter(m_TempCol As Collection, TempKey As String) As CJobParameter
On Error Resume Next
Dim Ei As CJobParameter
Static TempEi As CJobParameter

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CJobParameter
      End If
      Set GetJobParameter = TempEi
   Else
      Set GetJobParameter = Ei
   End If
End Function

Public Function GetLocation(m_TempCol As Collection, TempKey As String) As CLocation
On Error Resume Next
Dim Ei As CLocation
Static TempEi As CLocation

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLocation
      End If
      Set GetLocation = TempEi
   Else
      Set GetLocation = Ei
   End If
End Function

Public Function GetPartType(m_TempCol As Collection, TempKey As String) As CPartType
On Error Resume Next
Dim Ei As CPartType
Static TempEi As CPartType

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPartType
      End If
      Set GetPartType = TempEi
   Else
      Set GetPartType = Ei
   End If
End Function

Public Function GetSupplier(m_TempCol As Collection, TempKey As String) As CSupplier
On Error Resume Next
Dim Ei As CSupplier
Static TempEi As CSupplier

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CSupplier
      End If
      Set GetSupplier = TempEi
   Else
      Set GetSupplier = Ei
   End If
End Function

Public Function GetSupplierExt(m_TempCol As Collection, TempKey As String) As CSupplierExt
On Error Resume Next
Dim Ei As CSupplierExt
Static TempEi As CSupplierExt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CSupplierExt
      End If
      Set GetSupplierExt = TempEi
   Else
      Set GetSupplierExt = Ei
   End If
End Function

Public Function GetCustomer(m_TempCol As Collection, TempKey As String) As CCustomer
On Error Resume Next
Dim Ei As CCustomer
Static TempEi As CCustomer

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CCustomer
      End If
      Set GetCustomer = TempEi
   Else
      Set GetCustomer = Ei
   End If
End Function
Public Function GetEmployee(m_TempCol As Collection, TempKey As String) As CEmployee
On Error Resume Next
Dim Ei As CEmployee
Static TempEi As CEmployee

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CEmployee
      End If
      Set GetEmployee = TempEi
   Else
      Set GetEmployee = Ei
   End If
End Function

Public Function GetReceiptItem(m_TempCol As Collection, TempKey As String) As CReceiptItem
On Error Resume Next
Dim Ei As CReceiptItem
Static TempEi As CReceiptItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CReceiptItem
      End If
      Set GetReceiptItem = TempEi
   Else
      Set GetReceiptItem = Ei
   End If
End Function

Public Function GetBillingDiscount(m_TempCol As Collection, TempKey As String) As CBillingDiscount
On Error Resume Next
Dim Ei As CBillingDiscount
Static TempEi As CBillingDiscount

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBillingDiscount
      End If
      Set GetBillingDiscount = TempEi
   Else
      Set GetBillingDiscount = Ei
   End If
End Function

Public Function GetPackaging(m_TempCol As Collection, TempKey As String) As CPackaging
On Error Resume Next
Dim Ei As CPackaging
Static TempEi As CPackaging

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPackaging
      End If
      Set GetPackaging = TempEi
   Else
      Set GetPackaging = Ei
   End If
End Function

Public Function GetPurchaseExpense(m_TempCol As Collection, TempKey As String) As CPurchaseExpense
On Error Resume Next
Dim Ei As CPurchaseExpense
Static TempEi As CPurchaseExpense

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPurchaseExpense
      End If
      Set GetPurchaseExpense = TempEi
   Else
      Set GetPurchaseExpense = Ei
   End If
End Function

Public Function GetSupplierSpec(m_TempCol As Collection, TempKey As String) As CSupplierSpec
On Error Resume Next
Dim Ei As CSupplierSpec
Static TempEi As CSupplierSpec

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CSupplierSpec
      End If
      Set GetSupplierSpec = TempEi
   Else
      Set GetSupplierSpec = Ei
   End If
End Function

Public Function GetSupplierTrans(m_TempCol As Collection, TempKey As String) As CSupplierTranSport
On Error Resume Next
Dim Ei As CSupplierTranSport
Static TempEi As CSupplierTranSport

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CSupplierTranSport
      End If
      Set GetSupplierTrans = TempEi
   Else
      Set GetSupplierTrans = Ei
   End If
End Function

Public Function GetPartItem(m_TempCol As Collection, TempKey As String) As CPartItem
On Error Resume Next
Dim Ei As CPartItem
Static TempEi As CPartItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPartItem
      End If
      Set GetPartItem = TempEi
   Else
      Set GetPartItem = Ei
   End If
End Function

Public Function GetPartItemExt(m_TempCol As Collection, TempKey As String) As CPartItemExt
On Error Resume Next
Dim Ei As CPartItemExt
Static TempEi As CPartItemExt

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPartItemExt
      End If
      Set GetPartItemExt = TempEi
   Else
      Set GetPartItemExt = Ei
   End If
End Function

Public Function GetCostRaw(m_TempCol As Collection, TempKey As String) As CCostRaw
On Error Resume Next
Dim Ei As CCostRaw
Static TempEi As CCostRaw

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         'Set TempEi = New CCostRaw
      End If
      Set GetCostRaw = TempEi
   Else
      Set GetCostRaw = Ei
   End If
End Function

Public Function GetPartGroup(m_TempCol As Collection, TempKey As String) As CPartGroup
On Error Resume Next
Dim Ei As CPartGroup
Static TempEi As CPartGroup

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPartGroup
      End If
      Set GetPartGroup = TempEi
   Else
      Set GetPartGroup = Ei
   End If
End Function

Public Function GetFeature(m_TempCol As Collection, TempKey As String) As CFeature
On Error Resume Next
Dim Ei As CFeature
Static TempEi As CFeature

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CFeature
      End If
      Set GetFeature = TempEi
   Else
      Set GetFeature = Ei
   End If
End Function

Public Function GetBank(m_TempCol As Collection, TempKey As String) As CBank
On Error Resume Next
Dim Ei As CBank
Static TempEi As CBank

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBank
      End If
      Set GetBank = TempEi
   Else
      Set GetBank = Ei
   End If
End Function
Public Function GetBalanceAccum(m_TempCol As Collection, TempKey As String) As CBalanceAccum
On Error Resume Next
Dim Ei As CBalanceAccum
Static TempEi As CBalanceAccum

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBalanceAccum
      End If
      Set GetBalanceAccum = TempEi
   Else
      Set GetBalanceAccum = Ei
   End If
End Function

Public Function GetBankBranch(m_TempCol As Collection, TempKey As String) As CBankBranch
On Error Resume Next
Dim Ei As CBankBranch
Static TempEi As CBankBranch

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBankBranch
      End If
      Set GetBankBranch = TempEi
   Else
      Set GetBankBranch = Ei
   End If
End Function
Public Sub LoadPictureFromFile(FileName As String, Pc As PictureBox)
On Error Resume Next
    If Dir(FileName) <> "" Then
      Pc.Picture = LoadPicture(FileName)
   End If
End Sub
Public Function PatchWildCard(T As String) As String
   If Len(Trim(T)) <> 0 Then
      PatchWildCard = T & "%"
   Else
      PatchWildCard = T
   End If
End Function

Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
Dim LastDay As Byte
   LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", ValidDate) + 1, ValidDate))))
   LastDayOfMonth = LastDay
End Function

Public Sub GetFirstLastDate(D As Date, FD As Date, LD As Date, Optional add As Long = 0)
Dim MM As Long
Dim DD1 As Long
Dim DD2 As Long
Dim YYYY As Long

   D = DateAdd("m", add, D)
   MM = Month(D)
   DD1 = 1
   DD2 = LastDayOfMonth(D)
   YYYY = Year(D)
   
   FD = DateSerial(YYYY, MM, DD1)
   LD = DateSerial(YYYY, MM, DD2)
End Sub

Public Function PaymentTypeToText(id As PAYMENT_TYPE) As String
   If id = CASH_PMT Then
      PaymentTypeToText = MapText("เงินสด")
   ElseIf id = CHECK_PMT Then
      PaymentTypeToText = MapText("เช็ค")
   ElseIf id = CREDITCRD_PMT Then
      PaymentTypeToText = MapText("บัตรเครดิต")
   ElseIf id = BANKTRF_PMT Then
      PaymentTypeToText = MapText("โอนผ่านธนาคาร")
   ElseIf id = CASHRET_PMT Then
      PaymentTypeToText = MapText("เงินสดขายคืน")
   End If
End Function

Public Function RatioTypeToText(id As RATIO_TYPE) As String
   If id = RATIO_COST Then
      RatioTypeToText = MapText("ตามต้นทุนผลิต")
   ElseIf id = RATIO_QUANTITY Then
      RatioTypeToText = MapText("ตามปริมาณผลิต")
   ElseIf id = RATIO_RAW Then
      RatioTypeToText = MapText("จำนวนวัตถุดิบที่ใช้")
   ElseIf id = RATIO_VARY Then
      RatioTypeToText = MapText("แปรผันตามจำนวนผลิต")
   ElseIf id = RATIO_PERCENT Then
      RatioTypeToText = MapText("แปรผันตาม % ของมูลค่าวัตถุดิบที่ใช้")
   End If
End Function
Public Function PictureTypeToText(id As PICTURE_TYPE) As String
   If id = HEAD_ACCOUNT Then
      PictureTypeToText = MapText("ใบเปิดหน้าบัญชี")
   ElseIf id = HEAD_PART Then
     PictureTypeToText = MapText("รูปภาพประกอบ")
   End If
End Function

Public Function ParcelTypeToText(id As PARCEL_TYPE) As String
   If id = PARCEL_BAG Then
      ParcelTypeToText = MapText("BAG")
   ElseIf id = PARCEL_BULK Then
      ParcelTypeToText = MapText("BULK")
   ElseIf id = PARCEL_ALL Then
      ParcelTypeToText = MapText("ทั้งหมด")
   End If
End Function

Public Function ConvertPerPack(id As Integer) As String
   If id = 10 Then
      ConvertPerPack = MapText("BAG")
   ElseIf id = 21 Then
      ConvertPerPack = MapText("BULK")
   Else
     ConvertPerPack = MapText("")
   End If
End Function

Public Function MyDiffEx(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiffEx = 0
   Else
      MyDiffEx = D1 / D2
   End If
End Function


Public Sub ReArrangeRatio(Col As Collection)
Dim Fi As CFormulaItem
Dim Sum As Double

   Sum = 0
   For Each Fi In Col
      If Fi.Flag <> "D" Then
         Sum = Sum + Fi.REAL_AMOUNT
      End If
   Next Fi

   For Each Fi In Col
      If Fi.Flag <> "D" Then
         Fi.ITEM_PERCENT = MyDiffEx(Fi.REAL_AMOUNT, Sum) * 100

         If Fi.Flag <> "A" Then
            Fi.Flag = "E"
         End If
      End If
   Next Fi
End Sub

Public Function BillingDocType2Set(DocType As Long) As String
   If DocType = 1 Then
      BillingDocType2Set = "(1)"
   ElseIf DocType = 2 Then
      BillingDocType2Set = "(2)"
   Else
      BillingDocType2Set = "(1, 2)"
   End If
End Function

Public Sub StartExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
   Vsp.ExportFile = glbParameterObj.ReportFile
   Vsp.ExportFormat = vpxPlainHTML
End Sub

Public Sub CloseExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
End Sub

Public Function GetDoItem(m_TempCol As Collection, TempKey As String) As CDoItem
On Error Resume Next
Dim Ei As CDoItem
Static TempEi As CDoItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CDoItem
      End If
      Set GetDoItem = TempEi
   Else
      Set GetDoItem = Ei
   End If
End Function

Public Function GetInventoryDoc(m_TempCol As Collection, TempKey As String) As Object
On Error Resume Next
Dim Ei As Object

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      Set GetInventoryDoc = Nothing
   Else
      Set GetInventoryDoc = Ei
   End If
End Function
Public Function GetMasterRef(m_TempCol As Collection, TempKey As String) As CMasterRef
On Error Resume Next
Dim Ei As CMasterRef
Static TempEi As CMasterRef

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CMasterRef
      End If
      Set GetMasterRef = TempEi
   Else
      Set GetMasterRef = Ei
   End If
End Function

Public Function GetPaymentItem(m_TempCol As Collection, TempKey As String) As CPaymentItem
On Error Resume Next
Dim Ei As CPaymentItem
Static TempEi As CPaymentItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPaymentItem
      End If
      Set GetPaymentItem = TempEi
   Else
      Set GetPaymentItem = Ei
   End If
End Function

Private Function CompareKeyEx(D1 As Object, D2 As Object, CompareType As Long) As Boolean
Dim TempResult As Boolean

   If CompareType = 1 Then
      If D1.LOCATION_NO = D2.LOCATION_NO Then
         CompareKeyEx = D1.PART_NO < D2.PART_NO
      Else
         CompareKeyEx = D1.LOCATION_NO > D2.LOCATION_NO
      End If
   Else
      If D1.PART_NO = D2.PART_NO Then
         CompareKeyEx = D1.LOCATION_NO < D2.LOCATION_NO
      Else
         CompareKeyEx = D1.PART_NO < D2.PART_NO
      End If
   End If
End Function

Public Sub SelectionsortEx(List As Collection, MIN As Long, MAX As Long, CompareType As Long)
Dim I As Long
Dim J As Long
Dim best_value As Object
Dim Temp As Object
Dim best_j As Integer

    For I = MIN To MAX - 1
        Set best_value = List(I)
        best_j = I
        For J = I + 1 To MAX
            If CompareKeyEx(List(J), best_value, CompareType) Then
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
Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function
Public Function GenerateInsertSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String

   SQL = "INSERT INTO " & O.TableName & vbCrLf & " (" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ") " & vbCrLf & "VALUES " & vbCrLf & "(" & vbCrLf
         Else
            Sep = ", " & vbCrLf
         End If
         
         SQL = SQL & Tf.FieldName & Sep
      End If
   Next Tf
   
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ")"
         Else
            Sep = ", " & vbCrLf
         End If
'''Debug.Print "---" & Tf.FieldName
         SQL = SQL & Tf.TransformToSQLString & Sep
'''Debug.Print "---" & Tf.FieldName
      End If
   Next Tf
   
   GenerateInsertSQL = SQL
End Function

Public Function GenerateUpdateSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String
Dim TempKeyName As String
Dim TempKeyVal As Long

   SQL = "UPDATE " & O.TableName & " SET" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT And Tf.FieldCat <> CREATE_DATE_CAT And Tf.FieldCat <> CREATE_BY_CAT Then
         If Tf.FieldCat = ID_CAT Then
            TempKeyName = Tf.FieldName
            TempKeyVal = Tf.GetValue
         Else
            If Tf.FieldCat = MODIFY_BY_CAT Then
               Sep = "" & vbCrLf
            Else
               Sep = ", " & vbCrLf
            End If
            
            SQL = SQL & Tf.FieldName & " = " & Tf.TransformToSQLString & Sep
         End If
      End If
   Next Tf
      
   SQL = SQL & "WHERE " & TempKeyName & " = " & TempKeyVal
   
   GenerateUpdateSQL = SQL
End Function
Public Sub PopulateInternalField(ShowMode As SHOW_MODE_TYPE, O As Object)
Dim Tf As CTableField
Dim TempID As Long
Dim InternalDate As String

   For Each Tf In O.m_FieldList
      If Tf.FieldCat = ID_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetSeqID(O.SequenceName, TempID, glbErrorLog)
            Call Tf.SetValue(TempID)
         End If
      ElseIf Tf.FieldCat = CREATE_DATE_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         End If
      ElseIf Tf.FieldCat = MODIFY_DATE_CAT Then
         'If ShowMode = SHOW_EDIT Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         'End If
      ElseIf Tf.FieldCat = CREATE_BY_CAT Then
         If ShowMode = SHOW_ADD Then
            Call Tf.SetValue(glbLoginID)
         End If
      ElseIf Tf.FieldCat = MODIFY_BY_CAT Then
         'If ShowMode = SHOW_EDIT Then
            Call Tf.SetValue(glbLoginID)
         'End If
      End If
   Next Tf
End Sub

Public Function CashDocType2Text(id As CASH_DOC_TYPE) As String
   If id = CASH_DEPOSIT Then
      CashDocType2Text = "ใบนำฝากเงิน"
   ElseIf id = CASH_PITTYCASH Then
      CashDocType2Text = "ใบเบิก/เคลียร์เงินสดย่อย"
   ElseIf id = CASH_TRANSFER Then
      CashDocType2Text = "ใบโอนเงินระหว่างบัญชี"
   ElseIf id = CASH_WITHDRAW Then
      CashDocType2Text = "ใบถอนเงิน (ใช้เพื่อเป็นเงินสดย่อย)"
   ElseIf id = CASH_WHTHDRAW2 Then
      CashDocType2Text = "ใบถอนเงิน/โอนเงิน (ทั่วไป)"
   ElseIf id = CASH_DEPOSIT2 Then
      CashDocType2Text = "ใบนำฝากเงิน/โอนเงิน (ทั่วไป)"
   ElseIf id = POST_CHEQUE Then
      CashDocType2Text = "ใบยืนยันการ clearing เช็ค"
   ElseIf id = WAITING_CHEQUE Then
      CashDocType2Text = "ใบยืนยันการรับเช็คของซัพพลายเออร์"
   ElseIf id = PASSED_CHEQUE Then
      CashDocType2Text = "ใบยืนยันการ Clearing เช็คจ่าย"
   End If
   
End Function

Public Function ChequeDocType2Text(id As CHEQUE_DOC_TYPE) As String
   If id = CHECK_CHEQUE Then
      ChequeDocType2Text = "ใบตรวจสอบการรับ/คืนเช็ค"
End If
End Function
Public Function GetCashTran(m_TempCol As Collection, TempKey As String) As CCashTran
On Error Resume Next
Dim Ei As CCashTran
Static TempEi As CCashTran

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CCashTran
      End If
      Set GetCashTran = TempEi
   Else
      Set GetCashTran = Ei
   End If
End Function

Public Function GetObject(ClassName As String, Optional m_TempCol As Collection, Optional TempKey As String, Optional SetNew As Boolean = True) As Object
On Error Resume Next
Dim Ei As Object
Dim TempEi As Object

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing And SetNew Then
         Set TempEi = CreateObject(ClassName)
         If TempEi Is Nothing Then
            Set TempEi = GetNewClass(ClassName)
         End If
      End If
      Set GetObject = TempEi
   Else
      Set GetObject = Ei
   End If
End Function
Public Function GetNewClass(ClassName As String) As Object
   If ClassName = "CConfigDoc" Then
      Static m_CConfigDoc As CConfigDoc
      If m_CConfigDoc Is Nothing Then
         Set m_CConfigDoc = New CConfigDoc
      End If
      Set GetNewClass = m_CConfigDoc
   ElseIf ClassName = "CJobInput" Then
      Static m_CJobInput As CJobInput
      If m_CJobInput Is Nothing Then
         Set m_CJobInput = New CJobInput
      End If
      Set GetNewClass = m_CJobInput
   ElseIf ClassName = "CCostItemRaw" Then
      Static m_CCostItemRaw As CCostItemRaw
      If m_CCostItemRaw Is Nothing Then
         Set m_CCostItemRaw = New CCostItemRaw
      End If
      Set GetNewClass = m_CCostItemRaw
   ElseIf ClassName = "CExpenseDetail" Then
      Static m_CExpenseDetail As CExpenseDetail
      If m_CExpenseDetail Is Nothing Then
         Set m_CExpenseDetail = New CExpenseDetail
      End If
      Set GetNewClass = m_CExpenseDetail
   ElseIf ClassName = "CBillingDoc" Then
      Static m_CBillingDoc As CBillingDoc
      If m_CBillingDoc Is Nothing Then
         Set m_CBillingDoc = New CBillingDoc
      End If
      Set GetNewClass = m_CBillingDoc
   ElseIf ClassName = "CSupItem" Then
      Static m_CSupItem As CSupItem
      If m_CSupItem Is Nothing Then
         Set m_CSupItem = New CSupItem
      End If
      Set GetNewClass = m_CSupItem
   ElseIf ClassName = "CMonthlyAccum" Then
      Static m_CMonthlyAccum As CMonthlyAccum
      If m_CMonthlyAccum Is Nothing Then
         Set m_CMonthlyAccum = New CMonthlyAccum
      End If
      Set GetNewClass = m_CMonthlyAccum
   ElseIf ClassName = "CReceiptItem" Then
      Static m_CReceiptItem As CReceiptItem
      If m_CReceiptItem Is Nothing Then
         Set m_CReceiptItem = New CReceiptItem
      End If
      Set GetNewClass = m_CReceiptItem
   ElseIf ClassName = "CPlanningItem" Then
      Static m_CPlanningItem As CPlanningItem
      If m_CPlanningItem Is Nothing Then
         Set m_CPlanningItem = New CPlanningItem
      End If
      Set GetNewClass = m_CPlanningItem
   ElseIf ClassName = "CUserAccount" Then
      Static m_CUserAccount As CUserAccount
      If m_CUserAccount Is Nothing Then
         Set m_CUserAccount = New CUserAccount
      End If
      Set GetNewClass = m_CUserAccount
    ElseIf ClassName = "CSumComeIn" Then
      Static m_CSumComeIn As CSumComeIn
      If m_CSumComeIn Is Nothing Then
         Set m_CSumComeIn = New CSumComeIn
      End If
      Set GetNewClass = m_CSumComeIn
    ElseIf ClassName = "CSumInventoryAccount" Then
      Static m_CSumInventoryAccount As CSumInventoryAccount
      If m_CSumInventoryAccount Is Nothing Then
         Set m_CSumInventoryAccount = New CSumInventoryAccount
      End If
      Set GetNewClass = m_CSumInventoryAccount
    ElseIf ClassName = "CInventoryActItem" Then
      Static m_CInventoryActItem As CInventoryActItem
      If m_CInventoryActItem Is Nothing Then
         Set m_CInventoryActItem = New CInventoryActItem
      End If
      Set GetNewClass = m_CInventoryActItem
   ElseIf ClassName = "CTotalCommission" Then
      Static m_CTotalCommission As CTotalCommission
      If m_CTotalCommission Is Nothing Then
         Set m_CTotalCommission = New CTotalCommission
      End If
      Set GetNewClass = m_CTotalCommission
   
   ElseIf ClassName = "CCommissionCost" Then
      Static m_CCommissionCost As CCommissionCost
      If m_CCommissionCost Is Nothing Then
         Set m_CCommissionCost = New CCommissionCost
      End If
      Set GetNewClass = m_CCommissionCost
    ElseIf ClassName = "CWeight" Then
      Static m_CWeight As CWeight
      If m_CWeight Is Nothing Then
         Set m_CWeight = New CWeight
      End If
      Set GetNewClass = m_CWeight
   ElseIf ClassName = "CPartItem" Then
      Static m_CPartItem As CPartItem
      If m_CPartItem Is Nothing Then
         Set m_CPartItem = New CPartItem
      End If
      Set GetNewClass = m_CPartItem
   
    ElseIf ClassName = "CCommissionIncentive" Then
      Static m_CCommissionIncentive As CCommissionIncentive
      If m_CCommissionIncentive Is Nothing Then
         Set m_CCommissionIncentive = New CCommissionIncentive
      End If
      Set GetNewClass = m_CCommissionIncentive
   
   ElseIf ClassName = "CDoItem" Then
      Static m_CDoItem As CDoItem
      If m_CDoItem Is Nothing Then
         Set m_CDoItem = New CDoItem
      End If
      Set GetNewClass = m_CDoItem
      
   End If

End Function
Public Function SavePictureToServer(PathSource As String) As String
Dim id As Long
   Call glbDatabaseMngr.GetSeqID("PICTURE_FILE_SEQ", id, glbErrorLog)

   If Right(glbParameterObj.MapDrivePicture, 1) <> "\" Then
      glbParameterObj.MapDrivePicture = glbParameterObj.MapDrivePicture & "\"
   End If

   If Dir(glbParameterObj.MapDrivePicture, vbDirectory) = "" Then
      On Error GoTo NoPath
      MkDir (glbParameterObj.MapDrivePicture)
   End If

   If Dir(glbParameterObj.MapDrivePicture & Year(Now) & Format(Month(Now), "00"), vbDirectory) = "" Then
      On Error GoTo NoPath
      MkDir (glbParameterObj.MapDrivePicture & Year(Now) & Format(Month(Now), "00"))
   End If

   Call FileCopy(PathSource, glbParameterObj.MapDrivePicture & Year(Now) & Format(Month(Now), "00") & "\" & Format(id, "000000") & ".JPG")
   
   SavePictureToServer = Year(Now) & Format(Month(Now), "00") & "\" & Format(id, "000000") & ".JPG"
   
   Exit Function
NoPath:
   glbErrorLog.LocalErrorMsg = "Error ไม่มีเนื้อที่จะใช้บันทึก กรุณาติดต่อผู้ดูแลระบบเพื่อสร้างเนื้อที่ในการบันทึก"
   glbErrorLog.ShowUserError
End Function
Public Sub DeletePictureFromDisk(Path As String)
   If Right(glbParameterObj.MapDrivePicture, 1) <> "\" Then
      glbParameterObj.MapDrivePicture = glbParameterObj.MapDrivePicture & "\"
   End If
   Call Kill(glbParameterObj.MapDrivePicture & Path)
End Sub
Public Sub AddMemoNote()
Dim ItemCount As Long
Dim OKClick As Boolean

   frmAddEditMemoNote.HeaderText = MapText("เพิ่ม MEMO")
   frmAddEditMemoNote.ShowMode = SHOW_ADD
   Load frmAddEditMemoNote
   frmAddEditMemoNote.Show 1
   
   OKClick = frmAddEditMemoNote.OKClick
   
   Unload frmAddEditMemoNote
   Set frmAddEditMemoNote = Nothing
   
End Sub
Public Function CheckTask() As Boolean
Dim Mn As CMemoNote
Dim Rs As ADODB.Recordset
Dim ItemCount As Long
Dim TempStr As String

   Set Mn = New CMemoNote
   Set Rs = New ADODB.Recordset
   Call Mn.SetFieldValue("MEMO_NOTE_ID", -1)
   Call Mn.SetFieldValue("FROM_DATE_FINISH", Now)
   
   Call Mn.SetFieldValue("MEMO_NOTE_WARN", "Y")
   Call Mn.SetFieldValue("MEMO_NOTE_CREATE_TO", glbUser.REAL_USER_ID)
   Call Mn.SetFieldValue("ORDER_TYPE", 1)
   Call Mn.QueryData(2, Rs, ItemCount)
   
'   While Not (Rs.EOF)
'      Call Mn.PopulateFromRS(2, Rs)
'      TempStr = TempStr & "วันที่สร้าง " & DateToStringExtEx2(Mn.GetFieldValue("MEMO_NOTE_DATE_CREATE")) & "      วันที่คาดการสำเร็จ " & DateToStringExtEx2(Mn.GetFieldValue("MEMO_NOTE_DATE_FINISH")) & "     หัวข้อ " & Mn.GetFieldValue("MEMO_NOTE_SUBJECT")
'      TempStr = TempStr & vbCrLf
'      Rs.MoveNext
'   Wend
   If ItemCount > 0 Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านมีงานที่ค้างอยู่ จำนวน " & ItemCount & " งาน  ต้องการที่จำดูทันทีหรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Load frmMemoNote
         frmMemoNote.Show 1
      
         Unload frmMemoNote
         Set frmMemoNote = Nothing
         CheckTask = True
      Else
         CheckTask = False
      End If
   Else
      CheckTask = True
   End If
   Set Mn = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
End Function

Public Function GetSupItem(m_TempCol As Collection, TempKey As String) As CSupItem
On Error Resume Next
Dim Ei As CSupItem
Static TempEi As CSupItem

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CSupItem
      End If
      Set GetSupItem = TempEi
   Else
      Set GetSupItem = Ei
   End If
End Function
'Public Function GetAuthenPO_Approve(m_TempCol As Collection, UserName As String, PO_Price As Double) As Integer
'On Error Resume Next
'Dim Data As CAuthenPO
'GetAuthenPO_Approve = 0
'If m_TempCol.Count = 0 Then 'แสดงว่า ไม่มีผู้อนุมัติ
'   GetAuthenPO_Approve = 1
'   Exit Function
'End If
'For Each Data In m_TempCol
'    If (Data.AUTHEN_USER_NAME = UserName) And (Data.AUTHEN_PO_FROM < PO_Price) And (Data.AUTHEN_PO_TO > PO_Price) Then
'         GetAuthenPO_Approve = 2
'         Exit Function
'    End If
' Next Data
'End Function
'Public Function GetAuthenPO_Verify(m_TempCol As Collection, UserName As String, PO_Price As Double) As Integer
'On Error Resume Next
'Dim Data As CAuthenPO
'GetAuthenPO_Verify = 0
'If m_TempCol.Count = 0 Then 'แสดงว่า ไม่มีผู้ตรวจสอบ
'   GetAuthenPO_Verify = 1
'   Exit Function
'End If
'For Each Data In m_TempCol
'    If (Data.AUTHEN_USER_NAME = UserName) And (Data.AUTHEN_PO_FROM < PO_Price) And (Data.AUTHEN_PO_TO > PO_Price) Then
'         GetAuthenPO_Verify = 2
'         Exit Function
'    End If
' Next Data
'End Function
Public Function GetAuthenPO(m_TempCol As Collection, UserName As String, PO_Price As Double) As Integer
On Error Resume Next
Dim data As CAuthenPO
   GetAuthenPO = 0
   For Each data In m_TempCol
      If (data.AUTHEN_USER_NAME = UserName) Then
         GetAuthenPO = 2
         Exit Function
      End If
   Next data
End Function
Public Function WOY(MyDate As Date) As Integer    ' Week Of Year
  WOY = Format(MyDate, "ww", vbMonday, vbFirstFourDays)
  If WOY > 52 Then
    If Format(MyDate + 7, "ww", vbMonday, vbFirstFourDays) = 2 Then WOY = 1
  End If
End Function

Public Function NoFlag2Null(Flag As String) As String
   If Flag = "N" Then
      NoFlag2Null = ""
   Else
      NoFlag2Null = Flag
   End If
End Function

Public Function GetCustomerAccountList(m_TempCol As Collection, TempKey As String) As CCustomerAccountList
On Error Resume Next
Dim Ei As CCustomerAccountList

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      Set GetCustomerAccountList = Nothing
   Else
      Set GetCustomerAccountList = Ei
   End If
End Function
Public Sub GenerateReceiptHeader(Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, Offset As Long)
Dim Amt As Double
Dim VatAmt As Double
Dim NetAmt As Double

   Vsp.FontBold = True
   Vsp.TableBorder = TempBorder
   
    Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = Offset + (3.7 * 567)
    
    Vsp.TableCell(tcColWidth, , 1) = "19.cm"
      
   Vsp.EndTable
'------------------------------------------------------------------------------------------------------------> OFFSET
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.6cm"
   Vsp.TableBorder = TempBorder
   Vsp.TableCell(tcColWidth, , 1) = "3cm"
   Vsp.TableCell(tcColWidth, , 2) = "9cm"
   Vsp.TableCell(tcColWidth, , 3) = "2cm"
   Vsp.TableCell(tcColWidth, , 4) = "5cm"
   
   Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
   Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME")
   
   Vsp.TableCell(tcAlign, 1, 4) = taCenterBottom
   Vsp.TableCell(tcText, 1, 4) = DateToStringExtEx2(mcolParam("DOCUMENT_DATE"))
   
   Vsp.EndTable
   Vsp.FontBold = False
      
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.6cm"
   Vsp.TableBorder = TempBorder
   Vsp.TableCell(tcColWidth, , 1) = "1.5cm"
   Vsp.TableCell(tcColWidth, , 2) = "12.5cm"
   Vsp.TableCell(tcColWidth, , 3) = "2cm"
   Vsp.TableCell(tcColWidth, , 4) = "3cm"
   
   Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
   Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_ADDRESS")
   
   Vsp.TableCell(tcAlign, 1, 4) = taRightBottom
   Vsp.TableCell(tcText, 1, 4) = ""
   
   Vsp.EndTable
   Vsp.FontBold = False
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "1cm"
   Vsp.TableBorder = TempBorder
   Vsp.TableCell(tcColWidth, , 1) = "19cm"
   Vsp.EndTable
   Vsp.FontBold = False
   
  
End Sub
Public Function GenerateReceiptBody(Vsp As VSPrinter, mcolParam As Collection, BD As CBillingDoc, TempBorder As TableBorderSettings)
Dim Poi As CReceiptItem
Dim I As Long
Dim J As Long
Dim Sum As Double
   
   I = 0
   J = 0
   Sum = 0
   
   For Each Poi In BD.ReceiptItems
      I = I + 1
      J = J + 1
                          
      If J > 6 Then
         J = 1
         Vsp.NewPage
      End If
                       
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.7cm"
      Vsp.TableBorder = TempBorder
                 
      Vsp.TableCell(tcColWidth, , 1) = "2.cm"
      Vsp.TableCell(tcColWidth, , 2) = "12cm"
      Vsp.TableCell(tcColWidth, , 3) = "4.1cm"
      Vsp.TableCell(tcColWidth, , 4) = "0.9cm"
                        
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = Poi.DOCUMENT_NO
      
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(Poi.DOCUMENT_DATE)
                 
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = Left(FormatNumber(Poi.PAID_AMOUNT), Len(FormatNumber(Poi.PAID_AMOUNT)) - 3)
                 
      Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 4) = Right(FormatNumber(Poi.PAID_AMOUNT), 2)
      
      Sum = Sum + Poi.PAID_AMOUNT
            
      Vsp.EndTable
   Next Poi
   
   For I = 1 To (6 - J)
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.7cm"
      Vsp.TableBorder = TempBorder
                    
      Vsp.TableCell(tcColWidth, , 1) = "2.cm"
      Vsp.TableCell(tcColWidth, , 2) = "12cm"
      Vsp.TableCell(tcColWidth, , 3) = "4.1cm"
      Vsp.TableCell(tcColWidth, , 4) = "0.9cm"
       
      Vsp.EndTable
   Next I
   
   Vsp.StartTable
      Vsp.TableCell(tcCols) = 4
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.95cm"
      Vsp.TableBorder = TempBorder
                 
      Vsp.TableCell(tcColWidth, , 1) = "2.cm"
      Vsp.TableCell(tcColWidth, , 2) = "11cm"
      Vsp.TableCell(tcColWidth, , 3) = "5.1cm"
      Vsp.TableCell(tcColWidth, , 4) = "0.9cm"
                        
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = ""
      
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = "(" & ThaiBaht(Sum) & ")"
                 
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = Left(FormatNumber(Sum), Len(FormatNumber(Sum)) - 3)
                 
      Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
      Vsp.TableCell(tcText, 1, 4) = Right(FormatNumber(Sum), 2)
                       
      Vsp.EndTable
End Function
Public Sub GenerateReceiptFooter(Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, BD As CBillingDoc)
Dim Amt As Double
Dim VatAmt As Double
Dim NetAmt As Double
Dim Ct As CCashTran
   
   Vsp.FontBold = True
   Vsp.TableBorder = TempBorder
   
   Vsp.FontSize = 14
   
   For Each Ct In BD.Payments
      If Ct.GetFieldValue("PAYMENT_TYPE") = 1 Then   'เงินสด
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 4
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.6cm"
         Vsp.TableBorder = TempBorder
         Vsp.TableCell(tcColWidth, , 1) = "2cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
         Vsp.TableCell(tcColWidth, , 3) = "2cm"
         Vsp.TableCell(tcColWidth, , 4) = "14.6cm"
         
         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 2) = "X"
         
         Vsp.EndTable
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = 2 Then   'เงินโอน
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 4
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.6cm"
         Vsp.TableBorder = TempBorder
         Vsp.TableCell(tcColWidth, , 1) = "2cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
         Vsp.TableCell(tcColWidth, , 3) = "2cm"
         Vsp.TableCell(tcColWidth, , 4) = "14.6cm"
         
         Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
         Vsp.TableCell(tcText, 1, 4) = "( X )  เงินโอน   " & "ธนาคาร " & Ct.GetFieldValue("BANK_NAME") & " " & Ct.GetFieldValue("BRANCH_NAME") & "  เลขที่บัญชี   " & Ct.GetFieldValue("ACCOUNT_NAME")
         Vsp.EndTable
         
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then   'เช็ค
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 4
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.6cm"
         Vsp.TableBorder = TempBorder
         Vsp.TableCell(tcColWidth, , 1) = "2cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
         Vsp.TableCell(tcColWidth, , 3) = "2cm"
         Vsp.TableCell(tcColWidth, , 4) = "14.6cm"
         Vsp.EndTable
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 6
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.6cm"
         Vsp.TableBorder = TempBorder
         Vsp.TableCell(tcColWidth, , 1) = "2cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
         Vsp.TableCell(tcColWidth, , 3) = "3.5cm"
         Vsp.TableCell(tcColWidth, , 4) = "5cm"
         Vsp.TableCell(tcColWidth, , 5) = "2.5cm"
         Vsp.TableCell(tcColWidth, , 6) = "5.4cm"
         
         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 2) = "X"
         
         Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
         Vsp.TableCell(tcText, 1, 4) = Ct.GetFieldValue("BANK_NAME")
         
         Vsp.TableCell(tcAlign, 1, 6) = taLeftBottom
         Vsp.TableCell(tcText, 1, 6) = Ct.GetFieldValue("BRANCH_NAME")
         
         Vsp.EndTable
      
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 6
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.6cm"
         Vsp.TableBorder = TempBorder
         Vsp.TableCell(tcColWidth, , 1) = "2cm"
         Vsp.TableCell(tcColWidth, , 2) = "0.4cm"
         Vsp.TableCell(tcColWidth, , 3) = "3.5cm"
         Vsp.TableCell(tcColWidth, , 4) = "5cm"
         Vsp.TableCell(tcColWidth, , 5) = "2.5cm"
         Vsp.TableCell(tcColWidth, , 6) = "5.4cm"
         
         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 2) = ""
         
         Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
         Vsp.TableCell(tcText, 1, 4) = Ct.GetFieldValue("CHEQUE_NO")
         
         Vsp.TableCell(tcAlign, 1, 6) = taLeftBottom
         Vsp.TableCell(tcText, 1, 6) = DateToStringExtEx2(Ct.GetFieldValue("CHEQUE_DATE"))
         
         Vsp.EndTable
      
      
      End If
   Next Ct
   
   Vsp.FontBold = False
      
End Sub
Public Function CheckHaveValue(OldCheckHaveValue As Boolean, Amt As Double) As Boolean
   If (Amt <> 0) Or OldCheckHaveValue Then
      CheckHaveValue = True
   End If
End Function
Public Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim I As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For I = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next I
    WhereStr = WhereStr & ")"
    GenerateSearchLike = WhereStr
End Function
Public Function ChkLen(str As String, Size As Integer) As Boolean
 ChkLen = False
  If IsNumeric(str) Then
     If Len(str) < Size Then
          ChkLen = False
     Else
       ChkLen = True
     End If
End If
End Function
'Public Function getIPAddress() As String
'Dim WMI     As Object
'Dim qryWMI  As Object
'Dim Item    As Variant
'    Set WMI = GetObject("winmgmts:\\.\root\cimv2")
'    Set qryWMI = WMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration  WHERE IPEnabled = True")
'    For Each Item In qryWMI
'      getIPAddress = Item.IpAddress(0)
'    Next
'    Set WMI = Nothing
'    Set qryWMI = Nothing
'End Function
Public Sub CalculateSupItemSumComeIn(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim Rs As ADODB.Recordset
Dim ItemCount As Long
Dim I As Long
Dim J As Long
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim TempStr As String
Dim Sup As CSupItem
Dim PrevKey1 As String
Dim TempRo As CSupItem
Dim tempSupItem As CSupItem
Dim TempComeIn  As CSupItem
Dim TempData As CSumComeIn
Dim RoColl As Collection
Dim ComeInColl As Collection  'Query จำนวนวัตถุดิบที่เข้ามาก่อนหน้าวันที่ต้องการออก report 1 วัน  เพื่อทำยอดยกมา และ ยอดสะสม
Dim collSumSupItemByRo As Collection
Dim CarryForward As Double
Dim TmpValue As Double
Dim currentKey1  As String

   Set RoColl = New Collection
   Set ComeInColl = New Collection
   Set collSumSupItemByRo = New Collection
    
  Call LoadSupItemComeIn(Nothing, RoColl, ToDate, ToDate)
  Call LoadSupItemComeIn(Nothing, ComeInColl, -1, DateAdd("d", -1, ToDate))
   
   Call LoadSupItemPartItemByRo(Nothing, collSumSupItemByRo, -1, DateAdd("D", -1, ToDate), , "(100,101,102,103)")
   Set Rs = New ADODB.Recordset
   
   For J = 1 To UBound(Total1)
      Total1(J) = 0
   Next J
   
   Set Sup = New CSupItem
   
   
   Sup.SUP_ITEM_ID = -1
'   Sup.SUPPLIER_CODE = SupplierCode
  Sup.OrderBy = 7
   
   Sup.DOCUMENT_TYPE_SET = "(1000,1001,1002,1003)"
   Sup.CLOSE_FLAG = "N"
   Sup.PO_APPROVED_FLAG = "Y"
   Sup.TO_DATE = ToDate
   Call Sup.QueryData(112, Rs, ItemCount)
   I = 0
   
   PrevKey1 = ""
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
      
      Call Sup.PopulateFromRS(112, Rs)
      
      Set tempSupItem = GetObject("CSupItem", collSumSupItemByRo, Trim(Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
      If (Sup.TX_AMOUNT > tempSupItem.TX_AMOUNT) Then
         currentKey1 = Sup.PART_NO
         
         If PrevKey1 <> currentKey1 And I > 0 Then
            Set TempData = New CSumComeIn
            TempData.PART_NO = PrevKey1
            TempData.SUM_TX_AMOUNT = Total1(7)
            TempData.SUM_ACTUAL_UNIT_PRICE = Total1(8)
            If Not (C Is Nothing) Then
            End If
      
            If Not (Cl Is Nothing) Then
               Call Cl.add(TempData, Trim(PrevKey1))
            End If
            Set TempData = Nothing
            For J = 1 To UBound(Total1)
                Total1(J) = 0
            Next J
         ElseIf I = 0 Then
         
         End If
         
         I = I + 1
         PrevKey1 = Sup.PART_NO

         TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE, 2)

         TempStr = FormatNumberToNull(Sup.TX_AMOUNT, 2)
         Total1(5) = Total1(5) + Sup.TX_AMOUNT
 
         TempStr = Sup.UNIT_NAME
         Set TempRo = GetObject("CSupItem", RoColl, Trim(Sup.SUPPLIER_ID & "-" & Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
         Set TempComeIn = GetObject("CSupItem", ComeInColl, Trim(Sup.SUPPLIER_ID & "-" & Sup.BILLING_DOC_ID & "-" & Sup.PART_ITEM_ID))
         
' ค้างส่ง
         TmpValue = Sup.TX_AMOUNT - (TempRo.TX_AMOUNT + TempComeIn.TX_AMOUNT)
         TempStr = FormatNumberToNull(TmpValue, 2)
         Total1(7) = Total1(7) + TmpValue
         
         TempStr = FormatNumberToNull(Sup.ACTUAL_UNIT_PRICE * TmpValue)
         Total1(8) = Total1(8) + (Sup.ACTUAL_UNIT_PRICE * TmpValue)

      End If
      
      Rs.MoveNext
   Wend
    
   Set RoColl = Nothing
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set Sup = Nothing
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub CalculateSumInventoryAccountAndMove(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LOCATION_ID As Long = -1, Optional PART_TYPE As Long = -1, Optional PART_GROUP As Long = -1)
On Error GoTo ErrorHandler
Dim RName As String
Dim cData As CLotItem
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Rs As ADODB.Recordset
Dim TempData As CSumInventoryAccount

Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim Total1(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double
Dim InventoryBals1 As Collection
Dim Li As CLotItem
Dim TempLi1 As CLotItem
Dim TempLi2 As CLotItem
Dim Sum1 As Double
Dim Count1 As Double
Dim BalanceAccums As Collection
Dim LeftAmount1 As Double
Dim LeftAmount2 As Double
Dim TempLi As CLotItem
Dim Tot1 As Double
Dim Tot2 As Double
Dim TxValue As Double
Dim TempAmt As Double
Dim TempValue As Double


Dim m_PartTxtypes As Collection
Set m_PartTxtypes = New Collection
Dim m_PartTxtypeBas As Collection
Set m_PartTxtypeBas = New Collection

Dim BalanceLi As CLotItem
Dim TempLiBa1 As CLotItem
Dim TempLiBa2 As CLotItem

Set Rs = New ADODB.Recordset
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   'Call LoadPartTxTypeDocTypeAmount(Nothing, m_PartTxtypes, FromDate, ToDate)
   Call LoadPartTxTypeDocTypeAmount(Nothing, m_PartTxtypes, FromDate, ToDate, , LOCATION_ID, PART_TYPE, PART_GROUP)
   
   Set BalanceAccums = New Collection
   Set InventoryBals1 = New Collection
   If FromDate > 0 Then
      Dim MonthlyAccums  As Collection
      Dim YYYYMM As String
      Dim firstDate As Date
      Dim lastDate As Date
      Set MonthlyAccums = New Collection
      Call GetFirstLastDate(FromDate, firstDate, lastDate)
      YYYYMM = Format(Year(DateAdd("D", -1, firstDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, firstDate)), "00")
      Call LoadMonthlyBalancePartItem(Nothing, MonthlyAccums, YYYYMM, LOCATION_ID)
      Call glbDaily.CopyMonthlyAccumPartItem(MonthlyAccums, InventoryBals1)
      Set MonthlyAccums = Nothing
      
      If (firstDate <> FromDate) Then
         Call LoadPartTxTypeAmount(Nothing, m_PartTxtypeBas, firstDate, DateAdd("D", -1, FromDate), , LOCATION_ID)
      End If
   End If
   Set BalanceAccums = Nothing
   
   Dim Ma As CMonthlyAccum
   Set Ma = New CMonthlyAccum
   Ma.PART_ITEM_ID = -1
   Ma.TO_YYYYMM = Format(Year(ToDate), "0000") & "-" & Format(Month(ToDate), "00")
   'Ma.YYYYMM = YYYYMM
   Ma.LOCATION_ID = LOCATION_ID
   Ma.PART_TYPE = PART_TYPE
   Ma.PART_GROUP = PART_GROUP
   Ma.OrderBy = 1
   Call Ma.QueryData(4, Rs, iCount, False)
    I = 0

 If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

      Set Ma = New CMonthlyAccum
      While Not Rs.EOF
         Set TempData = New CSumInventoryAccount
         
          I = I + 1
         Call Ma.PopulateFromRS(4, Rs)
         Tot1 = 0
         Tot2 = 0
         TxValue = 0
         Set BalanceLi = GetLotItem(InventoryBals1, Trim(str(Ma.PART_ITEM_ID)))
         Set TempLiBa1 = GetLotItem(m_PartTxtypeBas, Ma.PART_ITEM_ID & "-" & "I")
         Set TempLiBa2 = GetLotItem(m_PartTxtypeBas, Ma.PART_ITEM_ID & "-" & "E")
         
'          If Ma.PART_NO = "2016" Then
'              ''Debug.Print Ma.PART_NO
'         End If
         ' ยกมา ====
         LeftAmount1 = BalanceLi.NEW_AMOUNT + TempLiBa1.TX_AMOUNT - TempLiBa2.TX_AMOUNT
         Total1(4) = Total1(4) + LeftAmount1
         Amt = LeftAmount1
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         
         Sum1 = BalanceLi.TOTAL_INCLUDE_PRICE + TempLiBa1.TOTAL_INCLUDE_PRICE - TempLiBa2.TOTAL_INCLUDE_PRICE
         Tot2 = Tot2 + Sum1
         '====
         
          '=== ซื้อเข้า รับเข้าทั่วไป
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-1")  'ใบรับเข้าวัตถุดิบ
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-23")  'ใบรับเข้าทั่วไป
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-18")  'รับคืน
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-19")  'ใบรับเข้าวัสดุอุปกรณ์
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-20")  'ใบรับเข้าจ่ายออกวัสดุอุปกรณ์
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1

          '==== รับจากการผลิต
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-12")  'รับจากการผลิต
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-13")  'รับจากการผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-14")  'รับจากการผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         
         '==== รับจากการผลิต
         
       '=== รับจากการโอน
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-3")  'โอนเข้า
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-22")  'โอนเข้า
         TempLi.TX_AMOUNT = TempLi.TX_AMOUNT + TempLi2.TX_AMOUNT
         TempLi.TOTAL_INCLUDE_PRICE = TempLi.TOTAL_INCLUDE_PRICE + TempLi2.TOTAL_INCLUDE_PRICE
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         '=== รับจากการโอน

         '=== รับเข้าจากการปรับยอดจากการชั่งตวง
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-5")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         
         '=== รับเข้าจากการปรับยอด

        '=== รับเข้าจากการตรวจนับ
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-4")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt

          '=== เบิก ผลิต
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-2")  'เบิก
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-12")  'ผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-13")  'ผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-14")  'ผลิต
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-20")  'รับเข้าจ่ายออก
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1

         '=== เบิก ผลิต
         
         '=== โอน
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-3")  'โอน
         Set TempLi2 = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-22")  'โอน
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         Amt = Amt + TempLi2.TX_AMOUNT
         Sum1 = Sum1 + TempLi2.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1
         
         '=== ปรับยอดจากการชั่งตวง
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-5")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1

         '=== ขายเชื่อ ขายสด
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-10")
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-21")
         Amt = Amt + TempLi.TX_AMOUNT
         Sum1 = Sum1 + TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1
         
         '=== ปรับยอดจากการตรวจนับ
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-4")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt

         '=== คงเหลือบัญชี
         ' Sum1 = BalanceLi.TOTAL_INCLUDE_PRICE + TempLiBa1.TOTAL_INCLUDE_PRICE - TempLiBa2.TOTAL_INCLUDE_PRICE
         TempData.SUM_INV_ACCOUNT = Tot1
         TempData.AVR_UNIT_ACCOUNT = MyDiffEx(Tot2, Tot1)
   
         '=== คงเหลือ Physical
         
         TempAmt = 0
         TempValue = 0
         
         '=== รับเข้าจากการตรวจนับ
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-I-4")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 + Amt
         Tot2 = Tot2 + Sum1
         
         TempAmt = TempAmt + Amt
         TempValue = TempValue + Sum1
         
         
         '=== ปรับยอดจากการตรวจนับ
         Set TempLi = GetLotItem(m_PartTxtypes, Ma.PART_ITEM_ID & "-E-4")  'ปรับยอด
         Amt = TempLi.TX_AMOUNT
         Sum1 = TempLi.TOTAL_INCLUDE_PRICE
         TxValue = TxValue + Amt
         Tot1 = Tot1 - Amt
         Tot2 = Tot2 - Sum1
         
         TempAmt = TempAmt - Amt
         TempValue = TempValue - Sum1
         
         '=== Diff
         TempData.SUM_INV_PHYSICAL = Tot1
         TempData.AVR_UNIT_PHYSICAL = MyDiffEx(Tot2, Tot1)
         If TxValue <> 0 Then
         Else
            I = I - 1
         End If
      TempData.PART_NO = Ma.PART_NO
      If Not (C Is Nothing) Then
      End If
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Ma.PART_NO))
      End If
      Set TempData = Nothing
      Rs.MoveNext
      Wend
      Set Li = Nothing
      Set Pi = Nothing
      Set cData = Nothing
      Set InventoryBals1 = Nothing
   
   'genDoc = True
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub CalculateSumInventoryAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LOCATION_ID As Long = -1, Optional PART_TYPE As Long = -1, Optional PART_GROUP As Long = -1)
On Error GoTo ErrorHandler
Dim RName As String
Dim cData As CLotItem
Dim I As Long
Dim J As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSumInventoryAccount
Dim Total1(100) As Double
Dim iCount As Long
Dim Amt As Double
Dim InventoryBals1 As Collection
Dim TempLi1 As CLotItem
Dim TempLi2 As CLotItem
Dim Sum1 As Double
Dim Count1 As Double
Dim BalanceAccums As Collection
Dim LeftAmount1 As Double
Dim LeftAmount2 As Double
Dim TempLi As CLotItem
Dim Tot1 As Double
Dim Tot2 As Double
Dim TxValue As Double
Dim TempAmt As Double
Dim TempValue As Double

Dim m_PartTxtypes As Collection
Set m_PartTxtypes = New Collection
Dim m_PartTxtypeBas As Collection
Set m_PartTxtypeBas = New Collection

Dim BalanceLi As CLotItem
Dim TempLiBa1 As CLotItem
Dim TempLiBa2 As CLotItem

Set Rs = New ADODB.Recordset
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Call LoadPartTxTypeDocTypeAmount(Nothing, m_PartTxtypes, FromDate, ToDate, , LOCATION_ID, PART_TYPE, PART_GROUP)
   
   Set BalanceAccums = New Collection
   Set InventoryBals1 = New Collection
   If FromDate > 0 Then
      Dim MonthlyAccums  As Collection
      Dim YYYYMM As String
      Dim firstDate As Date
      Dim lastDate As Date
      Set MonthlyAccums = New Collection
      Call GetFirstLastDate(FromDate, firstDate, lastDate)
      YYYYMM = Format(Year(DateAdd("D", -1, firstDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, firstDate)), "00")
      Call LoadMonthlyBalancePartItem(Nothing, MonthlyAccums, YYYYMM, LOCATION_ID)
      Call glbDaily.CopyMonthlyAccumPartItem(MonthlyAccums, InventoryBals1)
      Set MonthlyAccums = Nothing
      
      If (firstDate <> FromDate) Then
         Call LoadPartTxTypeAmount(Nothing, m_PartTxtypeBas, firstDate, DateAdd("D", -1, FromDate), , LOCATION_ID)
      End If
   End If
   Set BalanceAccums = Nothing
   
   Dim Ma As CMonthlyAccum
   Set Ma = New CMonthlyAccum
   Ma.PART_ITEM_ID = -1
   Ma.TO_YYYYMM = Format(Year(ToDate), "0000") & "-" & Format(Month(ToDate), "00")
   'Ma.YYYYMM = YYYYMM
   Ma.LOCATION_ID = LOCATION_ID
   Ma.PART_TYPE = PART_TYPE
   Ma.PART_GROUP = PART_GROUP
   Ma.OrderBy = 1
   Call Ma.QueryData(4, Rs, iCount, False)
    I = 0

 If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

      Set Ma = New CMonthlyAccum
      While Not Rs.EOF
         Set TempData = New CSumInventoryAccount
         
          I = I + 1
         Call Ma.PopulateFromRS(4, Rs)
         Tot1 = 0
         Tot2 = 0
         Sum1 = 0
         TxValue = 0
         Set BalanceLi = GetLotItem(InventoryBals1, Trim(str(Ma.PART_ITEM_ID)))
         Set TempLiBa1 = GetLotItem(m_PartTxtypeBas, Ma.PART_ITEM_ID & "-" & "I")
         Set TempLiBa2 = GetLotItem(m_PartTxtypeBas, Ma.PART_ITEM_ID & "-" & "E")
         
         ' ยกมา ====
         LeftAmount1 = BalanceLi.NEW_AMOUNT + TempLiBa1.TX_AMOUNT - TempLiBa2.TX_AMOUNT
         Sum1 = BalanceLi.TOTAL_INCLUDE_PRICE + TempLiBa1.TOTAL_INCLUDE_PRICE - TempLiBa2.TOTAL_INCLUDE_PRICE
         '=== คงเหลือบัญชี
         TempData.SUM_INV_ACCOUNT = LeftAmount1 'Tot1
         TempData.AVR_UNIT_ACCOUNT = FormatNumber(MyDiff(Sum1, LeftAmount1), 2)

         If TxValue <> 0 Then
         Else
            I = I - 1
         End If
      TempData.PART_NO = Ma.PART_NO
      If Not (C Is Nothing) Then
      End If
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Ma.PART_NO))
      End If
      Set TempData = Nothing
      Rs.MoveNext
      Wend
      Set cData = Nothing
      Set InventoryBals1 = Nothing
   
   'genDoc = True
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Function getBeginDay(dayName As String) As Date
Dim dayInt As Integer
Dim Date_Now As Integer
If dayName = "SUN" Then
   dayInt = 1
ElseIf dayName = "MON" Then
   dayInt = 2
ElseIf dayName = "TUE" Then
   dayInt = 3
ElseIf dayName = "WED" Then
   dayInt = 4
ElseIf dayName = "THU" Then
   dayInt = 5
ElseIf dayName = "FRI" Then
   dayInt = 6
ElseIf dayName = "SAT" Then
   dayInt = 7
End If
Date_Now = Format(Now, "w")
If Date_Now >= dayInt Then
   Date_Now = Format(Now, "w") - dayInt
Else
   Date_Now = 8 - dayInt
End If
getBeginDay = Now - Date_Now
End Function
Function ConcateSting(msgO As String, msgN As String) As String
Dim Arr() As String
Arr = Split(msgO, ",")

   If UBound(Arr) = -1 Then
      ConcateSting = msgN
   Else
      ConcateSting = msgO & "," & msgN
   End If
End Function
Public Function CapsLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyCapital)
    CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function
Public Function CheckCapsLock() As String
If CapsLockOn = True Then
        CheckCapsLock = "Capslock is on."
   Else
         CheckCapsLock = "Capslock is off."
   End If
End Function
Public Function GetTotalAmountPallet(Cl As Collection) As Double
Dim PD As CPalletDoc
Dim SumAmount As Double
'   SumAmount = 0
'   For Each PD In Cl
'      If PD.Flag <> "D" Then
'            SumAmount = SumAmount + PD.PALLET_CAP_LAST
'      End If
'   Next PD
'   GetTotalAmountPallet = SumAmount
End Function
Public Function GetTotalAmount(Cl As Collection) As Double
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim SumAmount As Double
   SumAmount = 0
   For Each LTD In Cl
      For Each PD In LTD.C_PalletDoc
         If PD.Flag <> "D" Then
            SumAmount = SumAmount + PD.CAPACITY_AMOUNT
         End If
      Next PD
   Next LTD
   GetTotalAmount = SumAmount
End Function
Public Function GetTotalAmount2(Cl As Collection, Optional DocumentType As Long = -1, Optional PartItemID As Long = -1) As Double
Dim m_CollPallet As Collection
Dim LTD As CLotDoc
Dim PD As CPalletDoc
Dim SumAmount As Double
   SumAmount = 0
   For Each LTD In Cl
      Set m_CollPallet = New Collection
      Call LoadPalletDocAmount(Nothing, m_CollPallet, LTD.LOT_ID, 2, , 2, "I", , , LTD.LOT_DOC_ID, LTD.HEAD_PACK_NO, LTD.LOT_ITEM_WH_ID, DocumentType, PartItemID)
      SumAmount = SumAmount + GetTotalAmountPallet(m_CollPallet) 'FormatNumber(GetTotalAmountPallet(m_CollPallet), 0)
      Set m_CollPallet = Nothing
   Next LTD
   GetTotalAmount2 = SumAmount
End Function
Public Function getFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    If dtmDate = 0 Then
        dtmDate = Date
    End If
    getFirstDayInMonth = DateSerial(Year(dtmDate), Month(dtmDate), 1)
End Function
Public Function getLastDayInMonth(Optional dtmDate As Date = 0) As Date
    If dtmDate = 0 Then
        dtmDate = Date
    End If
    getLastDayInMonth = DateSerial(Year(dtmDate), Month(dtmDate) + 1, 0)
End Function
'Public Function CreateLotAuto(Optional StartDate As Date = -1, Optional LotNoNew As Long = 0, Optional cLot As Collection = Nothing) As Long
'Dim LT As cLot
'Dim No As String
'Dim IsOK As Boolean
'Dim Rs As ADODB.Recordset
'Set Rs = New ADODB.Recordset
'
'      If StartDate = -1 Then
'         glbErrorLog.LocalErrorMsg = MapText("เอกสารนี้ไม่มีข้อมูลวันที่ผลิต กรุณาตรวจสอบ")
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
'      If LotNoNew = 0 Then
'         glbErrorLog.LocalErrorMsg = MapText("กรุณ ระบุ เลทที่ Lot การผลิต")
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
'
'      Set LT = New cLot
'
'      No = "LG" & Right(Format(Year(StartDate) + 543, "0000"), 2) & Format(StartDate, "mm") & Format(StartDate, "dd")
'      LT.LOT_NO = No & Format(LotNoNew, "000")
'      LT.LOT_DATE = StartDate
'
'      If CheckUniqueNs(LOT_UNIQUE, LT.LOT_NO, -1) Then 'ถ้ายังไม่มี
'        LT.AddEditMode = SHOW_ADD
'         Call LT.AddEditData
'         CreateLotAuto = LT.LOT_ID
'      Else
'         Call LT.QueryData(1, Rs, 0)
'         Set LT = Nothing
'         Set LT = New cLot
'         Call LT.PopulateFromRS(1, Rs)
'         CreateLotAuto = LT.LOT_ID
'      End If
'
'      Set LT = Nothing
'End Function
Public Function CreateLotAuto(Optional StartDate As Date = -1, Optional LotNoNew As Long = 0, Optional cLot As Collection = Nothing) As Long
Dim Lt As cLot
Dim No As String
Dim IsOK As Boolean
Dim SearchLotNo As cLot
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

      If StartDate = -1 Then
         glbErrorLog.LocalErrorMsg = MapText("เอกสารนี้ไม่มีข้อมูลวันที่ผลิต กรุณาตรวจสอบ")
         glbErrorLog.ShowUserError
         Exit Function
      End If
      If LotNoNew = 0 Then
         glbErrorLog.LocalErrorMsg = MapText("กรุณ ระบุ เลทที่ Lot การผลิต")
         glbErrorLog.ShowUserError
         Exit Function
      End If
      
      Set Lt = New cLot
      
      No = "LG" & Right(Format(Year(StartDate) + 543, "0000"), 2) & Format(StartDate, "mm") & Format(StartDate, "dd")
      Lt.LOT_NO = No & Format(LotNoNew, "000")
      Lt.LOT_DATE = StartDate
      
       Set SearchLotNo = GetObject("CLot", cLot, Trim(Lt.LOT_NO), False)
      If SearchLotNo Is Nothing Then 'ถ้ายังไม่มี
         Lt.AddEditMode = SHOW_ADD
         Call Lt.AddEditData
         CreateLotAuto = Lt.LOT_ID
      Else
         Call Lt.QueryData(1, Rs, 0)
         Set Lt = Nothing
         Set Lt = New cLot
         Call Lt.PopulateFromRS(1, Rs)
         CreateLotAuto = Lt.LOT_ID
      End If
      
      Set Lt = Nothing
End Function
Public Function CheckIntAscii(KeyAscii As Integer) As Integer
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8) Then
   KeyAscii = 0
  End If
  CheckIntAscii = KeyAscii
End Function
Public Function VerifyLockDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_DATE And uctlDate <= glbLockDate.TO_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_DATE And oldDate <= glbLockDate.TO_DATE))) Then
      VerifyLockDate = True
   Else
      VerifyLockDate = False
   End If
End Function
Public Function VerifyLockInventoryDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_INVENTORY_DATE And uctlDate <= glbLockDate.TO_INVENTORY_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_INVENTORY_DATE And oldDate <= glbLockDate.TO_INVENTORY_DATE))) Then
      VerifyLockInventoryDate = True
   Else
      VerifyLockInventoryDate = False
   End If
End Function
Public Function VerifyLockInvoiceDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_INVOICE_DATE And uctlDate <= glbLockDate.TO_INVOICE_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_INVOICE_DATE And oldDate <= glbLockDate.TO_INVOICE_DATE))) Then
      VerifyLockInvoiceDate = True
   Else
      VerifyLockInvoiceDate = False
   End If
End Function
Public Function VerifyLockReceiptDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_RECEIPT_DATE And uctlDate <= glbLockDate.TO_RECEIPT_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_RECEIPT_DATE And oldDate <= glbLockDate.TO_RECEIPT_DATE))) Then
      VerifyLockReceiptDate = True
   Else
      VerifyLockReceiptDate = False
   End If
End Function

Public Function InternalDateToDateExGrid(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long
   
Dim TempDate As Date


   
   If Len(IntDate) < 8 Then
      InternalDateToDateExGrid = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 7, 4)
   MStr = Mid(IntDate, 4, 2)
   DStr = Mid(IntDate, 1, 2)
   
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
   Y = Val(YStr) - 543
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateExGrid = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function
Public Sub getLockDate()
Dim m_Rs  As ADODB.Recordset
Dim iCount As Long
Set m_Rs = New ADODB.Recordset
   glbLockDate.LOCK_DATE_ID = -1
   glbLockDate.LOCK_TYPE = 1
   Call glbLockDate.QueryData(1, m_Rs, iCount)
   If Not m_Rs.EOF Then
      Call glbLockDate.PopulateFromRS(1, m_Rs)
   End If
End Sub
Public Sub EditStatusFlagInInventoryWHDoc(m_BillingDoc As CBillingDoc, Optional InvDelFlag As String = "N")
   Dim IvWhd As CInventoryWHDoc
   Dim Di As CDoItem
   Set IvWhd = New CInventoryWHDoc
   
   For Each Di In m_BillingDoc.DoItems
      If Di.Flag = "D" Then 'ถ้าใน list ของ DoItem ของ ใบ Inv มีการ ลบออกเพียงตัวเดียว ก็จะทำให้ ใบ So ไม่สมบูรณ์
         m_BillingDoc.LOAD_FLAG = "" 'เพื่อให้ไป update ในใบขึ้นอาหารว่า ยัง ออกใบส่งของไม่สำเร็จ
      End If
   Next Di
   
   IvWhd.INVENTORY_WH_DOC_ID = m_BillingDoc.INVENTORY_WH_DOC_ID
   If IvWhd.INVENTORY_WH_DOC_ID > 0 Then
      If InvDelFlag = "D" Then 'ถ้าเป็นลบใบส่งของ ก็ให้ไป update ใบขึ้นอาหาร ให้เป็น I ด้วย
         m_BillingDoc.LOAD_FLAG = ""
      End If
      If m_BillingDoc.LOAD_FLAG = "I" Or m_BillingDoc.LOAD_FLAG = "Y" Then
         IvWhd.LOAD_FLAG = "Y"
         IvWhd.SUCCESS_FLAG = "Y"
      Else
         IvWhd.LOAD_FLAG = "I"
         IvWhd.SUCCESS_FLAG = "N"
      End If
      m_BillingDoc.LOAD_FLAG = IvWhd.LOAD_FLAG
      m_BillingDoc.SUCCESS_FLAG = IvWhd.SUCCESS_FLAG
      m_BillingDoc.B_SUCCESS_FLAG = IvWhd.SUCCESS_FLAG
        If Not IvWhd.UpdateSuccessFlag Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call glbDaily.RollbackTransaction
            Exit Sub
         End If
   End If
End Sub
Public Sub EditStatusFlagInBillingDoc(m_BillingDoc As CBillingDoc)
   Dim BD As CBillingDoc
   Dim Di As CDoItem
   Dim PrevKey As Long
   Dim CountUpdateC As Long
   Dim CountUpdateY As Long
   PrevKey = 0

   For Each Di In m_BillingDoc.DoItems
      If Di.Flag = "D" Then 'ถ้าใน list ของ DoItem ของ ใบ Inv มีการ ลบออกเพียงตัวเดียว ก็จะทำให้ ใบ So ไม่สมบูรณ์
         m_BillingDoc.LOAD_FLAG = "" 'เพื่อให้ไป update ในใบขึ้นอาหารว่า ยัง ออกใบส่งของไม่สำเร็จ
      End If
   Next Di
   
   
   For Each Di In m_BillingDoc.DoItems
      If PrevKey <> Di.BILLING_DOC_SO_ID Then
         Set BD = New CBillingDoc
         BD.BILLING_DOC_ID = Di.BILLING_DOC_SO_ID
         PrevKey = Di.BILLING_DOC_SO_ID
            If m_BillingDoc.B_SUCCESS_FLAG = "Y" Then
                Call BD.UpdateSuccessFlag("C") 'update สถานะ ของใบ so ว่า Compleate แล้ว
                CountUpdateC = CountUpdateC + 1
            Else
                Call BD.UpdateSuccessFlag("Y") 'update สถานะ ของใบ so ว่า load ไปใช้ในใบขึ้นอาหารเรียบร้อยแล้ว
                CountUpdateY = CountUpdateY + 1
             End If
       End If
    Next Di
    Set Di = Nothing
   Set BD = Nothing
   If CountUpdateC > 0 And CountUpdateY = 0 Then 'ถ้า ใบ So ทุกใบ Compleat หมดแล้วก็ ให้เปลี่ยน สถานะในใบส่งของเป็น Compleate ด้วย
      m_BillingDoc.SUCCESS_FLAG = "C" 'ใบส่งของออกใบส่งของเรียบร้อย
   Else
      m_BillingDoc.SUCCESS_FLAG = "N" 'ใบส่งของรอดึงใบขึ้นอาหาร
   End If
End Sub

Public Function CheckLastVersionProgram(LastVerPro As String) As String
On Error GoTo ErrorHandler
Dim ErrorObj As clsErrorLog
Dim m_Rs  As ADODB.Recordset
Dim iCount As Long
Dim VerPro As CVersionProgram
Set VerPro = New CVersionProgram
Set m_Rs = New ADODB.Recordset

VerPro.VERSION_ID = 1
Call VerPro.QueryData(m_Rs, iCount)
Set VerPro = Nothing
If Not m_Rs.EOF Then
   Set VerPro = New CVersionProgram
    Call VerPro.PopulateFromRS(1, m_Rs)
    If LastVerPro > VerPro.VERSION_NAME Then
       VerPro.AddEditMode = SHOW_EDIT
       VerPro.VERSION_ID = 1
       VerPro.VERSION_NAME = Trim(LastVerPro)
       Call VerPro.AddEditData
    End If
     CheckLastVersionProgram = VerPro.VERSION_NAME
 End If
Exit Function

ErrorHandler:
   Set ErrorObj = New clsErrorLog
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
End Function
Public Function calExWorksPrice(Pi As CPartItem, DELIVERY_CUS_ITEM_ID As Long, CUSTOMER_ID As Long, PRICE_THINK_TYPE As Long, m_ExWorkPricesItem As Collection, m_ExDeliveryCostItem As Collection, m_Customers As Collection, ByRef TempData As Collection, m_ExPromotionPartItem As Collection, m_ExPromotionDlcItem As Collection, ByRef RatePromotionPart As Double, ByRef RatePromotionDlc As Double) As Double
Dim RateWotkPrice As Double
Dim RateDeliveryCost As Double
Dim RateOther As Double 'Pi
'Dim RatePromotionPart As Double
'Dim RatePromotionDlc As Double


Dim RateWotkPriceFlag As Boolean
Dim RateDeliveryCostFlag As Boolean
Dim RateOtherFlag As Boolean
Dim RatePromotionPartFlag As Boolean
Dim RatePromotionDlcFlag As Boolean


Dim TempD As CExWorksPrice
Dim TempD2 As CCustomer

Dim tempRate As CDoItem
Set tempRate = New CDoItem

calExWorksPrice = 0
Set TempData = New Collection
'calExDeliveryCost
      If Pi.PART_TYPE = 10 Or Pi.PART_TYPE = 21 Then 'จะเข้าทำเฉพาะ Bag กับ Bulk เท่านั้น
       'คิดราคาสินค้า
         Set TempD = GetObject("CExWorksPrice", m_ExWorkPricesItem, Trim(str(Pi.PART_ITEM_ID)), False)  'ค้นหาราคาสินค้าตาม id สินค้า
         If Not TempD Is Nothing Then
            RateWotkPrice = TempD.PACKAGE_RATE
            
            tempRate.EX_WORKS_PRICE_ITEM_ID = TempD.EX_WORKS_PRICE_ITEM_ID
            tempRate.PACKAGE_RATE = RateWotkPrice
            RateWotkPriceFlag = True
         Else
           RateWotkPriceFlag = False
         End If
         'จบคิดราคาสินค้า
         
         
         'คิดราคาค่าขนส่ง
         Dim EX_DELIVERY_COST_ITEM_ID As Long
         Call calExDeliveryCost(DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, Pi.WEIGHT_PER_PACK, Pi.PART_TYPE, PRICE_THINK_TYPE, m_ExDeliveryCostItem, m_Customers, RateDeliveryCost, RateDeliveryCostFlag, EX_DELIVERY_COST_ITEM_ID)
         
         tempRate.EX_DELIVERY_COST_ITEM_ID = EX_DELIVERY_COST_ITEM_ID
         tempRate.RATE_CUSTOMER = RateDeliveryCost

          'จบการคิดราคาค่าขนส่ง
          
           'คิดโปรโมชั่นสินค้า
          'm_ExPromotionPartItem
         Set TempD = GetObject("CExWorksPrice", m_ExPromotionPartItem, Trim(str(CUSTOMER_ID)) & "-" & Trim(str(Pi.PART_ITEM_ID)), False)  'ค้นหาราคาสินค้าตาม id สินค้า
         If Not TempD Is Nothing Then
            RatePromotionPart = TempD.DISCOUNT_AMOUNT

            tempRate.EX_PROMOTION_PART_ITEM_ID = TempD.EX_PROMOTION_PART_ITEM_ID
            tempRate.DISCOUNT_AMOUNT_PART = RatePromotionPart
            RatePromotionPartFlag = True
         Else
           RatePromotionPartFlag = False
         End If
          'จบคิดโปรโมชั่นสินค้า
          
         'คิดโปรโมชั่นค่าขนส่ง
          'm_ExPromotionPartItem
         Dim EX_PROMOTION_DLC_ITEM_ID As Long
         Call calExPromotionDls(DELIVERY_CUS_ITEM_ID, CUSTOMER_ID, Pi.WEIGHT_PER_PACK, Pi.PART_TYPE, PRICE_THINK_TYPE, m_ExPromotionDlcItem, m_Customers, RatePromotionDlc, RatePromotionDlcFlag, EX_PROMOTION_DLC_ITEM_ID)
         
         tempRate.EX_PROMOTION_DLC_ITEM_ID = EX_PROMOTION_DLC_ITEM_ID
         tempRate.DISCOUNT_AMOUNT_DLC = RatePromotionDlc
          'จบคิดโปรโมชั่นค่าขนส่ง
         
         Set TempD2 = GetObject("CCustomer", m_Customers, Trim(str(CUSTOMER_ID)), False)
         If Not TempD2 Is Nothing Then
           If Pi.PART_TYPE = 10 Then
               RateOther = TempD2.PRO_COMMISSION_BAG + TempD2.PRO_CHEER_BAG + TempD2.PRO_DST_BAG + TempD2.PRO_OTHER1_BAG + TempD2.PRO_OTHER2_BAG + TempD2.PRO_OTHER3_BAG
                 
               tempRate.PRO_COMMISSION_BAG = TempD2.PRO_COMMISSION_BAG
               tempRate.PRO_CHEER_BAG = TempD2.PRO_CHEER_BAG
               tempRate.PRO_DST_BAG = TempD2.PRO_DST_BAG
               tempRate.PRO_OTHER1_BAG = TempD2.PRO_OTHER1_BAG
               tempRate.PRO_OTHER2_BAG = TempD2.PRO_OTHER2_BAG
               tempRate.PRO_OTHER3_BAG = TempD2.PRO_OTHER3_BAG
               tempRate.PRO_COMMISSION_KG = 0
               tempRate.PRO_CHEER_KG = 0
               tempRate.PRO_DST_KG = 0
               tempRate.PRO_OTHER1_KG = 0
               tempRate.PRO_OTHER2_KG = 0
               tempRate.PRO_OTHER3_KG = 0
               tempRate.SUM_RATE_OTHER_BAG = RateOther
           ElseIf Pi.PART_TYPE = 21 Then
               RateOther = TempD2.PRO_COMMISSION_KG + TempD2.PRO_CHEER_KG + TempD2.PRO_DST_KG + TempD2.PRO_OTHER1_KG + TempD2.PRO_OTHER2_KG + TempD2.PRO_OTHER3_KG
            
               tempRate.PRO_COMMISSION_BAG = 0
               tempRate.PRO_CHEER_BAG = 0
               tempRate.PRO_DST_BAG = 0
               tempRate.PRO_OTHER1_BAG = 0
               tempRate.PRO_OTHER2_BAG = 0
               tempRate.PRO_OTHER3_BAG = 0
               tempRate.PRO_COMMISSION_KG = TempD2.PRO_COMMISSION_KG
               tempRate.PRO_CHEER_KG = TempD2.PRO_CHEER_KG
               tempRate.PRO_DST_KG = TempD2.PRO_DST_KG
               tempRate.PRO_OTHER1_KG = TempD2.PRO_OTHER1_KG
               tempRate.PRO_OTHER2_KG = TempD2.PRO_OTHER2_KG
               tempRate.PRO_OTHER3_KG = TempD2.PRO_OTHER3_KG
               tempRate.SUM_RATE_OTHER_KG = RateOther
            End If
            RateOtherFlag = True
         Else
            RateOtherFlag = False
         End If
         
         Call TempData.add(tempRate, Trim(str(CUSTOMER_ID)))
         
         
          If PRICE_THINK_TYPE = 1 Then
            If Not RateWotkPriceFlag Then
               glbErrorLog.LocalErrorMsg = "ยังไม่ได้กำหนดราคาสินค้า กรุณากำหนดราคาสินค้าก่อน"
               glbErrorLog.ShowUserError
               calExWorksPrice = -1
               Exit Function
            End If
         ElseIf PRICE_THINK_TYPE = 2 Then
             If Not RateWotkPriceFlag Then
               glbErrorLog.LocalErrorMsg = "ยังไม่ได้กำหนดราคาสินค้า กรุณากำหนดราคาสินค้าก่อน"
               glbErrorLog.ShowUserError
               calExWorksPrice = -1
               Exit Function
            End If
            
            If Not RateDeliveryCostFlag Then
               glbErrorLog.LocalErrorMsg = "ยังไม่ได้กำหนดราคาค่าขนส่ง กรุณากำหนดราคาค่าขนส่งก่อน"
               glbErrorLog.ShowUserError
               calExWorksPrice = -1
               Exit Function
            End If
         ElseIf PRICE_THINK_TYPE = 3 Then
             If Not RateWotkPriceFlag Then
               glbErrorLog.LocalErrorMsg = "ยังไม่ได้กำหนดราคาสินค้า กรุณากำหนดราคาสินค้าก่อน"
               glbErrorLog.ShowUserError
               calExWorksPrice = -1
               Exit Function
            End If
         End If
'         RateWotkPrice = RateWotkPrice - RatePromotionPart
         If PRICE_THINK_TYPE = 1 Then
            calExWorksPrice = RateWotkPrice + RateOther
         ElseIf PRICE_THINK_TYPE = 2 Then
            RateDeliveryCost = RateDeliveryCost - RatePromotionDlc 'ถ้ามีส่วนลดค่าขนส่งก็ให้ลบออกไปด้วย
            calExWorksPrice = RateWotkPrice + RateDeliveryCost + RateOther
         ElseIf PRICE_THINK_TYPE = 3 Then
            calExWorksPrice = RateWotkPrice + RateOther
         Else
           calExWorksPrice = "0"
         End If
      End If
End Function
Public Sub calExDeliveryCost(DELIVERY_CUS_ITEM_ID As Long, CUSTOMER_ID As Long, WEIGHT_PER_PACK As Long, PART_TYPE As Long, PRICE_THINK_TYPE As Long, m_ExDeliveryCostItem As Collection, m_Customers As Collection, ByRef RateDeliveryCost As Double, ByRef RateDeliveryCostFlag As Boolean, ByRef EX_DELIVERY_COST_ITEM_ID As Long)
Dim TempD As CExWorksPrice
Dim TempD2 As CCustomer

   Set TempD2 = GetObject("CCustomer", m_Customers, Trim(str(CUSTOMER_ID)), False)
   
      If Not TempD2 Is Nothing Then
      If TempD2.CAL_RATE_DELIVERY_TYPE = 1 Then 'คิดตามปริมาณที่คิดลูกค้า
        If PART_TYPE = 10 Then
             Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(WEIGHT_PER_PACK)), False)   'ค้นหาราคาค่าขนส่ง ตาม id สถานที่จัดส่ง และน้ำหนักจริงของอาหารเบอร์นั้นๆ
            If Not TempD Is Nothing Then
               RateDeliveryCost = TempD.RATE_CUSTOMER
               EX_DELIVERY_COST_ITEM_ID = TempD.EX_DELIVERY_COST_ITEM_ID
               RateDeliveryCostFlag = True
            Else
               Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(30)), False)  'ค้นหาราคาค่าขนส่ง  ที่ 30 กิโลเลย ถ้ามีก็ให้คำนวณบนฐาน 30 กก.
               If Not TempD Is Nothing Then
                  RateDeliveryCost = (TempD.RATE_CUSTOMER * WEIGHT_PER_PACK) / 30
                  EX_DELIVERY_COST_ITEM_ID = TempD.EX_DELIVERY_COST_ITEM_ID
                  RateDeliveryCostFlag = True
               Else
                  RateDeliveryCostFlag = False
               End If
            End If
         ElseIf PART_TYPE = 21 Then
            Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(1)), False)  'ค้นหาราคาค่าขนส่ง  ที่ 1 กิโลเลย
            If Not TempD Is Nothing Then
               RateDeliveryCost = TempD.RATE_CUSTOMER
               EX_DELIVERY_COST_ITEM_ID = TempD.EX_DELIVERY_COST_ITEM_ID
               RateDeliveryCostFlag = True
            Else
               RateDeliveryCostFlag = False
            End If
         End If
       ElseIf TempD2.CAL_RATE_DELIVERY_TYPE = 2 Then 'คิดตามเที่ยวที่คิดลูกค้า
         Set TempD = GetObject("CExWorksPrice", m_ExDeliveryCostItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(999)), False)  'ค้นหาราคาค่าขนส่ง  ที่ 999 กิโลเลย
            If Not TempD Is Nothing Then
               RateDeliveryCost = TempD.RATE_CUSTOMER
               EX_DELIVERY_COST_ITEM_ID = TempD.EX_DELIVERY_COST_ITEM_ID
               RateDeliveryCostFlag = True
            Else
               RateDeliveryCostFlag = False
            End If
         End If
      End If
         
      If Not RateDeliveryCostFlag And PRICE_THINK_TYPE > 1 Then
         glbErrorLog.LocalErrorMsg = "ยังไม่ได้กำหนดราคาค่าขนส่ง กรุณากำหนดราคาค่าขนส่งก่อน"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   
End Sub
Public Sub calExPromotionDls(DELIVERY_CUS_ITEM_ID As Long, CUSTOMER_ID As Long, WEIGHT_PER_PACK As Long, PART_TYPE As Long, PRICE_THINK_TYPE As Long, m_ExPromotionDlcItem As Collection, m_Customers As Collection, ByRef DISCOUNT_AMOUNT As Double, ByRef RatePromotionDlcFlag As Boolean, ByRef EX_PROMOTION_DLC_ITEM_ID As Long)
Dim TempD As CExWorksPrice
Dim TempD2 As CCustomer

   Set TempD2 = GetObject("CCustomer", m_Customers, Trim(str(CUSTOMER_ID)), False)
   
      If Not TempD2 Is Nothing Then
      If TempD2.CAL_RATE_DELIVERY_TYPE = 1 Then 'คิดตามปริมาณที่คิดลูกค้า
        If PART_TYPE = 10 Then
             Set TempD = GetObject("CExWorksPrice", m_ExPromotionDlcItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(WEIGHT_PER_PACK)), False)   'ค้นหาราคาค่าขนส่ง ตาม id สถานที่จัดส่ง และน้ำหนักจริงของอาหารเบอร์นั้นๆ
            If Not TempD Is Nothing Then
               DISCOUNT_AMOUNT = TempD.DISCOUNT_AMOUNT
               EX_PROMOTION_DLC_ITEM_ID = TempD.EX_PROMOTION_DLC_ITEM_ID
               RatePromotionDlcFlag = True
            Else
               Set TempD = GetObject("CExWorksPrice", m_ExPromotionDlcItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(30)), False)  'ค้นหาราคาค่าขนส่ง  ที่ 30 กิโลเลย ถ้ามีก็ให้คำนวณบนฐาน 30 กก.
               If Not TempD Is Nothing Then
                  DISCOUNT_AMOUNT = (TempD.DISCOUNT_AMOUNT * WEIGHT_PER_PACK) / 30
                  EX_PROMOTION_DLC_ITEM_ID = TempD.EX_PROMOTION_DLC_ITEM_ID
                  RatePromotionDlcFlag = True
               Else
                  RatePromotionDlcFlag = False
               End If
            End If
         ElseIf PART_TYPE = 21 Then
            Set TempD = GetObject("CExWorksPrice", m_ExPromotionDlcItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(1)), False)  'ค้นหาราคาค่าขนส่ง  ที่ 1 กิโลเลย
            If Not TempD Is Nothing Then
               DISCOUNT_AMOUNT = TempD.DISCOUNT_AMOUNT
               EX_PROMOTION_DLC_ITEM_ID = TempD.EX_PROMOTION_DLC_ITEM_ID
               RatePromotionDlcFlag = True
            Else
               RatePromotionDlcFlag = False
            End If
         End If
       ElseIf TempD2.CAL_RATE_DELIVERY_TYPE = 2 Then 'คิดตามเที่ยวที่คิดลูกค้า
         Set TempD = GetObject("CExWorksPrice", m_ExPromotionDlcItem, Trim(str(DELIVERY_CUS_ITEM_ID)) & "-" & Trim(str(999)), False)  'ค้นหาราคาค่าขนส่ง  ที่ 999 กิโลเลย
            If Not TempD Is Nothing Then
               DISCOUNT_AMOUNT = TempD.DISCOUNT_AMOUNT
               EX_PROMOTION_DLC_ITEM_ID = TempD.EX_PROMOTION_DLC_ITEM_ID
               RatePromotionDlcFlag = True
            Else
               RatePromotionDlcFlag = False
            End If
         End If
      End If
         
'      If Not RatePromotionDlcFlag And PRICE_THINK_TYPE <> 1 Then
'         glbErrorLog.LocalErrorMsg = "ยังไม่ได้กำหนดราคาค่าขนส่ง กรุณากำหนดราคาค่าขนส่งก่อน"
'         glbErrorLog.ShowUserError
'         Exit Sub
'      End If
End Sub
Public Function CheckProOtherName(Key1 As String, Key2 As String, Key3 As String) As String
Dim tStr As String
Dim isHasPrev  As Boolean
tStr = ""
 If Len(Key1) > 0 Then
   tStr = "อื่นๆ1=" & Key1
   isHasPrev = True
 End If
 
  If Len(Key2) > 0 Then
      If isHasPrev Then
         tStr = tStr & ",อื่นๆ2=" & Key2
      Else
         tStr = "อื่นๆ2=" & Key2
      End If
      isHasPrev = True
 End If
 
 If Len(Key3) > 0 Then
      If isHasPrev Then
         tStr = tStr & ",อื่นๆ3=" & Key3
      Else
         tStr = "อื่นๆ3=" & Key3
      End If
 End If
 CheckProOtherName = tStr
End Function
Public Function CheckNewMounth() As Boolean
Dim ServerDateTime As String
CheckNewMounth = False
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      If Mid$(ServerDateTime, 9, 2) <= 10 Then
         CheckNewMounth = True
      End If
End Function
Public Function isDiff(Value As Double) As Double
    isDiff = 0
   If Value > 0 Then
      isDiff = MyDiff(Value * 100, 102)
   End If
End Function
Public Function getKeyCbo(data As String) As String
Dim strKey() As String
strKey = Split(data, "k:")
If UBound(strKey) > 0 Then
   getKeyCbo = strKey(1)
Else
   getKeyCbo = ""
End If
End Function
