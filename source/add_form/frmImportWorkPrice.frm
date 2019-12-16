VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportWorkPrice 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13560
   Icon            =   "frmImportWorkPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8805
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   15531
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3150
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   3600
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   6
         Top             =   3930
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9780
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1005
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlFromActiveDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   16
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
      End
      Begin prjFarmManagement.uctlDate uctlToValidDate 
         Height          =   495
         Left            =   1860
         TabIndex        =   17
         Top             =   2100
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
      End
      Begin VB.Label lblFromActiveDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblToValidDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   2955
         Left            =   480
         TabIndex        =   15
         Top             =   5160
         Width           =   12585
      End
      Begin VB.Label lblInventoryActDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblInventoryActDate"
         Height          =   315
         Left            =   480
         TabIndex        =   14
         Top             =   1095
         Width           =   1305
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   12480
         TabIndex        =   1
         Top             =   3150
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportWorkPrice.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   4470
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportWorkPrice.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   12
         Top             =   4050
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   3660
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFileName"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   3180
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10935
         TabIndex        =   4
         Top             =   4470
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   9285
         TabIndex        =   3
         Top             =   4470
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportWorkPrice.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportWorkPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public Area As Long

Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private PartUctlColls As Collection
Private PartColls As Collection
Private m_Customers As Collection
Private m_DeliveryCus As Collection
Private temp_DeliveryCus As Collection
Private PartLabColls  As Collection
Private PartLabUpdateColls  As Collection

Private m_PartItems As Collection

Private isSave As Boolean

Private Sub cmdFileName_Click()
 On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub



Private Sub cmdOK_Click()
   OKClick = True
   If isSave Then
      glbDatabaseMngr.DBConnection.CommitTrans
   End If
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   If Area = 1 Or Area = 2 Or Area = 3 Or Area = 4 Then
      If Not VerifyDate(lblInventoryActDate, uctlDocumentDate, False) Then
         Exit Sub
      End If
   End If
   
   Call EnableForm(Me, False)

   If Area = 1 Then
     Call LoadPartItem(Nothing, PartColls, , , , 2)
      Call ImportExWorksPrice
   ElseIf Area = 2 Then
      Call LoadCustomer(Nothing, m_Customers, 2)
      Call LoadDeliveryCus(Nothing, m_DeliveryCus, , , , 2)
      Call ImportRateDelivery
   ElseIf Area = 3 Then
     Call LoadCustomer(Nothing, m_Customers, 2)
     Call LoadPartItem(Nothing, PartColls, , , , 2)
     Call ImportPromotionExWorksPrice
   ElseIf Area = 4 Then
      Call LoadCustomer(Nothing, m_Customers, 2)
      Call LoadDeliveryCus(Nothing, m_DeliveryCus, , , , 2)
      Call ImportPromotionDelivery
   End If
   Call EnableForm(Me, True)
End Sub
'ImportPromotionExWorksPrice
Private Sub ImportPromotionDelivery()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim I As Long
Dim J As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim EWP As CExWorksPrice
Dim EPDI As CExPromotionDlcItem
Dim DC As CDeliveryCus
Dim DC2 As CDeliveryCus
Dim IsOK As Boolean
Dim SearchCusID As CCustomer
Dim SearchCusDly As CDeliveryCus
Dim SearchCusDly2 As CDeliveryCus
Dim DlyCusItem As CExPromotionDlcItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim Beginrow As Long
Dim CusID As Long
Dim TempRateType As Long
Dim IsFind As Boolean

 Dim tempDCI_Code As String
Dim R As Long


   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   StatusInsert = True
   HasBegin = False

   ID = 1
   Beginrow = 5
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 0
   prgProgress.MAX = (MaxRow * 3) + 1
   
   uctlDocumentDate.ShowDate = m_ExcelSheet.Cells(1, 2).Value
   uctlFromActiveDate.ShowDate = m_ExcelSheet.Cells(2, 2).Value
   uctlToValidDate.ShowDate = m_ExcelSheet.Cells(3, 2).Value
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlFromActiveDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีมีผล ") & " " & uctlFromActiveDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdStart.Enabled = True
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlToValidDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีสิ้นสุด ") & " " & uctlToValidDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdStart.Enabled = True
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   For row = Beginrow To MaxRow - 1
      DoEvents
      Me.Refresh
      
   If Len(m_ExcelSheet.Cells(row, 1).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 4).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 5).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 6).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 7).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 8).Value) = 0 Then
      If Trim(m_ExcelSheet.Cells(row, 8).Value) = 1 Then  'ใช้งาน
         MsgBox "เอกสาร : <" & txtFileName.Text & ">  บรรทัดที่ " & row & " ต้องไม่เป็นช่องว่าง กรุณาตรวจสอบ"
         Call EnableForm(Me, True)
         cmdStart.Enabled = True
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Exit Sub
      End If
   End If
   
'      If Trim(m_ExcelSheet.Cells(row, 1).Value) <> "***" Then
'         For J = row + 1 To MaxRow 'ตรวจสอบบรรทัดที่เบอร์ซ้ำกัน
''            If Trim(m_ExcelSheet.Cells(row, 3).Value) = Trim(m_ExcelSheet.Cells(J, 3).Value) And Len(m_ExcelSheet.Cells(row, 3).Value) > 0 And Trim(m_ExcelSheet.Cells(J, 1).Value) <> "***" Then
'             If Trim(m_ExcelSheet.Cells(row, 3).Value) = Trim(m_ExcelSheet.Cells(J, 3).Value) And Len(m_ExcelSheet.Cells(row, 3).Value) > 0 Then
'                 MsgBox "เอกสาร : <" & txtFileName.Text & ">  รหัสสถานที่ : " & m_ExcelSheet.Cells(row, 3).Value & " บรรทัดที่ " & row & " ซ้ำกับบรรทัดที่ " & J & " กรุณาแก้ไขไม่ให้ซ้ำกัน"
'                  Call EnableForm(Me, True)
'                  cmdStart.Enabled = True
'                  cmdExit.Enabled = True
'                  cmdOK.Enabled = True
'                  Exit Sub
'            End If
'         Next J
'   Else
'      Exit For
'   End If
   
  
   
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   prgProgress.MAX = 100
   ProgressCount = 0
   
   For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh
      If Trim(m_ExcelSheet.Cells(row, 8).Value) = 1 Then 'ใช้งาน
         If Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
            Set SearchCusID = GetObject("CCustomer", m_Customers, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
            If Not SearchCusID Is Nothing Then
               CusID = SearchCusID.CUSTOMER_ID
            Else
               CusID = 0
            End If
            
            If CusID > 0 Then
             Set DC2 = New CDeliveryCus
                  R = 1
                  tempDCI_Code = ""
                  
                  While R <> -1
                     If Len(Trim(m_ExcelSheet.Cells(row, 3).Value)) > 0 Then
                     tempDCI_Code = Trim(m_ExcelSheet.Cells(row, 3).Value)
                   Else
                     tempDCI_Code = Trim(m_ExcelSheet.Cells(row, 1).Value) & "-" & Format(R, "00")
                     DC2.Flag = "N"
                   End If
                   
                    Set SearchCusDly = GetObject("CDeliveryCus", m_DeliveryCus, Trim(tempDCI_Code) & "-" & Trim(str(CusID)), False)
                        If Not SearchCusDly Is Nothing Then
                          If SearchCusDly.Flag = "I" And DC2.Flag <> "N" Then  'กรณีเจอ แล้วข้อมูลอยู่ใน database
                               IsFind = True
                               DC2.KEY_ID = row
                               DC2.DELIVERY_CUS_ITEM_CODE = tempDCI_Code
                               Call temp_DeliveryCus.add(DC2, Trim(str(row)))
                               R = -1
                          Else
                              IsFind = True
                              R = R + 1
                              tempDCI_Code = Trim(m_ExcelSheet.Cells(row, 1).Value) & "-" & Format(R, "00")
                           End If
                        Else
                           IsFind = False
                           
                           DC2.KEY_ID = row
                           DC2.DELIVERY_CUS_ITEM_CODE = tempDCI_Code
                           Call temp_DeliveryCus.add(DC2, Trim(str(row)))
                           
                           R = -1
                        End If
                 Wend
          
               If Not IsFind Then
                  Set DC = New CDeliveryCus
         
                  DC.DELIVERY_CUS_ITEM_CODE = tempDCI_Code
                  DC.DELIVERY_CUS_ITEM_NAME = Trim(m_ExcelSheet.Cells(row, 4).Value)
                  DC.CUSTOMER_ID = CusID
               
                  'ใส่เสมอ
                  DC.AddEditMode = SHOW_ADD
                  DC.Flag = "A"
                  
                  
                  If Not glbDaily.AddEditDeliveryCus(DC, IsOK, False, glbErrorLog) Then
                      Call EnableForm(Me, True)
                      HasBegin = False
                   End If
                    If Not IsOK Then
                         Call EnableForm(Me, True)
                         glbErrorLog.ShowUserError
                         glbDatabaseMngr.DBConnection.RollbackTrans
                     Else
                         Call m_DeliveryCus.add(DC, Trim(DC.DELIVERY_CUS_ITEM_CODE) & "-" & Trim(str(CusID)))
                        isSave = True
                     End If
                  End If
            End If
            
         ProgressCount = ProgressCount + 1
         prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
         txtPercent.Text = prgProgress.Value
         Else
               row = row + 2
         End If
      End If
   Next row

   Set DC = Nothing
   ProgressCount = 0
   Set EWP = New CExWorksPrice
   EWP.AddEditMode = SHOW_ADD
'   If Area = 1 Or Area = 2 Then
'      EWP.EX_WORKS_PRICE_DATE = uctlDocumentDate.ShowDate
'   Else
'      EWP.EX_WORKS_PRICE_DATE = Now
'   End If
   EWP.EX_WORKS_PRICE_DATE = uctlDocumentDate.ShowDate
   EWP.EX_WORKS_PRICE_TYPE = Area
   EWP.EX_WORKS_PRICE_LEVEL = "Y"
   EWP.EX_WORKS_PRICE_CODE = AutoGenName
   EWP.EX_WORKS_PRICE_DESC = "IMPORTED " & DateToStringExtEx3(Now)
   EWP.EX_WORKS_PRICE_STATUS = 0

   EWP.FROM_ACTIVE_DATE = uctlFromActiveDate.ShowDate
   EWP.TO_VALID_DATE = uctlToValidDate.ShowDate
   
   Call LoadCustomer(Nothing, m_Customers, 2)
   Call LoadDeliveryCus(Nothing, m_DeliveryCus, , , , 2)
   
     
     For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh
      
      If Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set SearchCusID = GetObject("CCustomer", m_Customers, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
         If Not SearchCusID Is Nothing Then
            CusID = SearchCusID.CUSTOMER_ID
         Else
            CusID = 0
         End If
         
         If CusID > 0 Then
            Set SearchCusDly2 = GetObject("CDeliveryCus", temp_DeliveryCus, Trim(str(row)), False)
            If Not SearchCusDly2 Is Nothing Then
               Set SearchCusDly = GetObject("CDeliveryCus", m_DeliveryCus, Trim(SearchCusDly2.DELIVERY_CUS_ITEM_CODE) & "-" & Trim(str(CusID)), False)
               If Not SearchCusDly Is Nothing Then
                  IsFind = True
               Else
                  IsFind = False
               End If
            Else
               IsFind = False
            End If
                    
            If IsFind Then
               Set DlyCusItem = New CExPromotionDlcItem
               DlyCusItem.DISCOUNT_AMOUNT = Val(m_ExcelSheet.Cells(row, 5).Value)
               
               If Val(m_ExcelSheet.Cells(row, 6).Value) = 1 Then 'เป็นถุง
                  DlyCusItem.RATE_TYPE_CUS = 1
                  DlyCusItem.WEIGHT_PER_PACK_CUS = Val(m_ExcelSheet.Cells(row, 7).Value)
               ElseIf Val(m_ExcelSheet.Cells(row, 6).Value) = 2 Then 'เป็นกิโลกรัม
                  DlyCusItem.RATE_TYPE_CUS = 2
                  DlyCusItem.WEIGHT_PER_PACK_CUS = "1"
               ElseIf Val(m_ExcelSheet.Cells(row, 6).Value) = 3 Then 'เป็นเที่ยว
                  DlyCusItem.RATE_TYPE_CUS = 3
                  DlyCusItem.WEIGHT_PER_PACK_CUS = "999"
               End If
               
               DlyCusItem.CUSTOMER_ID = SearchCusDly.CUSTOMER_ID
               DlyCusItem.DELIVERY_CUS_ITEM_ID = SearchCusDly.DELIVERY_CUS_ITEM_ID
               'ใส่เสมอ
               DlyCusItem.AddEditMode = SHOW_ADD
               DlyCusItem.Flag = "A"
               
               Call EWP.ExPromotionDlc.add(DlyCusItem)
            End If
         End If
         
      ProgressCount = ProgressCount + 1
      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
      txtPercent.Text = prgProgress.Value
      End If
   Next row
   
   If Not glbDaily.AddEditExWorksPrice(EWP, IsOK, False, glbErrorLog) Then
      Call EnableForm(Me, True)
      HasBegin = False
   End If
      If Not IsOK Then
         Call EnableForm(Me, True)
         glbErrorLog.ShowUserError
         glbDatabaseMngr.DBConnection.RollbackTrans
     Else
        isSave = True
     End If
  
   Set DlyCusItem = Nothing
   

   prgProgress.Value = prgProgress.MAX
   Set m_ExcelSheet = Nothing
   
   'cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   m_ExcelApp.Workbooks.Close
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub ImportRateDelivery()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim I As Long
Dim J As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim EWP As CExWorksPrice
Dim EWPI As CExWorksPriceItem
Dim DC As CDeliveryCus
Dim DC2 As CDeliveryCus
Dim IsOK As Boolean
Dim SearchCusID As CCustomer
Dim SearchCusDly As CDeliveryCus
Dim SearchCusDly2 As CDeliveryCus
Dim DlyCusItem As CExDeliveryCostItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim Beginrow As Long
Dim CusID As Long
Dim TempRateType As Long
Dim IsFind As Boolean

 Dim tempDCI_Code As String
Dim R As Long


   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   StatusInsert = True
   HasBegin = False

   ID = 1
   Beginrow = 5
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 0
   prgProgress.MAX = (MaxRow * 3) + 1
   
   uctlDocumentDate.ShowDate = m_ExcelSheet.Cells(1, 2).Value
   uctlFromActiveDate.ShowDate = m_ExcelSheet.Cells(2, 2).Value
   uctlToValidDate.ShowDate = m_ExcelSheet.Cells(3, 2).Value
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlFromActiveDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีมีผล ") & " " & uctlFromActiveDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdStart.Enabled = True
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlToValidDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีสิ้นสุด ") & " " & uctlToValidDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdStart.Enabled = True
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   For row = Beginrow To MaxRow - 1
      DoEvents
      Me.Refresh
      
   If Len(m_ExcelSheet.Cells(row, 1).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 4).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 5).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 6).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 7).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 8).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 9).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 10).Value) = 0 Then 'Or Len(m_ExcelSheet.Cells(row, 3).Value) = 0
      If Trim(m_ExcelSheet.Cells(row, 11).Value) = 1 Then  'ใช้งาน
         MsgBox "เอกสาร : <" & txtFileName.Text & ">  บรรทัดที่ " & row & " ต้องไม่เป็นช่องว่าง กรุณาตรวจสอบ"
         Call EnableForm(Me, True)
         cmdStart.Enabled = True
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Exit Sub
      End If
   End If
   
'      If Trim(m_ExcelSheet.Cells(row, 1).Value) <> "***" Then
'         For J = row + 1 To MaxRow 'ตรวจสอบบรรทัดที่เบอร์ซ้ำกัน
''            If Trim(m_ExcelSheet.Cells(row, 3).Value) = Trim(m_ExcelSheet.Cells(J, 3).Value) And Len(m_ExcelSheet.Cells(row, 3).Value) > 0 And Trim(m_ExcelSheet.Cells(J, 1).Value) <> "***" Then
'             If Trim(m_ExcelSheet.Cells(row, 3).Value) = Trim(m_ExcelSheet.Cells(J, 3).Value) And Len(m_ExcelSheet.Cells(row, 3).Value) > 0 Then
'                 MsgBox "เอกสาร : <" & txtFileName.Text & ">  รหัสสถานที่ : " & m_ExcelSheet.Cells(row, 3).Value & " บรรทัดที่ " & row & " ซ้ำกับบรรทัดที่ " & J & " กรุณาแก้ไขไม่ให้ซ้ำกัน"
'                  Call EnableForm(Me, True)
'                  cmdStart.Enabled = True
'                  cmdExit.Enabled = True
'                  cmdOK.Enabled = True
'                  Exit Sub
'            End If
'         Next J
'   Else
'      Exit For
'   End If
   
  
   
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   prgProgress.MAX = 100
   ProgressCount = 0
   
   For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh
      If Trim(m_ExcelSheet.Cells(row, 11).Value) = 1 Then 'ใช้งาน
         If Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
            Set SearchCusID = GetObject("CCustomer", m_Customers, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
            If Not SearchCusID Is Nothing Then
               CusID = SearchCusID.CUSTOMER_ID
            Else
               CusID = 0
            End If
            
            If CusID > 0 Then
             Set DC2 = New CDeliveryCus
                  R = 1
                  tempDCI_Code = ""
                  
                  While R <> -1
                     If Len(Trim(m_ExcelSheet.Cells(row, 3).Value)) > 0 Then
                     tempDCI_Code = Trim(m_ExcelSheet.Cells(row, 3).Value)
                   Else
                     tempDCI_Code = Trim(m_ExcelSheet.Cells(row, 1).Value) & "-" & Format(R, "00")
                     DC2.Flag = "N"
                   End If
                   
                    Set SearchCusDly = GetObject("CDeliveryCus", m_DeliveryCus, Trim(tempDCI_Code) & "-" & Trim(str(CusID)), False)
                        If Not SearchCusDly Is Nothing Then
                          If SearchCusDly.Flag = "I" And DC2.Flag <> "N" Then  'กรณีเจอ แล้วข้อมูลอยู่ใน database
                               IsFind = True
                               DC2.KEY_ID = row
                               DC2.DELIVERY_CUS_ITEM_CODE = tempDCI_Code
                               Call temp_DeliveryCus.add(DC2, Trim(str(row)))
                               R = -1
                          Else
                              IsFind = True
                              R = R + 1
                              tempDCI_Code = Trim(m_ExcelSheet.Cells(row, 1).Value) & "-" & Format(R, "00")
                           End If
                        Else
                           IsFind = False
                           
                           DC2.KEY_ID = row
                           DC2.DELIVERY_CUS_ITEM_CODE = tempDCI_Code
                           Call temp_DeliveryCus.add(DC2, Trim(str(row)))
                           
                           R = -1
                        End If
                 Wend
          
               If Not IsFind Then
                  Set DC = New CDeliveryCus
         
                  DC.DELIVERY_CUS_ITEM_CODE = tempDCI_Code
                  DC.DELIVERY_CUS_ITEM_NAME = Trim(m_ExcelSheet.Cells(row, 4).Value)
                  DC.CUSTOMER_ID = CusID
               
                  'ใส่เสมอ
                  DC.AddEditMode = SHOW_ADD
                  DC.Flag = "A"
                  
                  
                  If Not glbDaily.AddEditDeliveryCus(DC, IsOK, False, glbErrorLog) Then
                      Call EnableForm(Me, True)
                      HasBegin = False
                   End If
                    If Not IsOK Then
                         Call EnableForm(Me, True)
                         glbErrorLog.ShowUserError
                         glbDatabaseMngr.DBConnection.RollbackTrans
                     Else
                         Call m_DeliveryCus.add(DC, Trim(DC.DELIVERY_CUS_ITEM_CODE) & "-" & Trim(str(CusID)))
                        isSave = True
                     End If
                  End If
            End If
            
         ProgressCount = ProgressCount + 1
         prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
         txtPercent.Text = prgProgress.Value
         Else
               row = row + 2
         End If
      End If
   Next row

   Set DC = Nothing
   ProgressCount = 0
   Set EWP = New CExWorksPrice
   EWP.AddEditMode = SHOW_ADD
   If Area = 1 Or Area = 2 Then
      EWP.EX_WORKS_PRICE_DATE = uctlDocumentDate.ShowDate
   Else
      EWP.EX_WORKS_PRICE_DATE = Now
   End If
   EWP.EX_WORKS_PRICE_TYPE = Area
   EWP.EX_WORKS_PRICE_LEVEL = "Y"
   EWP.EX_WORKS_PRICE_CODE = AutoGenName
   EWP.EX_WORKS_PRICE_DESC = "IMPORTED " & DateToStringExtEx3(Now)
   EWP.EX_WORKS_PRICE_STATUS = 0

   EWP.FROM_ACTIVE_DATE = uctlFromActiveDate.ShowDate
   EWP.TO_VALID_DATE = uctlToValidDate.ShowDate
   
   Call LoadCustomer(Nothing, m_Customers, 2)
   Call LoadDeliveryCus(Nothing, m_DeliveryCus, , , , 2)
   
     
     For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh
      
      If Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set SearchCusID = GetObject("CCustomer", m_Customers, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
         If Not SearchCusID Is Nothing Then
            CusID = SearchCusID.CUSTOMER_ID
         Else
            CusID = 0
         End If
         
         If CusID > 0 Then
            Set SearchCusDly2 = GetObject("CDeliveryCus", temp_DeliveryCus, Trim(str(row)), False)
            If Not SearchCusDly2 Is Nothing Then
               Set SearchCusDly = GetObject("CDeliveryCus", m_DeliveryCus, Trim(SearchCusDly2.DELIVERY_CUS_ITEM_CODE) & "-" & Trim(str(CusID)), False)
               If Not SearchCusDly Is Nothing Then
                  IsFind = True
               Else
                  IsFind = False
               End If
            Else
               IsFind = False
            End If
                    
            If IsFind Then
               Set DlyCusItem = New CExDeliveryCostItem
               DlyCusItem.RATE_DELIVERY = Val(m_ExcelSheet.Cells(row, 5).Value)
               DlyCusItem.RATE_CUSTOMER = Val(m_ExcelSheet.Cells(row, 8).Value)
               If Val(m_ExcelSheet.Cells(row, 6).Value) = 1 Then 'เป็นถุง
                  DlyCusItem.RATE_TYPE = 1
                  DlyCusItem.WEIGHT_PER_PACK = Val(m_ExcelSheet.Cells(row, 7).Value)
               ElseIf Val(m_ExcelSheet.Cells(row, 6).Value) = 2 Then 'เป็นกิโลกรัม
                  DlyCusItem.RATE_TYPE = 2
                  DlyCusItem.WEIGHT_PER_PACK = "1"
               ElseIf Val(m_ExcelSheet.Cells(row, 6).Value) = 3 Then 'เป็นเที่ยว
                  DlyCusItem.RATE_TYPE = 3
                  DlyCusItem.WEIGHT_PER_PACK = "999"
               End If
               
               If Val(m_ExcelSheet.Cells(row, 9).Value) = 1 Then 'เป็นถุง
                  DlyCusItem.RATE_TYPE_CUS = 1
                  DlyCusItem.WEIGHT_PER_PACK_CUS = Val(m_ExcelSheet.Cells(row, 10).Value)
               ElseIf Val(m_ExcelSheet.Cells(row, 9).Value) = 2 Then 'เป็นกิโลกรัม
                  DlyCusItem.RATE_TYPE_CUS = 2
                  DlyCusItem.WEIGHT_PER_PACK_CUS = "1"
               ElseIf Val(m_ExcelSheet.Cells(row, 9).Value) = 3 Then 'เป็นเที่ยว
                  DlyCusItem.RATE_TYPE_CUS = 3
                  DlyCusItem.WEIGHT_PER_PACK_CUS = "999"
               End If
               
               DlyCusItem.CUSTOMER_ID = SearchCusDly.CUSTOMER_ID
               DlyCusItem.DELIVERY_CUS_ITEM_ID = SearchCusDly.DELIVERY_CUS_ITEM_ID
               'ใส่เสมอ
               DlyCusItem.AddEditMode = SHOW_ADD
               DlyCusItem.Flag = "A"
               
               Call EWP.ExDeliveryCost.add(DlyCusItem)
            End If
         End If
         
      ProgressCount = ProgressCount + 1
      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
      txtPercent.Text = prgProgress.Value
      End If
   Next row
   
   If Not glbDaily.AddEditExWorksPrice(EWP, IsOK, False, glbErrorLog) Then
      Call EnableForm(Me, True)
      HasBegin = False
   End If
      If Not IsOK Then
         Call EnableForm(Me, True)
         glbErrorLog.ShowUserError
         glbDatabaseMngr.DBConnection.RollbackTrans
     Else
        isSave = True
     End If
  
   Set DlyCusItem = Nothing
   

   prgProgress.Value = prgProgress.MAX
   Set m_ExcelSheet = Nothing
   
   'cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   m_ExcelApp.Workbooks.Close
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub ImportExWorksPrice()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim I As Long
Dim J As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim EWP As CExWorksPrice
Dim EWPI As CExWorksPriceItem
Dim IsOK As Boolean
Dim SearchItemNo As CPartItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim SearchItemNo2 As CExWorksPriceItem
Dim Beginrow As Long

   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   StatusInsert = True
   HasBegin = False

   ID = 1
   Beginrow = 5
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 0
   prgProgress.MAX = (MaxRow * 3) + 1

   
   uctlDocumentDate.ShowDate = m_ExcelSheet.Cells(1, 2).Value
   uctlFromActiveDate.ShowDate = m_ExcelSheet.Cells(2, 2).Value
   uctlToValidDate.ShowDate = m_ExcelSheet.Cells(3, 2).Value
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlFromActiveDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีมีผล ") & " " & uctlFromActiveDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlToValidDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีสิ้นสุด ") & " " & uctlToValidDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   For row = Beginrow To MaxRow - 1
      DoEvents
      Me.Refresh
      If Len(m_ExcelSheet.Cells(row, 1).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 5).Value) = 0 Or Val(m_ExcelSheet.Cells(row, 5).Value) < 0 Then
         MsgBox "เอกสาร : <" & txtFileName.Text & ">  บรรทัดที่ " & row & " ต้องไม่เป็นช่องว่างและต้องมีค่ามากกว่า 0 กรุณาตรวจสอบ"
         Call EnableForm(Me, True)
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Exit Sub
      End If
   
      For J = row + 1 To MaxRow 'ตรวจสอบบรรทัดที่เบอร์ซ้ำกัน
         If Trim(m_ExcelSheet.Cells(row, 1).Value) = Trim(m_ExcelSheet.Cells(J, 1).Value) And Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
              MsgBox "เอกสาร : <" & txtFileName.Text & ">  เบอร์วัตถุดิบ : " & m_ExcelSheet.Cells(row, 1).Value & " บรรทัดที่ " & row & " ซ้ำกับบรรทัดที่ " & J & " กรุณาแก้ไขไม่ให้ซ้ำกัน"
               Call EnableForm(Me, True)
               cmdExit.Enabled = True
               cmdOK.Enabled = True
               Exit Sub
         End If
      Next J
   Next row

   Set EWP = New CExWorksPrice
   EWP.AddEditMode = SHOW_ADD
   
   EWP.EX_WORKS_PRICE_DATE = uctlDocumentDate.ShowDate
   EWP.FROM_ACTIVE_DATE = uctlFromActiveDate.ShowDate
   EWP.TO_VALID_DATE = uctlToValidDate.ShowDate
   EWP.EX_WORKS_PRICE_TYPE = Area
   EWP.EX_WORKS_PRICE_LEVEL = "Y"
   EWP.EX_WORKS_PRICE_CODE = AutoGenName
   EWP.EX_WORKS_PRICE_DESC = "IMPORTED " & DateToStringExtEx3(Now)
   EWP.EX_WORKS_PRICE_STATUS = 0

   prgProgress.MAX = 100
   For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh

      If Len(m_ExcelSheet.Cells(row, 1).Value) > 0 Then
         Set EWPI = New CExWorksPriceItem
         EWPI.Flag = "A"
         
       EWPI.PACKAGE_RATE = Val(m_ExcelSheet.Cells(row, 5).Value)
       
       Set SearchItemNo = GetObject("CPartItem", PartColls, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
         If Not SearchItemNo Is Nothing Then
            EWPI.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         Else
            EWPI.Flag = "I"
         End If
         EWPI.RATE_TYPE = 1
         
        Call EWP.ExWorksPriceItem.add(EWPI)
         Set EWPI = Nothing
      End If
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
      txtPercent.Text = prgProgress.Value
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If StatusInsert = True Then
      If Not glbDaily.AddEditExWorksPrice(EWP, IsOK, False, glbErrorLog) Then
          Call EnableForm(Me, True)
          HasBegin = False
       End If
       If Not IsOK Then
          Call EnableForm(Me, True)
          glbErrorLog.ShowUserError
      Else
         isSave = True
       End If
   Else
      lblNote.Caption = CodeForInsert
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      glbDatabaseMngr.DBConnection.RollbackTrans
      Exit Sub
   End If
   
   Set EWP = Nothing
   prgProgress.Value = prgProgress.MAX
   
   Set m_ExcelSheet = Nothing

   cmdExit.Enabled = True
   cmdOK.Enabled = True
   m_ExcelApp.Workbooks.Close
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub ImportPromotionExWorksPrice()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim I As Long
Dim J As Long
Dim row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim HasBegin As Boolean
Dim EWP As CExWorksPrice
Dim EPPI As CExPromotionPartItem
Dim IsOK As Boolean
Dim SearchCusID As CCustomer
Dim SearchItemNo As CPartItem
Dim StatusInsert As Boolean
Dim CodeForInsert As String
Dim SearchItemNo2 As CExWorksPriceItem
Dim Beginrow As Long

   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   StatusInsert = True
   HasBegin = False

   ID = 1
   Beginrow = 5
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
   
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 0
   prgProgress.MAX = (MaxRow * 3) + 1


   uctlDocumentDate.ShowDate = InternalDateToDate(NVLS(m_ExcelSheet.Cells(1, 2).Value, ""))
   uctlFromActiveDate.ShowDate = InternalDateToDate(NVLS(m_ExcelSheet.Cells(2, 2).Value, ""))
   uctlToValidDate.ShowDate = InternalDateToDate(NVLS(m_ExcelSheet.Cells(3, 2).Value, ""))
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlFromActiveDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีมีผล ") & " " & uctlFromActiveDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   If Not CheckUniqueNs(WORKS_PRICE_ACTIVE_DATE_UNIQUE, Trim(DateToStringInt(uctlToValidDate.ShowDate)), ID, , 4, , Area) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล ช่วงวันทีสิ้นสุด ") & " " & uctlToValidDate.ShowDate & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      Exit Sub
   End If
   
   For row = Beginrow To MaxRow - 1
      DoEvents
      Me.Refresh
      If Len(m_ExcelSheet.Cells(row, 1).Value) = 0 Or Len(m_ExcelSheet.Cells(row, 7).Value) = 0 Or Val(m_ExcelSheet.Cells(row, 7).Value) < 0 Then
         MsgBox "เอกสาร : <" & txtFileName.Text & ">  บรรทัดที่ " & row & " ต้องไม่เป็นช่องว่างและต้องมีค่ามากกว่า 0 กรุณาตรวจสอบ"
         Call EnableForm(Me, True)
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Exit Sub
      End If
   
      For J = row + 1 To MaxRow 'ตรวจสอบบรรทัดที่เบอร์ซ้ำกัน
         If Trim(m_ExcelSheet.Cells(row, 1).Value) & "-" & Trim(m_ExcelSheet.Cells(row, 3).Value) = Trim(m_ExcelSheet.Cells(J, 1).Value) & "-" & Trim(m_ExcelSheet.Cells(J, 3).Value) And Len(Trim(m_ExcelSheet.Cells(row, 1).Value) & Trim(m_ExcelSheet.Cells(row, 3).Value)) > 0 Then
              MsgBox "เอกสาร : <" & txtFileName.Text & ">  เบอร์วัตถุดิบ : " & m_ExcelSheet.Cells(row, 1).Value & " บรรทัดที่ " & row & " ซ้ำกับบรรทัดที่ " & J & " กรุณาแก้ไขไม่ให้ซ้ำกัน"
               Call EnableForm(Me, True)
               cmdExit.Enabled = True
               cmdOK.Enabled = True
               Exit Sub
         End If
      Next J
   Next row

   Set EWP = New CExWorksPrice
   EWP.AddEditMode = SHOW_ADD
   
   EWP.EX_WORKS_PRICE_DATE = uctlDocumentDate.ShowDate
   EWP.FROM_ACTIVE_DATE = uctlFromActiveDate.ShowDate
   EWP.TO_VALID_DATE = uctlToValidDate.ShowDate
   EWP.EX_WORKS_PRICE_TYPE = Area
   EWP.EX_WORKS_PRICE_LEVEL = "Y"
   EWP.EX_WORKS_PRICE_CODE = AutoGenName
   EWP.EX_WORKS_PRICE_DESC = "IMPORTED " & DateToStringExtEx3(Now)
   EWP.EX_WORKS_PRICE_STATUS = 0

   prgProgress.MAX = 100
   For row = Beginrow To MaxRow
      DoEvents
      Me.Refresh

      If Len(Trim(m_ExcelSheet.Cells(row, 1).Value) & Trim(m_ExcelSheet.Cells(row, 3).Value)) > 0 Then
         Set EPPI = New CExPromotionPartItem
         EPPI.Flag = "A"
         
       EPPI.DISCOUNT_AMOUNT = Val(m_ExcelSheet.Cells(row, 7).Value)
       
       Set SearchItemNo = GetObject("CPartItem", PartColls, Trim(m_ExcelSheet.Cells(row, 3).Value), False)
         If Not SearchItemNo Is Nothing Then
            EPPI.PART_ITEM_ID = SearchItemNo.PART_ITEM_ID
         Else
            EPPI.Flag = "I"
         End If
         EPPI.RATE_TYPE = 1
         
          Set SearchCusID = GetObject("CCustomer", m_Customers, Trim(m_ExcelSheet.Cells(row, 1).Value), False)
         If Not SearchCusID Is Nothing Then
            EPPI.CUSTOMER_ID = SearchCusID.CUSTOMER_ID
         Else
            EPPI.CUSTOMER_ID = -1
         End If
         
        Call EWP.ExPromotionPart.add(EPPI)
         Set EPPI = Nothing
      End If
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = MyDiff(ProgressCount, MaxRow) * 100
      txtPercent.Text = prgProgress.Value
   Next row
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If StatusInsert = True Then
      If Not glbDaily.AddEditExWorksPrice(EWP, IsOK, False, glbErrorLog) Then
          Call EnableForm(Me, True)
          HasBegin = False
       End If
       If Not IsOK Then
          Call EnableForm(Me, True)
          glbErrorLog.ShowUserError
      Else
         isSave = True
       End If
   Else
      lblNote.Caption = CodeForInsert
      Call EnableForm(Me, True)
      cmdExit.Enabled = True
      cmdOK.Enabled = True
      glbDatabaseMngr.DBConnection.RollbackTrans
      Exit Sub
   End If
   
   Set EWP = Nothing
   prgProgress.Value = prgProgress.MAX
   
   Set m_ExcelSheet = Nothing

   cmdExit.Enabled = True
   cmdOK.Enabled = True
   m_ExcelApp.Workbooks.Close
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If Area = 1 Or Area = 2 Or Area = 3 Or Area = 4 Then
         uctlDocumentDate.SetFocus
         uctlDocumentDate.ShowDate = Now
         uctlFromActiveDate.ShowDate = Now
         uctlToValidDate.ShowDate = Now
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "อิมพอร์ต" & HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblInventoryActDate, MapText("วันที่ประกาศ"))
   Call InitNormalLabel(lblFromActiveDate, MapText("วันที่มีผล"))
   Call InitNormalLabel(lblToValidDate, MapText("วันที่สิ้นสุด"))
   Dim str As String

   If Area = 2 Then
      Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
       str = "- เริ่ม Import ที่ Row ที่ 5 ,Column A ของ Sheet1 เท่านั้น โดย" & vbCrLf & "Col A2 = วันที่มีผล, Col A แถวที่ 5 = รหัสลูกค้า, Col B2 = วันที่สิ้นสุด, Col B แถวที่ 5 = ชื่อลูกค้า, Col C แถวที่ 5 = รหัสทีจัดส่ง" & vbCrLf & "Col D แถวที่ 5 = ชื่อสถานทีจัดส่ง, Col E แถวที่ 5 = เลทคิดขนส่ง, Col F แถวที่ 5 = เลทคิดลูกค้า, " & vbCrLf & " *** โดยห้ามให้รหัสลูกค้าและรหัสสถานทีจัดส่งซ้ำกัน ***" & vbCrLf & "*** เลทจัดส่งBAG และจัดส่งBULK แยกด้วยเครื่องหมาย *** เท่านั้น ***"
   Else
      Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
       str = "- เริ่ม Import ที่ Row ที่ 4 ,Column A ของ Sheet1 เท่านั้น โดย" & vbCrLf & "Col A แถวที่ 4 = รหัสวัตถุดิบ, Col B1 = วันที่มีผล, Col B2 = วันที่สิ้นสุด, Col B แถวที่ 4 = ชื่อวัตถุดิบ, Col C แถวที่ 4 = ราคา" & vbCrLf & " ***โดยห้ามให้เบอร์วัตถุดิบซ้ำกัน***"
   End If
   
   Call InitNormalLabel(lblNote, str)
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   If isSave Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set PartUctlColls = New Collection
   Set PartColls = New Collection
   Set m_Customers = New Collection
   Set m_DeliveryCus = New Collection
   Set temp_DeliveryCus = New Collection
   Set PartLabColls = New Collection
   Set PartLabUpdateColls = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set PartColls = Nothing
   Set m_Customers = Nothing
   Set m_DeliveryCus = Nothing
   Set temp_DeliveryCus = Nothing
   Set PartLabColls = Nothing
   Set PartLabUpdateColls = Nothing
   Set PartUctlColls = Nothing
End Sub
Private Function SearchLabCode(SearchItemNo As CPartItem, PartNo As String, PartName As String) As Boolean
   SearchLabCode = True
   Set SearchItemNo = GetObject("CPartItem", PartColls, PartNo, False)
   If SearchItemNo Is Nothing Then
      Set SearchItemNo = GetObject("CPartItem", PartLabColls, PartNo, False)
      If SearchItemNo Is Nothing Then
         Set SearchItemNo = GetObject("CPartItem", PartLabUpdateColls, Trim(PartNo), False)
         If SearchItemNo Is Nothing Then
            'LoadForm
            Set SearchItemNo = New CPartItem
            Set frmMapPlcProductItem.PartItem = SearchItemNo
'            Set frmMapPlcProductItem.mPartItemColl = PartUctlColls
            If Trim(PartNo) = Trim(PartName) Then
               MsgBox MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo)
            Else
               MsgBox MapText("MAP ข้อมูล รหัสสินค้า/วัตถุดิบ " & PartNo & "-" & PartName & "กรุณาติดต่อบัญชีให้เพิ่มรหัสข้อมูลเข้าระบบ")
            End If

'            Unload frmMapPlcProductItem
'            Set frmMapPlcProductItem = Nothing

            'AddDataTo PartPlcUpdateColls
            If Len(Trim(SearchItemNo.PART_NO)) <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบรหัสที่อ้างอิง วัตถุดิบ สำหรับ " & PartNo & "-" & PartName
               glbErrorLog.ShowUserError

               SearchLabCode = False
               Exit Function
            End If
            SearchItemNo.NUMBER_LAB_ID = Trim(PartNo)
'            Call PartLabUpdateColls.add(SearchItemNo, Trim(PartNo))
         End If
      End If
   End If
End Function

Private Sub uctlFromActiveDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToValidDate_HasChange()
   m_HasModify = True
End Sub
Private Function AutoGenName() As String
Dim No As String
            Call glbDatabaseMngr.GenerateNumber(EX_WORKS_PRICE, No, glbErrorLog)
            If Area = 1 Then
              No = "P" & No
            ElseIf Area = 2 Then
              No = "D" & No
            ElseIf Area = 3 Then
              No = "PP" & No
            ElseIf Area = 4 Then
              No = "PD" & No
            End If
            AutoGenName = No
End Function
