VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditReceiptChequeDoc 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   Icon            =   "frmAddEditReceiptChequeDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11850
   StartUpPosition =   1  'CenterOwner
   Begin prjFarmManagement.uctlTextLookup uctlBankBranch 
      Height          =   435
      Left            =   1680
      TabIndex        =   49
      Top             =   3360
      Width           =   4995
      _ExtentX        =   9499
      _ExtentY        =   767
   End
   Begin prjFarmManagement.uctlTextLookup uctlBank 
      Height          =   435
      Left            =   1680
      TabIndex        =   48
      Top             =   2880
      Width           =   4995
      _ExtentX        =   9499
      _ExtentY        =   767
   End
   Begin prjFarmManagement.uctlTextBox txtAmountCheque 
      Height          =   435
      Left            =   9360
      TabIndex        =   45
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   1575
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   -120
      TabIndex        =   14
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlPassChequeDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         Top             =   1920
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlDate uctlBadChequeDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   41
         Top             =   2400
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3360
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5927
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditReceiptChequeDoc.frx":27A2
         Column(2)       =   "frmAddEditReceiptChequeDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditReceiptChequeDoc.frx":290E
         FormatStyle(2)  =   "frmAddEditReceiptChequeDoc.frx":2A6A
         FormatStyle(3)  =   "frmAddEditReceiptChequeDoc.frx":2B1A
         FormatStyle(4)  =   "frmAddEditReceiptChequeDoc.frx":2BCE
         FormatStyle(5)  =   "frmAddEditReceiptChequeDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditReceiptChequeDoc.frx":2D5E
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2085
         Left            =   12960
         TabIndex        =   26
         Top             =   5520
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   3678
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboBankBranch 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1170
            Width           =   4035
         End
         Begin VB.ComboBox cboBank 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   720
            Width           =   4035
         End
         Begin VB.ComboBox cboPaymentType 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   270
            Width           =   2325
         End
         Begin prjFarmManagement.uctlTextBox txtCheckNo 
            Height          =   435
            Left            =   7470
            TabIndex        =   28
            Top             =   210
            Width           =   2625
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlCheckDate 
            Height          =   405
            Left            =   7470
            TabIndex        =   30
            Top             =   660
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin VB.Label lblCheckDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5940
            TabIndex        =   36
            Top             =   690
            Width           =   1395
         End
         Begin VB.Label lblBankBranch 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   35
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   34
            Top             =   810
            Width           =   1275
         End
         Begin VB.Label lblCheckNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5700
            TabIndex        =   33
            Top             =   270
            Width           =   1665
         End
         Begin VB.Label lblPaymentType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   32
            Top             =   360
            Width           =   1275
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2085
         Left            =   270
         TabIndex        =   20
         Top             =   5730
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3678
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjFarmManagement.uctlTextLookup uctlResource 
            Height          =   435
            Left            =   1740
            TabIndex        =   21
            Top             =   120
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtPaidFor 
            Height          =   435
            Left            =   1770
            TabIndex        =   24
            Top             =   1080
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlPaidDate 
            Height          =   405
            Left            =   7530
            TabIndex        =   23
            Top             =   570
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtPvNo 
            Height          =   435
            Left            =   1770
            TabIndex        =   22
            Top             =   600
            Width           =   2625
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin Threed.SSCommand cmdPvNo 
            Height          =   405
            Left            =   4440
            TabIndex        =   40
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditReceiptChequeDoc.frx":2F36
            ButtonStyle     =   3
         End
         Begin VB.Label lblPvNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   660
            Width           =   1545
         End
         Begin VB.Label lblPaidFor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   1140
            Width           =   1515
         End
         Begin VB.Label lblPaidDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6000
            TabIndex        =   37
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label lblResource 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            TabIndex        =   25
            Top             =   180
            Width           =   1635
         End
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlChequeDocDate 
         Height          =   405
         Left            =   6600
         TabIndex        =   2
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtChequeDocNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   930
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   1080
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   16
         Top             =   -120
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblBankBranchName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         TabIndex        =   51
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblBankName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   2880
         Width           =   1335
      End
      Begin Threed.SSCommand cmdEditStatusPass 
         Height          =   525
         Left            =   7440
         TabIndex        =   47
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblAmountCheque 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7800
         TabIndex        =   46
         Top             =   2400
         Width           =   1455
      End
      Begin Threed.SSCheck chkPassCheque 
         Height          =   435
         Left            =   6000
         TabIndex        =   6
         Top             =   1920
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPassChequeDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblBadChequeDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   1575
      End
      Begin Threed.SSCommand cmdCustomer 
         Height          =   405
         Left            =   7260
         TabIndex        =   5
         Top             =   1380
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceiptChequeDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4410
         TabIndex        =   1
         Top             =   930
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceiptChequeDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkBadCheque 
         Height          =   435
         Left            =   6000
         TabIndex        =   3
         Top             =   2400
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label lblChequeDocDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         Top             =   990
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceiptChequeDoc.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   10
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceiptChequeDoc.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3360
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceiptChequeDoc.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblChequeDocNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   990
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditReceiptChequeDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_ChequeDoc As CChequeDoc

Private m_Customers As Collection
Private m_Employees As Collection
Private m_Resources As Collection
Private m_Banks As Collection
Private m_BankBranchs As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long
Public ReceiptType As Long
Public Area As Long
Public DocumentType As Long
Public CUSTOMER_ID  As String
Public PassChequeFlag As Long

Private FileName As String
Private m_SumUnit As Double
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_ChequeDoc.CHEQUE_DOC_ID = id
      If Not glbDaily.QueryChequeDoc(m_ChequeDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ChequeDoc.PopulateFromRS(8, m_Rs)
      
      uctlChequeDocDate.ShowDate = m_ChequeDoc.CHEQUE_DOC_DATE
      txtChequeDocNo.Text = m_ChequeDoc.CHEQUE_DOC_NO
       uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_ChequeDoc.CUSTOMER_ID)
       uctlBadChequeDate.ShowDate = m_ChequeDoc.BADCHEQUE_DATE
       uctlPassChequeDate.ShowDate = m_ChequeDoc.PASSCHEQUE_DATE
       chkPassCheque.Value = FlagToCheck(m_ChequeDoc.PASSCHEQUE_FLAG)
       chkBadCheque.Value = FlagToCheck(m_ChequeDoc.BADCHEQUE_FLAG)
       txtAmountCheque.Text = m_ChequeDoc.AMOUNT_CHEQUE
       PassChequeFlag = FlagToCheck(m_ChequeDoc.PASSCHEQUE_FLAG)
       uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, m_ChequeDoc.BANK_ID)
       uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, m_ChequeDoc.BANK_BRANCH_ID)
      
       If ShowMode = SHOW_EDIT And FlagToCheck(m_ChequeDoc.PASSCHEQUE_FLAG) = 1 Then
       ' เมื่ออยู่ใน SHOW_EDIT  และอาจเช็คผิดพลาด ว่าสถานะCheque ผ่าน แล้วเผลอ save ต้องการแก้ไข ต้องเป็นผู้มีสิทธิ์เท่านั้น ใส่ user password
             chkPassCheque.Enabled = False
         End If
      
      Call EnableDisableButton(True)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
      
   Call EnableForm(Me, True)
End Sub


Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CDoItem

   For Each Di In BD.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CDoItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CLotItem
Dim Ei As CLotItem

      For Each Ei In Ivd.ImportExports
         If Ei.LINK_ID = GuiID Then
            Set GetExportItem = Ei
            Exit Function
         End If
      Next Ei
End Function

Private Function VerifyJournalItem(BD As CBillingDoc) As Boolean
Dim Gl As CGLDetail
Dim SumDr As Double
Dim SumCr As Double

   SumDr = 0
   SumCr = 0
   For Each Gl In BD.GlDetails
      If Gl.Flag <> "D" Then
         If Gl.GetFieldValue("GL_TYPE") = 1 Then
            SumDr = SumDr + Gl.GetFieldValue("GL_AMOUNT")
         ElseIf Gl.GetFieldValue("GL_TYPE") = 2 Then
            SumCr = SumCr + Gl.GetFieldValue("GL_AMOUNT")
         End If
      End If
   Next Gl
   
   If FormatNumber(SumDr) <> FormatNumber(SumCr) Then
      VerifyJournalItem = False
   Else
      VerifyJournalItem = True
   End If
End Function

Private Function MyCountItem(Col As Collection) As Long
Dim I As Long
Dim Count As Long
Dim Ji As CCashTran

   Count = 0
   For I = 1 To Col.Count
      Set Ji = Col.Item(I)
      If (Ji.Flag <> "D") And (Ji.Cheque.GetFieldValue("EFFECTIVE_DATE") > 0) Then
         Count = Count + 1
      End If
   Next I
   
   MyCountItem = Count
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
Dim ct As CCashTran
'Dim RCD As CReceiptChequeDoc
Dim CR As CReceiptChequeDoc
Dim Sum1 As Double
Dim CheckDate As Long
Dim CheckDate2 As Long
'Public PassCheqDate As Long
'Public CheqDate As Long


  If Not CheckUniqueNs(CHEQUEDOC_UNIQUE, txtChequeDocNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtChequeDocNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If ShowMode = SHOW_EDIT And PassChequeFlag = 1 Then 'chkPassCheque.Enabled = False
      glbErrorLog.LocalErrorMsg = MapText("มีเอกสารที่อ้างอิงเช็คใบนี้ ไม่สามารถแก้ไขข้อมูลได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   
   

'   If ShowMode = SHOW_EDIT Then
'      If Area = 1 Then
'         If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
'      ElseIf Area = 2 Then
'         If Not VerifyAccessRight("LEDGER_BUY" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
'      End If
'   End If
   If Not VerifyTextControl(lblChequeDocNo, txtChequeDocNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblChequeDocDate, uctlChequeDocDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
     If Not VerifyTextControl(lblAmountCheque, txtAmountCheque, False) Then
      Exit Function
   End If
   
    If Not VerifyCombo(lblBankName, uctlBank.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankBranchName, uctlBankBranch.MyCombo, False) Then
      Exit Function
   End If
'    If chkBadCheque.Value = 0 And chkPassCheque.Value = 0 Then
'      glbErrorLog.LocalErrorMsg = "กรุณาเลือกใส่การตรวจสอบเช็ค ผ่าน หรือ ไม่ผ่าน "
'      glbErrorLog.ShowUserError
'      Exit Function
'    End If


  If chkBadCheque.Value = 1 Then
    If Not VerifyDate(lblBadChequeDate, uctlBadChequeDate, False) Then
      Exit Function
   End If

   CheckDate2 = DateDiff("D", uctlChequeDocDate.ShowDate, uctlBadChequeDate.ShowDate)
 
 If CheckDate2 < 0 Then
    glbErrorLog.LocalErrorMsg = MapText("กรุณาตรวจสอบ วันที่คืนเช็ค ควรมากว่าหรือเท่าวันที่ของเช็ค")
   glbErrorLog.ShowUserError
   Exit Function
End If


End If

  If chkPassCheque.Value = 1 Then
  
   If Not VerifyDate(lblPassChequeDate, uctlPassChequeDate, False) Then
      Exit Function
    End If
    
  CheckDate = DateDiff("D", uctlChequeDocDate.ShowDate, uctlPassChequeDate.ShowDate)
 If CheckDate < 0 Then
    glbErrorLog.LocalErrorMsg = MapText("กรุณาตรวจสอบ วันที่ผ่านเช็ค ควรมากว่าหรือเท่าวันที่ของเช็ค")
   glbErrorLog.ShowUserError
   Exit Function
 End If
End If


  

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   




   m_ChequeDoc.AddEditMode = ShowMode
   m_ChequeDoc.CHEQUE_DOC_ID = id
   m_ChequeDoc.CHEQUE_DOC_DATE = uctlChequeDocDate.ShowDate
   m_ChequeDoc.CHEQUE_DOC_NO = txtChequeDocNo.Text
   m_ChequeDoc.BADCHEQUE_FLAG = Check2Flag(chkBadCheque.Value)
   m_ChequeDoc.PASSCHEQUE_FLAG = Check2Flag(chkPassCheque.Value)
   m_ChequeDoc.PASSCHEQUE_DATE = uctlPassChequeDate.ShowDate
   m_ChequeDoc.BADCHEQUE_DATE = uctlBadChequeDate.ShowDate
   m_ChequeDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_ChequeDoc.AMOUNT_CHEQUE = Val(txtAmountCheque.Text)
    m_ChequeDoc.BANK_ID = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
     m_ChequeDoc.BANK_BRANCH_ID = uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex))
'  For Each CR In m_ChequeDoc.ChequeDoc
''    ''Debug.Print (CR.PAID_AMOUNT)
'    Sum1 = Sum1 + CR.PAID_AMOUNT
' Next CR


    For Each CR In m_ChequeDoc.ChequeDoc
      If CR.Flag <> "D" Then
        Sum1 = Sum1 + CR.PAID_AMOUNT
     End If
   Next CR
   
    If Val(txtAmountCheque.Text) < Sum1 Then
     glbErrorLog.LocalErrorMsg = "ไม่สามารถบันทึกข้อมูลได้เนื่องจากยอดเช็คที่ชำระน้อยกว่าบิล "
      glbErrorLog.ShowUserError
      Exit Function
    ElseIf Val(txtAmountCheque.Text) > Sum1 Then
     glbErrorLog.LocalErrorMsg = MapText("ใส่ยอดชำระเกินยอดบิล")
      glbErrorLog.ShowUserError
      Exit Function
    End If
   
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditChequeDoc(m_ChequeDoc, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If

   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

'Private Sub cboBank_Click()
'Dim BankID As Long
'
'   BankID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
'   If BankID > 0 Then
'      Call LoadBankBranch(cboBankBranch, , BankID)
'   End If
'
'   m_HasModify = True
'End Sub
'
'Private Sub cboBankBranch_Click()
'   m_HasModify = True
'End Sub

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub


Private Sub uctlBank1_Change()
Dim TempID As Long
Dim BB As CBankBranch
   TempID = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
   
   If TempID > 0 Then
      Call LoadBankBranch(uctlBankBranch.MyCombo, m_BankBranchs, TempID)
      Set uctlBankBranch.MyCollection = m_BankBranchs
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlBankAccountLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlBankBranch_Change()
Dim TempID1 As Long
Dim TempID2 As Long
   
   TempID1 = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
   TempID2 = uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex))
   
'   If TempID2 > 0 Then
'      Call LoadMaster(uctlBankAccountLookup.MyCombo, m_BankAccounts, BANK_ACCOUNT, TempID1, TempID2)
'      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
'   End If
   
   m_HasModify = True
End Sub

Private Sub cboPaymentType_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Public Sub RefreshGrid()
   Call GetTotalPrice

   GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
   GridEX1.Rebind
End Sub

Private Sub chkPayFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkBadCheque_Click(Value As Integer)
  m_HasModify = True
  chkPassCheque.Value = 0
  If Value = 1 Then
    chkPassCheque.Enabled = False
'    uctlPassChequeDate.Enable = False
  Else
     chkPassCheque.Enabled = True
'     uctlPassChequeDate.Enable = True
   
  End If

End Sub

Private Sub chkPassCheque_Click(Value As Integer)
  m_HasModify = True
  chkBadCheque.Value = 0
  If Value = 1 Then
     chkBadCheque.Enabled = False
'     uctlBadChequeDate.Enable = False
  Else
     chkBadCheque.Enabled = True
'     uctlBadChequeDate.Enable = True
  End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
         
 If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo) Then
         Exit Sub
   End If


      If TabStrip1.SelectedItem.Index = 1 Then
  
        frmAddReceiptDocItem.Area = Area
         frmAddReceiptDocItem.CustomerID = CUSTOMER_ID
         Set frmAddReceiptDocItem.TempCollection = m_ChequeDoc.ChequeDoc '''''ลอง
          frmAddReceiptDocItem.ShowMode = SHOW_ADD
         frmAddReceiptDocItem.HeaderText = MapText("เพิ่มรายการเช็ค")
         Load frmAddReceiptDocItem
          frmAddReceiptDocItem.Show 1
   
         OKClick = frmAddReceiptDocItem.OKClick
   
         Unload frmAddReceiptDocItem
         Set frmAddReceiptDocItem = Nothing
   
   
        If OKClick Then
            Call GetTotalPriceCheq
            GridEX1.ItemCount = CountItem(m_ChequeDoc.ChequeDoc)
            GridEX1.Rebind
          End If
          
          
'       ElseIf TabStrip1.SelectedItem.Index = 2 Then
'
''        frmAddReceiptDocItem.Area = Area
''         frmAddReceiptDocItem.CustomerID = CUSTOMER_ID
''         Set frmAddReceiptDocItem.TempCollection = m_ChequeDoc.BankInfo '''''ลอง
'          frmAddEditBankChequeDoc.ShowMode = SHOW_ADD
'         frmAddEditBankChequeDoc.HeaderText = MapText("เพิ่มรายละเอียดธนาคาร")
'         Load frmAddEditBankChequeDoc
'          frmAddEditBankChequeDoc.Show 1
'
'         OKClick = frmAddEditBankChequeDoc.OKClick
'
'         Unload frmAddEditBankChequeDoc
'         Set frmAddEditBankChequeDoc = Nothing
'
'
'        If OKClick Then
'            GridEX1.ItemCount = CountItem(m_ChequeDoc.BankInfo)
'            GridEX1.Rebind
'          End If
          
 End If
          
   If OKClick Then
      m_HasModify = True
   End If
End Sub
Private Sub cmdAuto_Click()
Dim No As String

'   If Trim(txtChequeDocNo.Text) = "" Then
'      Call glbDatabaseMngr.GenerateNumber(CHEQUE_DOC_NUMBER, No, glbErrorLog)
'      txtDocumentNo.Text = No
'   End If
End Sub

Private Sub cmdCustomer_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CCustomer

   Set TempCol = New Collection
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ค้นหา", "-", "เพิ่มข้อมูลใหม่")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      Set frmQueryCustomer.TempCollection = TempCol
      frmQueryCustomer.ShowMode = SHOW_ADD
      Load frmQueryCustomer
      frmQueryCustomer.Show 1
      
      OKClick = frmQueryCustomer.OKClick
      
      Unload frmQueryCustomer
      Set frmQueryCustomer = Nothing
      
      If OKClick Then
         Set Cs = TempCol(1)
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, Cs.CUSTOMER_ID)
         m_HasModify = True
      End If
   ElseIf lMenuChosen = 3 Then
      frmAddEditCustomer.ShowMode = SHOW_ADD
      frmAddEditCustomer.HeaderText = MapText("เพิ่มลูกค้า")
      Load frmAddEditCustomer
      frmAddEditCustomer.Show 1
      
      OKClick = frmAddEditCustomer.OKClick
      Call EnableForm(Me, False)
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      Call EnableForm(Me, True)
      
      Unload frmAddEditCustomer
      Set frmAddEditCustomer = Nothing
   End If
   
   Set TempCol = Nothing

End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
         If ID1 <= 0 Then
            m_ChequeDoc.ChequeDoc.Remove (ID2)
         Else
            m_ChequeDoc.ChequeDoc.Item(ID2).Flag = "D"
         End If
   
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_ChequeDoc.ChequeDoc)
         GridEX1.Rebind
         m_HasModify = True

   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
         
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Area = 1 Then
      If Not VerifyDate(lblChequeDocDate, uctlChequeDocDate) Then
         Exit Sub
      End If
   End If
   
   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ReceiptType = 1 Then

         frmAddEditDoItem.DocumentDate = uctlChequeDocDate.ShowDate
         frmAddEditDoItem.SubscriberID = -1
         frmAddEditDoItem.Area = Area
         frmAddEditDoItem.id = id
         frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem.HeaderText = MapText("แก้ไขรายการใบเสร็จ")
         frmAddEditDoItem.ParentShowMode = ShowMode
         frmAddEditDoItem.ShowMode = SHOW_EDIT
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
   
         OKClick = frmAddEditDoItem.OKClick
   
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
   
         If OKClick Then
            Call GetTotalPriceEx
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      End If
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'      frmAddEditCashTran.Area = Area
'      Set frmAddEditCashTran.ParentForm = Me
'      frmAddEditCashTran.ID = ID
'      frmAddEditCashTran.HeaderText = "แก้ไขรายการการชำระเงิน"
'      frmAddEditCashTran.ShowMode = SHOW_EDIT
'      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
'      Load frmAddEditCashTran
'      frmAddEditCashTran.Show 1
'
'      OKClick = frmAddEditCashTran.OKClick
'
'      Unload frmAddEditCashTran
'      Set frmAddEditCashTran = Nothing
'
'      If OKClick Then
'         m_HasModify = True
'
'         GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
'         Call GridEX1.Rebind
'
'         Call GetTotalPrice
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'      Set frmAddEditGlDetail.ParentForm = Me
'      frmAddEditGlDetail.ID = ID
'      frmAddEditGlDetail.HeaderText = "แก้ไขรายการสมุดรายวัน"
'      frmAddEditGlDetail.ShowMode = SHOW_EDIT
'      Set frmAddEditGlDetail.TempCollection = m_BillingDoc.GlDetails
'      Load frmAddEditGlDetail
'      frmAddEditGlDetail.Show 1
'
'      OKClick = frmAddEditGlDetail.OKClick
'
'      Unload frmAddEditGlDetail
'      Set frmAddEditGlDetail = Nothing
'
'      If OKClick Then
'         m_HasModify = True
'
'         GridEX1.ItemCount = CountItem(m_BillingDoc.GlDetails)
'         Call GridEX1.Rebind
'
'         Call GetTotalPrice
'      End If
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub CalculateIncludePrice()
Dim II As CLotItem
Dim AvgFee As Double

'   If m_SumUnit > 0 Then
'      AvgFee = Val(txtTotalAmount.Text) / m_SumUnit
'   Else
'      AvgFee = 0
'   End If
'
'   For Each II In m_BillingDoc.DoItems
'      If II.Flag <> "D" Then
'         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + AvgFee
'      End If
'   Next II
End Sub

Private Sub cmdEditStatusPass_Click()
  frmVerifyAccRight.AccName = "STATUS_CONTROL"
   frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงสถานะการผ่านของเช็คได้"
   Load frmVerifyAccRight
   frmVerifyAccRight.Show 1
   
   If frmVerifyAccRight.GrantRight Then
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
     PassChequeFlag = 0
   Else
      Unload frmVerifyAccRight
      Set frmVerifyAccRight = Nothing
      Exit Sub
   End If
            
   chkPassCheque.Enabled = True
   
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Function VerifyOnwerVersionMenu(Menu As Long, Owner As String) As Boolean
   VerifyOnwerVersionMenu = True
   
   If (Menu <> 1) And (Menu <> 2) Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_RECEIPT_PREFORM_PRINT", True) Then
         VerifyOnwerVersionMenu = False
         Exit Function
      End If
   End If
End Function

'Private Sub cmdPrint_Click()
'Dim lMenuChosen As Long
'Dim oMenu As cPopupMenu
'Dim ReportFlag As Boolean
'Dim ReportKey As String
'Dim Report As CReportInterface
'Dim Rc As CReportConfig
'Dim iCount As Long
'Dim EditMode As SHOW_MODE_TYPE
'Dim ReportMode As Long
'Dim Programowner As String
'   Programowner = glbParameterObj.Programowner
'
'   ReportMode = 1
'
'   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
'      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
'      glbErrorLog.ShowUserError
'      Exit Sub
'   End If
'
'   ReportFlag = False
'
'   Call LoadPictureFromFile(glbParameterObj.ReceiptPicture1, Picture2)
'
'   Set oMenu = New cPopupMenu
'   If DocumentType = 8 Then
'      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItemsSpacialBuy)
'   Else
'      If ReceiptType = 1 Then
'         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItems)
'      Else
'         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.RCPrintMenuItemsSpacial)
'      End If
'   End If
'   If lMenuChosen = 0 Then
'      Exit Sub
'   End If
'   Set oMenu = Nothing
'
'
'      If lMenuChosen = 1 Then
'      ReportKey = "CReportNormalRcp001"
'
'      Set Report = New CReportNormalRcp001
'      ReportFlag = True
'   ElseIf lMenuChosen = 2 Then
'      ReportKey = "CReportNormalRcp001"
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   ElseIf lMenuChosen = 4 Then
'      ReportKey = "CReportFormReceipt001"
'
'      Set Report = New CReportFormReceipt001
'      Call Report.AddParam(1, "DO_TYPE")
'
'      ReportFlag = True
'   ElseIf lMenuChosen = 5 Then
'      ReportKey = "CReportFormReceipt001"
'
'      Set Report = New CReportFormReceipt001
'    Call Report.AddParam(1, "DO_TYPE")
'      ReportFlag = True
'      ElseIf lMenuChosen = 6 Then
'      ReportKey = "CReportFormReceipt001"
'
'      Set Report = New CReportFormReceipt001
'      Call Report.AddParam(0, "DO_TYPE")
'      ReportFlag = True
'   ElseIf lMenuChosen = 7 Then
'      ReportKey = "CReportFormReceipt001"
'
'      Set Report = New CReportFormReceipt001
'      Call Report.AddParam(0, "DO_TYPE")
'      ReportFlag = True
'   ElseIf lMenuChosen = 8 Then
'      ReportKey = "CReportFormReceipt001"
'      ReportMode = 2
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   ElseIf lMenuChosen = 10 Then
'      ReportKey = "CReportNormalRcpHead"
'      Set Report = New CReportNormalRcpHead
'      ReportFlag = True
'   ElseIf lMenuChosen = 11 Then
'      ReportKey = "CReportNormalRcpHead"
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   ElseIf lMenuChosen = 22 Then
'      ReportKey = "CReportVoucherReceive"
'      Set Report = New CReportVoucherReceive
'      ReportFlag = True
'   ElseIf lMenuChosen = 23 Then
'      ReportKey = "CReportVoucherReceive"
'      ReportMode = 2
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบสำคัญรับ")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   ElseIf lMenuChosen = 25 Then
'      ReportKey = "CReportVoucherPay"
'      Set Report = New CReportVoucherPay
'      ReportFlag = True
'   ElseIf lMenuChosen = 26 Then
'      ReportKey = "CReportVoucherPay"
'      ReportMode = 2
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบสำคัญรับ")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   End If
'
'   If Not Report Is Nothing Then
'      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
'      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
'      Call Report.AddParam(ReportKey, "REPORT_KEY")
'      Call Report.AddParam(MapText("ใบเสร็จรับเงิน"), "REPORT_HEADER")
'      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
'      Call Report.AddParam("", "ACCEPT_NAME")
'   ElseIf lMenuChosen = 5 Then
'      ReportKey = "CReportFormPO001"
'
'      Set Rc = New CReportConfig
'      Rc.REPORT_KEY = ReportKey
'      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
'      Call Rc.QueryData(m_Rs, iCount)
'      HeaderText = MapText("ใบเสร็จรับเงิน")
'      If Not m_Rs.EOF Then
'         Call Rc.PopulateFromRS(1, m_Rs)
'         EditMode = SHOW_EDIT
'      Else
'         EditMode = SHOW_ADD
'      End If
'   End If
'
'   Call EnableForm(Me, False)
'   If ReportFlag Then
'      Set frmReport.ReportObject = Report
'      frmReport.HeaderText = ""
'      Load frmReport
'      frmReport.Show 1
'
'      Unload frmReport
'      Set frmReport = Nothing
'      Set Report = Nothing
'   Else
'      frmReportConfig.ReportMode = ReportMode
'      frmReportConfig.ShowMode = EditMode
'      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
'      frmReportConfig.ReportKey = ReportKey
'      frmReportConfig.HeaderText = HeaderText
'      Load frmReportConfig
'      frmReportConfig.Show 1
'
'      Unload frmReportConfig
'      Set frmReportConfig = Nothing
'   End If
'   Call EnableForm(Me, True)
'End Sub

Private Sub cmdPvNo_Click()
   ''Debug.Print
   
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_ChequeDoc.CHEQUE_DOC_ID
   m_ChequeDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
        Call EnableForm(Me, False)
        Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
  
  
       Call LoadBank(uctlBank.MyCombo, m_Banks)
      Set uctlBank.MyCollection = m_Banks
      
      Call LoadBankBranch(uctlBankBranch.MyCombo, m_BankBranchs)
      Set uctlBankBranch.MyCollection = m_BankBranchs
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_ChequeDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlChequeDocDate.ShowDate = Now
         m_ChequeDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_Resources = Nothing
   Set m_Banks = Nothing
   Set m_BankBranchs = Nothing
   
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2415
   Col.Caption = MapText("เลขที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2250
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2460
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน")

'   Set Col = GridEX1.Columns.add '6
'   Col.Width = 1920
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("ส่วนลดเงินสด")
End Sub
'Private Sub InitGrid2()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX1.Columns.add '3
'   Col.Width = 2415
'   Col.Caption = MapText("ธนาคาร")
'
'   Set Col = GridEX1.Columns.add '4
'   Col.Width = 2250
'   Col.Caption = MapText("สาขา")
'
'
'End Sub
Private Sub GetTotalPrice()
Dim II As CReceiptItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum7 As Double
Dim Pm As CCashTran

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   For Each II In m_BillingDoc.ReceiptChequeDocItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.PAID_AMOUNT
         Sum2 = Sum2 + II.VAT_AMOUNT
         Sum3 = Sum3 + II.DISCOUNT_AMOUNT
         Sum4 = Sum4 + II.DEPOSIT_AMOUNT
      End If
   Next II
   
   Sum7 = 0
   For Each Pm In m_BillingDoc.Payments
      Sum7 = Sum7 + Pm.GetFieldValue("AMOUNT") - Pm.GetFieldValue("INTERREST_PAY") + Pm.GetFieldValue("WH_PAY")
   Next Pm
   
   
   
   

End Sub


Private Sub GetTotalPriceEx()
Dim II As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum2 = 0
   Sum1 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.TOTAL_PRICE
         Sum2 = Sum2 + II.DISCOUNT_AMOUNT
         Sum3 = Sum3 + II.DEPOSIT_AMOUNT
      End If
   Next II

   
End Sub
Private Sub GetTotalPriceCheq()
Dim II As CReceiptChequeDoc
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum2 = 0
   Sum1 = 0
   For Each II In m_ChequeDoc.ChequeDoc
      If II.Flag <> "D" Then
'         Sum1 = Sum1 + II.TOTAL_PRICE
'         Sum2 = Sum2 + II.DISCOUNT_AMOUNT
         Sum1 = Sum1 + II.PAID_AMOUNT
      End If
   Next II

   txtAmountCheque.Text = Sum1
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblChequeDocNo, MapText("เลขที่ใบเช็ค"))
   

   Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
      cmdAuto.Visible = False
      cmdCustomer.Visible = True

   Call InitNormalLabel(lblChequeDocDate, MapText("วันที่เช็ค"))
   Call InitNormalLabel(lblBadChequeDate, MapText("วันที่คืนเช็ค"))
   Call InitNormalLabel(lblPassChequeDate, MapText("วันที่เช็คผ่าน"))
   Call InitNormalLabel(lblAmountCheque, MapText("ยอดรวมเช็ค"))
   Call InitNormalLabel(lblBankName, MapText("ธนาคาร"))
   Call InitNormalLabel(lblBankBranchName, MapText("สาขา"))
   
   Call InitCheckBox(chkBadCheque, "เช็คไม่ผ่าน")
   Call InitCheckBox(chkPassCheque, "เช็คผ่าน")
   
   
   
   
   
   If Area = 1 Then

      chkPassCheque.Visible = True
      chkBadCheque.Visible = True

   End If
   
   Call txtChequeDocNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
    Call txtAmountCheque.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True
   SSFrame2.Visible = False
   SSFrame3.Visible = False
   

   Call InitCombo(cboPaymentType)
'   Call InitCombo(cboBank)
'   Call InitCombo(cboBankBranch)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCustomer.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPvNo.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditStatusPass.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
'   Call InitMainButton(cmdSave, MapText("บันทึก"))
'   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCustomer, MapText("F"))
   Call InitMainButton(cmdPvNo, MapText("P"))
   Call InitMainButton(cmdEditStatusPass, MapText("แก้ไข STATUS"))
  
    If ShowMode = SHOW_EDIT Then
    ' เมื่ออยู่ใน SHOW_EDIT คือหลังจากมีการ Save แล้ว หรือ เข้าสู่การแก้ไข จะล็อกการเปลี่ยนแปลงcustomer เพื่อ ไม่ให้มีการเปลี่ยนแปลงข้างในที่หลังเมื่อมีบิล หรือ การตัดบิลไปแล้ว
   ' หากต้องการเปลี่ยนแปลง หรือ มีข้อผิดพลาดให้ลบ ใบนี้ทิ้งแล้ว สร้างใหม่
       uctlCustomerLookup.Enabled = False
       cmdCustomer.Enabled = False
    
    Else
       uctlCustomerLookup.Enabled = True
       cmdCustomer.Enabled = True
    End If


   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการใบเสร็จ")
'    TabStrip1.Tabs.add().Caption = MapText("เพิ่มเติม")
 Call InitGrid1
' Call InitGrid2

   
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_BillingDoc = New CBillingDoc
  Set m_ChequeDoc = New CChequeDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   Set m_Resources = New Collection
    Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   If TabStrip1.SelectedItem.Index = 5 Then
'      RowBuffer.RowStyle = RowBuffer.Value(7)
'   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

      
       If TabStrip1.SelectedItem.Index = 1 Then
      If m_ChequeDoc.ChequeDoc Is Nothing Then
         Exit Sub
      End If


      If RowIndex <= 0 Then
         Exit Sub
      End If
      

         Dim RCD As CReceiptChequeDoc
         If m_ChequeDoc.ChequeDoc.Count <= 0 Then
            Exit Sub
         End If
         Set RCD = GetItem(m_ChequeDoc.ChequeDoc, RowIndex, RealIndex)
         If RCD Is Nothing Then
            Exit Sub
         End If
   
   
   
   
         Values(1) = RCD.RECEIPT_CHEQUE_DOC_ID 'RCD.BILLING_DOC_ID
         Values(2) = RealIndex
         Values(3) = RCD.RECEIPT_CHEQUE_DOC_NO
          Values(4) = DateToStringExtEx2(RCD.RECEIPT_CHEQUE_DOC_DATE)
         Values(5) = FormatNumber(RCD.PAID_AMOUNT)
'         Values(6) = FormatNumber(RCD.CASH_DISCOUNT)
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If m_ChequeDoc.BankInfo Is Nothing Then
'         Exit Sub
'      End If
'
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'
'         Dim Cd As CChequeDoc
'         If m_ChequeDoc.BankInfo.Count <= 0 Then
'            Exit Sub
'         End If
'         Set Cd = GetItem(m_ChequeDoc.BankInfo, RowIndex, RealIndex)
'         If Cd Is Nothing Then
'            Exit Sub
'         End If
'
'
'
'
'         Values(1) = Cd.RECEIPT_CHEQUE_DOC_ID 'RCD.BILLING_DOC_ID
'         Values(2) = RealIndex
'         Values(3) = Cd.BANK_NAME
'          Values(4) = Cd.BANK_BRANCH_NAME
'

    End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub EnableDisableButton(En As Boolean)
   If En Then
'      If ShowMode = SHOW_EDIT Then
''         cmdAdd.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
''         cmdEdit.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
''         cmdDelete.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
''      Else
''         cmdAdd.Enabled = True
''         cmdDelete.Enabled = True
'      End If
'      If ((ReceiptType = 3) Or (ReceiptType = 5)) And Not (TabStrip1.SelectedItem.Index = 3 Or TabStrip1.SelectedItem.Index = 4) Then
'         cmdEdit.Enabled = False
'      Else
'         cmdEdit.Enabled = True
'      End If
'   Else
'      cmdAdd.Enabled = En
'      cmdDelete.Enabled = En
'      cmdEdit.Enabled = En
'
      cmdAdd.Enabled = En
      cmdDelete.Enabled = En
      cmdEdit.Enabled = En
      
      If ShowMode = SHOW_EDIT Then
         cmdEdit.Enabled = False
      End If
      
      
   End If
End Sub



Private Sub TabStrip1_Click()
'   GridEX1.Top = 5670
'   GridEX1.Left = 150
'   GridEX1.Visible = False
'
'   SSFrame2.Top = 5670
'   SSFrame2.Left = 150
'   SSFrame2.Visible = False
'
'   SSFrame3.Top = 5670
'   SSFrame3.Left = 150
'   SSFrame3.Visible = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Call EnableDisableButton(True)
      GridEX1.Visible = True
         Call GetTotalPrice
         Call InitGrid1
         GridEX1.ItemCount = CountItem(m_ChequeDoc.ChequeDoc)
         GridEX1.Rebind
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Call EnableDisableButton(True)
'      GridEX1.Visible = True
'
'         Call InitGrid2
'         GridEX1.ItemCount = CountItem(m_ChequeDoc.BankInfo)
'         GridEX1.Rebind

   End If
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtAmountCheque_Change()
 m_HasModify = True
' Call GetTotalPriceCheq
End Sub

Private Sub txtCheckNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeposit_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtDocumentNo_Change()
'   txtPvNo.Text = txtDocumentNo.Text
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtIncludeDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtIncludeVat_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtIncludeWH_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtChequeDocNo_Change()
  m_HasModify = True
End Sub

Private Sub txtPaidFor_Change()
   m_HasModify = True
End Sub

Private Sub txtPvNo_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalRcp_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtVatAmount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub CalculateAmount()
'   txtIncludeDiscount.Text = Val(txtNetTotal.Text) - Val(Replace(txtDiscount.Text, ",", ""))
   If (ReceiptType <> 5) Then
'      txtVatAmount.Text = Val(txtVatPercent.Text) * Val(Replace(txtIncludeDiscount.Text, ",", "")) / 100
   End If
'   txtIncludeVat.Text = Val(Replace(txtIncludeDiscount.Text, ",", "")) + Val(txtVatAmount.Text)
'   txtWHAmount.Text = Val(txtWH.Text) * Val(txtIncludeDiscount.Text) / 100
'   txtIncludeWH.Text = Val(txtIncludeVat.Text) - txtWHAmount.Text
'   txtDipRcp.Text = Val(txtIncludeWH.Text) - Val(txtTotalRcp.Text)
End Sub

Private Sub txtVatPercent_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtWH_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtWHAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlBadChequeDate_HasChange()
  m_HasModify = True
End Sub

Private Sub uctlCheckDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlChequeDocDate_HasChange()
  m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long
Dim Customer As CCustomer

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   CUSTOMER_ID = CustomerID
   If CustomerID > 0 Then
'      If Area = 1 Then
         Set Customer = m_Customers(Trim(str(CustomerID)))
'         Call LoadAccount(cboAccount, , CustomerID)
'         cboAccount.ListIndex = 1
   
'         Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
'         If Customer.RESPONSE_BY > 0 Then
'            uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, Customer.RESPONSE_BY)
'         Else
'            uctlSellByLookup.MyCombo.ListIndex = -1
'         End If
'      ElseIf Area = 2 Then
''         Call LoadAccount(cboAccount, , CustomerID)
''         cboAccount.ListIndex = -1
'
'         Call LoadSupplierAddress(cboCustomerAddress, , CustomerID, True)
'      End If
   Else
'      cboAccount.ListIndex = -1
'      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlPaidDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPassChequeDate_HasChange()
  m_HasModify = True

End Sub

Private Sub uctlResource_Change()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
