VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdjustSellPrice 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   Icon            =   "frmAdjustSellPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   11730
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4965
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   8758
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboGroup 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1320
         Width           =   3495
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   3210
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   465
         Left            =   30
         TabIndex        =   11
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   820
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1920
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtCustomer 
         Height          =   465
         Left            =   1920
         TabIndex        =   2
         Top             =   2280
         Width           =   1695
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   465
         Left            =   1920
         TabIndex        =   1
         Top             =   1800
         Width           =   1695
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   555
         Left            =   1920
         TabIndex        =   20
         Top             =   720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   979
         _Version        =   131073
         CaptionStyle    =   1
         Begin Threed.SSOption radFeature 
            Height          =   375
            Left            =   30
            TabIndex        =   23
            Top             =   90
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption4"
         End
         Begin Threed.SSOption radStock 
            Height          =   375
            Left            =   1950
            TabIndex        =   22
            Top             =   90
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption4"
         End
         Begin Threed.SSOption radCustom 
            Height          =   375
            Left            =   3960
            TabIndex        =   21
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   131073
            Caption         =   "SSOption4"
         End
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   465
         Left            =   7440
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   465
         Left            =   4920
         TabIndex        =   27
         Top             =   1800
         Width           =   6135
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   10440
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3720
         TabIndex        =   26
         Top             =   1920
         Width           =   1095
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   435
         Left            =   11115
         TabIndex        =   25
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustSellPrice.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5760
         TabIndex        =   24
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5760
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1920
         TabIndex        =   6
         Top             =   4140
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustSellPrice.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   15
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   3270
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   3690
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   2880
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9975
         TabIndex        =   7
         Top             =   4140
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAdjustSellPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private ExcelColl As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.xls)|*.xls;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdStart_Click()
On Error GoTo ErrorHandler

Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean

Dim Mn As CMenuItem
Dim ErrorObj As clsErrorLog
   
   Call GenerateExcelColl
   
   Call glbDaily.StartTransaction
   
   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   I = 0
    
   prgProgress.Value = 0
   txtPercent.Text = "LOADING"
   txtPercent.Refresh
   
   
   Dim Bd As CDoItem
   Dim Rs As ADODB.Recordset
   
   Set Bd = New CDoItem
   Set Rs = New ADODB.Recordset
   
   Bd.FROM_DATE = uctlFromDate.ShowDate
   Bd.TO_DATE = uctlToDate.ShowDate
   Bd.DISPLAY_ID = GetDisplayID
   Bd.CUSTOMER_CODE = txtCustomer.Text
   Bd.DOCUMENT_NO = txtDocumentNo.Text
   Bd.SET_PRODUCT_ID = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
   Call Bd.QueryData(33, Rs, iCount)
   
   I = 0
   
   While Not Rs.EOF
         I = I + 1
         Percent = MyDiffEx(I, iCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         txtPercent.Refresh
         Call Bd.PopulateFromRS(33, Rs)
         
         If Len(txtFileName.Text) > 0 Then
            Set Mn = GetObject("CMenuItem", ExcelColl, Trim(Bd.DOCUMENT_NO), False)
            If Not (Mn Is Nothing) Then
               Bd.AVG_PRICE = Bd.AVG_PRICE + Val(txtPrice.Text)
               Bd.PRICE_PER_PACK = Val(Format(Bd.AVG_PRICE * Bd.WEIGHT_PER_PACK, "0.00"))
               Bd.TOTAL_PRICE = Val(Format(Bd.AVG_PRICE * Bd.ITEM_AMOUNT, "0.00")) - Bd.DISCOUNT_AMOUNT - Bd.EXTRA_DISCOUNT
               Call Bd.UpdateAvgSellPrice
            End If
         Else
            Bd.AVG_PRICE = Bd.AVG_PRICE + Val(txtPrice.Text)
            Bd.PRICE_PER_PACK = Val(Format(Bd.AVG_PRICE * Bd.WEIGHT_PER_PACK, "0.00"))
            Bd.TOTAL_PRICE = Val(Format(Bd.AVG_PRICE * Bd.ITEM_AMOUNT, "0.00")) - Bd.DISCOUNT_AMOUNT - Bd.EXTRA_DISCOUNT
            Call Bd.UpdateAvgSellPrice
         End If
         Rs.MoveNext
      Wend
      
      
      
   Call glbDaily.CommitTransaction
      
   glbErrorLog.LocalErrorMsg = "การอัพเดดเสร็จสิ้น"
   glbErrorLog.ShowUserError
   
   OKClick = True
   Set Bd = Nothing
   Exit Sub
   
ErrorHandler:
   Call glbDaily.RollbackTransaction
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboGroup, , SET_PRODUCT)
      
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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
   pnlHeader.Caption = MapText("ระบบปรับราคาย้อนหลัง")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblPrice, "ราคาเฉลี่ย(+,-)")
   Call InitNormalLabel(lblCustomer, "ลูกค้า")
   Call InitNormalLabel(lblFromDate, "จากวันที่")
   Call InitNormalLabel(lblToDate, "จากวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblDocumentNo, "หมายเลขเอกสาร")
   Call InitNormalLabel(lblGroup, "กลุ่มวัตถุดิบ")
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   
   Call InitOptionEx(radFeature, "สินค้า/บริการ")
   Call InitOptionEx(radStock, "สินค้า/วัตถุดิบ")
   Call InitOptionEx(radCustom, "กำหนดเอง")
   
   radStock.Value = True
   txtFileName.Enabled = False
   Call InitCombo(cboGroup)
   
   txtCustomer.SetKeySearch ("CUSTOMER_CODE")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Call EnableForm(Me, False)
   m_HasActivate = False
   Set ExcelColl = New Collection
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub DeleteBalance(ToDate As Date)
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
Dim WhereStr2 As String
   
   WhereStr = ""
   WhereStr2 = ""
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr = "" Then
         WhereStr = " WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM BALANCE_ACCUM " & WhereStr
   m_Conn.Execute (SQL1)
   
   WhereStr = ""
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr = "" Then
         WhereStr = " WHERE (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr2 = "" Then
         WhereStr2 = " WHERE (J.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr2 = WhereStr2 & " AND (J.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM JOB_INOUT II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB_VERIFY II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB_RESOURCE II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB_PARAMETER II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB J " & WhereStr2
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM LOT_ITEM II WHERE II.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
      
   SQL1 = "UPDATE BILLING_DOC BD SET BD.COMMIT_FLAG = 'Y',BD.INVENTORY_DOC_ID = NULL WHERE BD.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM INVENTORY_DOC IVD " & WhereStr
   m_Conn.Execute (SQL1)
   
End Sub
Private Function GetDisplayID() As Long
   If radFeature.Value Then
      GetDisplayID = 2
   ElseIf radStock.Value Then
      GetDisplayID = 3
   ElseIf radCustom.Value Then
      GetDisplayID = 1
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set ExcelColl = Nothing
End Sub

Private Sub radCustom_Click(Value As Integer)
   If radStock.Value Then
      cboGroup.Enabled = True
   Else
      cboGroup.Enabled = False
   End If
End Sub

Private Sub radFeature_Click(Value As Integer)
   If radStock.Value Then
      cboGroup.Enabled = True
   Else
      cboGroup.Enabled = False
   End If
End Sub

Private Sub radStock_Click(Value As Integer)
   If radStock.Value Then
      cboGroup.Enabled = True
   Else
      cboGroup.Enabled = False
   End If
End Sub
Private Sub GenerateExcelColl()
Dim Mn As CMenuItem
Dim j As Long
Dim TempNo As String

   If Len(txtFileName.Text) <= 0 Then
      Exit Sub
   End If
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(1)
   
   TempNo = "Begin"
   j = 2
   While Len(TempNo) > 0
      TempNo = m_ExcelSheet.Cells(j, 2).Value
      Set Mn = New CMenuItem
      Mn.KEYWORD = TempNo
      Call ExcelColl.add(Mn, TempNo)
      Set Mn = Nothing
      j = j + 1
   Wend
   m_ExcelApp.Workbooks.Close
End Sub


