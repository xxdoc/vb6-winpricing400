VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionSale 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmAddEditCommissionSale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10815
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   6588
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboSaleType 
         Height          =   315
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2715
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSaleAmount 
         Height          =   435
         Left            =   3060
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSaleFrom 
         Height          =   435
         Left            =   3060
         TabIndex        =   0
         Top             =   1080
         Width           =   1455
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSaleTo 
         Height          =   435
         Left            =   7380
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin VB.Label lblSaleTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4440
         TabIndex        =   12
         Top             =   1170
         Width           =   2775
      End
      Begin VB.Label lblSaleType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6120
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblSaleFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1170
         Width           =   2775
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2925
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionSale.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblSaleAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2775
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6225
         TabIndex        =   6
         Top             =   2760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4575
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionSale.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCommissionSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ParentTag As String
Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object

Public TempCollection As Collection
Private Sub cboSaleType_Click()
   m_HasModify = True
End Sub

Private Sub cboSaleType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub
Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid(ParentTag)
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      txtSaleFrom.Text = ""
      txtSaleTo.Text = ""
      txtSaleAmount.Text = ""
      'cboSaleType.ListIndex = -1
   End If
   
   Call txtSaleFrom.SetFocus
   Call ParentForm.RefreshGrid(ParentTag)
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
      
   OKClick = True
   Unload Me

End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim PaymentType As Long
   
   If Flag Then
      Call EnableForm(Me, False)
      
      Dim Cmss As CCommissionSale
      Set Cmss = TempCollection.Item(ID)
      
      txtSaleFrom.Text = Cmss.SELL_FROM
      txtSaleTo.Text = Cmss.SELL_TO
      txtSaleAmount.Text = Cmss.COMMISSION_SALE_AMOUNT
      
      cboSaleType.ListIndex = IDToListIndex(cboSaleType, Cmss.COMMISSION_SALE_TYPE)
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblSaleFrom, txtSaleFrom, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblSaleTo, txtSaleTo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblSaleAmount, txtSaleAmount, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSaleType, cboSaleType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   
   Dim Cmss As CCommissionSale
   
   If ShowMode = SHOW_ADD Then
      Set Cmss = New CCommissionSale
      Cmss.Flag = "A"
      Call TempCollection.add(Cmss)
   Else
      Set Cmss = TempCollection.Item(ID)
      If Cmss.Flag <> "A" Then
         Cmss.Flag = "E"
      End If
   End If
   
   Cmss.SELL_FROM = Val(txtSaleFrom.Text)
   Cmss.SELL_TO = Val(txtSaleTo.Text)
   Cmss.COMMISSION_SALE_AMOUNT = Val(txtSaleAmount.Text)
   
   Cmss.COMMISSION_SALE_TYPE = cboSaleType.ItemData(Minus2Zero(cboSaleType.ListIndex))
   Cmss.COMMISSION_SALE_TYPE_NAME = cboSaleType.Text
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call InitCommissionSaleType(cboSaleType)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         m_HasModify = False
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
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
   ElseIf Shift = 0 And KeyCode = 117 Then
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
   End If
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblSaleFrom, MapText("ยอดขาย(มากกว่าเท่ากับ)"))
   Call InitNormalLabel(lblSaleTo, MapText("ยอดขาย(น้อยกว่า)"))
   Call InitNormalLabel(lblSaleAmount, MapText("COMMISSION"))
   Call InitNormalLabel(lblSaleType, MapText("หน่วย"))
   
   
   Call txtSaleFrom.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtSaleTo.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtSaleAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboSaleType)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub txtBatchNO_Change()
   m_HasModify = True
End Sub


Private Sub txtSaleAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtSaleFrom_Change()
   m_HasModify = True
End Sub

Private Sub txtSaleTo_Change()
   m_HasModify = True
End Sub
