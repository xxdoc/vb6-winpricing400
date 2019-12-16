VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditAmount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmAddEditAmount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4683
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   2160
         TabIndex        =   1
         Top             =   1320
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   2160
         TabIndex        =   0
         Top             =   840
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   767
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblCurrent3 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrent3"
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label lblCurrent1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrent1"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1875
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3840
         TabIndex        =   3
         Top             =   1920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2160
         TabIndex        =   2
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAmount.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean


Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public tax1 As Double
Public tax2 As Double
Public tax3 As Double

Public TempCollection As Collection
Private Ma As CDoItem



Private Sub cmdOK_Click()
 If SaveData Then
   Call QueryData(True)
   
   OKClick = True
   Unload Me
End If
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc

   If Not VerifyTextControl(lblCurrent3, txtAmount, False) Then
      Exit Function
   End If
   
  
   Call EnableForm(Me, False)
    Set Ma = TempCollection.Item(ID)
    
     If Val(txtAmount.Text) > Ma.OLD_PACK_AMOUNT Then
      glbErrorLog.LocalErrorMsg = "จำนวนใหม่ต้องน้อยกว่าหรือเท่ากับจำนวนเก่า"
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Function
   End If
   
     Ma.PACK_AMOUNT = txtAmount.Text
     If Ma.PART_ITEM_ID = -1 Then
      Ma.ITEM_AMOUNT = Ma.PACK_AMOUNT * Ma.WEIGHT_PER_PACK_SO
     Else
      Ma.ITEM_AMOUNT = Ma.PACK_AMOUNT * Ma.WEIGHT_PER_PACK
     End If
     Ma.TX_AMOUNT = Ma.ITEM_AMOUNT
     Ma.TOTAL_PRICE = Ma.PACK_AMOUNT * Ma.PRICE_PER_PACK
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
'If Not VerifyTextControl(lblCurrent3, txtAmount, True) Then
'      Exit Sub
'   End If

If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Set Ma = TempCollection.Item(ID)
         
         txtPartNo.Text = Ma.PART_NO
         txtAmount.Text = Ma.PACK_AMOUNT
         
      End If
   End If
     
   Call EnableForm(Me, True)
   
   End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
       
       Call QueryData(True)
      
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
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCurrent1, MapText("ชื่อสินค้า"))
   Call InitNormalLabel(lblCurrent3, MapText("จำนวน"))
   Call InitNormalLabel(Label2, MapText("ถุง"))
   
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   txtPartNo.Enabled = False
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
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

