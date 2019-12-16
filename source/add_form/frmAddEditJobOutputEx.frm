VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobOutputEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditJobOutputEx.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4515
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7964
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   1620
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   2100
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3000
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   2550
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdAmount 
         Height          =   435
         Left            =   5700
         TabIndex        =   5
         Top             =   1620
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1170
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   5700
         TabIndex        =   3
         Top             =   1170
         Width           =   1485
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   180
         TabIndex        =   21
         Top             =   1230
         Width           =   1545
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   1230
         Width           =   1905
      End
      Begin VB.Label lblStdAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   1680
         Width           =   1905
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   60
         TabIndex        =   18
         Top             =   2580
         Width           =   1665
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   30
         TabIndex        =   17
         Top             =   3030
         Width           =   1695
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   2130
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   14
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   270
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2370
         TabIndex        =   9
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4020
         TabIndex        =   10
         Top             =   3720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobOutputEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public id As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection5 As Collection
Public COMMIT_FLAG As String
Public typeInput As Long

Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Locations As Collection

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblProduct, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblAmount, MapText("จำนวนผลิตจริง"))
   Call InitNormalLabel(lblPlace, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblSerialNo, MapText("ซีเรียล"))
   Call InitNormalLabel(lblRef, MapText("หมายเลขอ้างอิง"))
   Call InitNormalLabel(lblStdAmount, MapText("จำนวนมาตรฐาน"))
   Call InitNormalLabel(lblWeightPerPack, MapText("น้ำหนักต่อถุง"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนถุง"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtRef.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtStdAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
           Dim Ma As CJobInput
         Set Ma = TempCollection.Item(id)

        uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, Ma.PART_TYPE_ID)
        uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
        txtWeightPerPack.Text = Ma.WEIGHT_PER_PACK
        txtPackAmount.Text = Ma.PACK_AMOUNT
        txtAmount.Text = Ma.TX_AMOUNT
        uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, Ma.LOCATION_ID)
        txtSerialNo.Text = Ma.SERIAL_NUMBER
        txtRef.Text = Ma.INOUT_REF
        txtStdAmount.Text = Ma.STD_AMOUNT
        
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      Else
      
      If typeInput = 1 Then
           Dim IWD As CInventoryWHDoc
           Dim LWH As CLotItemWH
           Dim LTD As CLotDoc
           Dim PD As CPalletDoc
            Set IWD = TempCollection5.Item(1)
            If IWD.C_LotItemsWH Is Nothing Then
               Call EnableForm(Me, True)
               Exit Sub
            End If

            If (IWD.C_LotItemsWH.Count = 0) Or (IWD.C_LotItemsWH Is Nothing) Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
   
             id = 1
            Set LWH = IWD.C_LotItemsWH.Item(id)
            
            txtWeightPerPack.Text = LWH.WEIGHT_PER_PACK
            
            If CountItem(LWH.C_LotDoc) <= 0 Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
   
            If LWH.C_LotDoc.Item(id) Is Nothing Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
            
            Dim SumPD As Double
            For Each LTD In LWH.C_LotDoc
               For Each PD In LTD.C_PalletDoc
                  SumPD = SumPD + PD.CAPACITY_AMOUNT
               Next PD
            Next LTD
            txtPackAmount.Text = SumPD * LWH.WEIGHT_PER_PACK
      End If
      
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   
   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblStdAmount, txtStdAmount, False) Then
      Exit Function
   End If
  If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Function
   End If
         
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CJobInput
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInput
   Else
      Set Ma = TempCollection.Item(id)
   End If
   
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   Ma.PART_TYPE_NAME = uctlPartTypeLookup.MyCombo.Text
   Ma.TX_AMOUNT = txtAmount.Text
   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   Ma.SERIAL_NUMBER = txtSerialNo.Text
   Ma.INOUT_REF = txtRef.Text
   Ma.TX_TYPE = "I"
   Ma.STD_AMOUNT = Val(txtStdAmount.Text)
   Ma.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
   Ma.PACK_AMOUNT = Val(txtPackAmount.Text)
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
    
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2)
      Set uctlPlaceLookup.MyCollection = m_Locations
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
         Call QueryData(True)
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Input_combo = New Collection
   Set m_Input1_combo = New Collection
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
   txtAmount.Text = Val(txtPackAmount.Text) * Val(txtWeightPerPack.Text)
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightPerPack_Change()
   m_HasModify = True
   txtAmount.Text = Val(txtPackAmount.Text) * Val(txtWeightPerPack.Text)
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlProductLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlProductLookup.MyCollection = m_PartItems
   
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlPlaceLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
   txtStdAmount.Text = txtAmount.Text
End Sub

Private Sub txtLink_Change()
   m_HasModify = True
End Sub

Private Sub txtRef_Change()
   m_HasModify = True
End Sub

Private Sub txtSerialNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlProductLookup_Change()
Dim Pi As CPartItem
Dim PartItemID As Long

   PartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_PartItems, Trim(str(PartItemID)))
      txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
   End If
   
   m_HasModify = True
End Sub
