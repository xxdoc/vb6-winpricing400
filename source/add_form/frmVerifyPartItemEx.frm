VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmVerifyPartItemEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerifyPartItemEx.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3345
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5900
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtCode 
         Height          =   435
         Left            =   1815
         TabIndex        =   0
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrevDesc 
         Height          =   465
         Left            =   1830
         TabIndex        =   2
         Top             =   1320
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtPrevCode 
         Height          =   435
         Left            =   1830
         TabIndex        =   1
         Top             =   870
         Width           =   3525
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrevQuantity 
         Height          =   435
         Left            =   1830
         TabIndex        =   3
         Top             =   1800
         Width           =   5295
         _ExtentX        =   2937
         _ExtentY        =   767
      End
      Begin VB.Label lblPrevQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   30
         TabIndex        =   11
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label lblPrevCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   60
         TabIndex        =   10
         Top             =   930
         Width           =   1695
      End
      Begin VB.Label lblPrevDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   9
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   75
         TabIndex        =   8
         Top             =   300
         Width           =   1665
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2220
         TabIndex        =   4
         Top             =   2550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmVerifyPartItemEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3870
         TabIndex        =   5
         Top             =   2550
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmVerifyPartItemEx"
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

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public COMMIT_FLAG As String
Public Area As Long

Private m_Sp As CSystemParam
Private m_Features As Collection
Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Pigs As Collection
Private m_PigStatuss As Collection
Private m_SubLotItems As Collection
Private m_ManualFlag As Boolean

Private m_SocID As Long
Public AccountID As Long
Public SubscriberID As Long
Public UsageDate As Date
Public Inputs As Collection

Public DocumentType As Long
Public ParentForm As Form

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdExit_Click()
'   If Not ConfirmExit(m_HasModify) Then
'      Exit Sub
'   End If
   
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
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCode, MapText("รหัสวัตถุดิบ"))
   
   Call InitNormalLabel(lblPrevQuantity, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblPrevCode, MapText("รหัสวัตถุดิบ"))
   Call InitNormalLabel(lblPrevDesc, MapText("ชื่อวัตถุดิบ"))
   
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtCode.SetEnableDisableKeyPress (False)
   
   Call txtPrevQuantity.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPrevQuantity.Enabled = False
   Call txtPrevCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPrevCode.Enabled = False
   Call txtPrevDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPrevDesc.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   cmdExit.Enabled = False
End Sub

Private Sub CalculatePrice()
'   txtNetTotal.Text = Format(Val(txtTotalPrice.Text) - Val(txtDiscount.Text), "0.00")
'   txtLeft.Text = Format(Val(txtNetTotal.Text) - Val(txtDeposit.Text), "0.00")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Ivd As CInventoryDoc
Dim iCount As Long
Dim Ei As CLotItem

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
         
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call LoadPartItem(Nothing, m_Parts, , "")
      Call LoadFeature(Nothing, m_Features)
            
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
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
   Set m_Sp = GetSystemParam(glbSystemParams, "BARCODE_FLAG")
   OKClick = False
   Call InitFormLayout
   
   m_ManualFlag = False
   m_HasActivate = False
   m_HasModify = False
   
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Features = New Collection
   Set m_Locations = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PartTypes = New Collection
   Set m_SubLotItems = New Collection
   Set m_Features = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Features = Nothing
   Set m_Locations = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatuss = Nothing
   Set m_PartTypes = Nothing
   Set m_SubLotItems = Nothing
   Set m_Features = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub radFeature_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub txtAvgPrice_Change()
'   m_HasModify = True
'   txtTotalPrice.Text = Val(txtAvgPrice.Text) * Val(txtQuantity.Text)
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtManual_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   m_HasModify = True
   Call CalculatePrice
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalPrice_Change()
'   m_HasModify = True
'   If Val(txtQuantity.Text) > 0 Then
'      txtAvgPrice.Text = Val(txtTotalPrice.Text) / Val(txtQuantity.Text)
'   Else
'      txtAvgPrice.Text = "0.00"
'   End If
   Call CalculatePrice
End Sub

Private Function GetFeature(Col As Collection, FtCode As String) As CFeature
Dim Ft As CFeature

   For Each Ft In Col
      If Ft.FEATURE_CODE = FtCode Then
         Set GetFeature = Ft
         Exit Function
      End If
   Next Ft
   
   Set GetFeature = Nothing
End Function

Private Function GetPartItem(Col As Collection, PiCode As String) As CPartItem
Dim Pi As CPartItem

   For Each Pi In Col
      If Pi.PART_NO = PiCode Then
         Set GetPartItem = Pi
         Exit Function
      End If
   Next Pi
   
   Set GetPartItem = Nothing
End Function

Private Sub PopulateGui(Ft As CFeature, Pi As CPartItem, Di As CDoItem)
   If Not (Ft Is Nothing) Then
      txtPrevCode.Text = Ft.FEATURE_CODE
      txtPrevDesc.Text = Ft.FEATURE_DESC
   Else
      txtPrevCode.Text = Pi.PART_NO
      txtPrevDesc.Text = Pi.PART_DESC
   End If
   txtPrevQuantity.Text = Di.ITEM_AMOUNT
End Sub

Private Function CreateConfigFlag(Ft As CFeature, Pi As CPartItem) As String
Dim Flag1 As String
Dim Flag2 As String
Dim Flag3 As String

   Flag1 = "N"
   If Not (Ft Is Nothing) Then
      Flag1 = "Y"
   End If
   
   Flag2 = "N"
   If Not (Pi Is Nothing) Then
      Flag2 = "Y"
   End If
   
   Flag3 = "N"
   CreateConfigFlag = Flag1 & Flag2 & Flag3
End Function

Private Function GetDisplayID(Ft As CFeature, Pi As CPartItem) As Long
   If Not (Ft Is Nothing) Then
      GetDisplayID = 2
   ElseIf Not (Pi Is Nothing) Then
      GetDisplayID = 3
   End If
End Function

Private Sub AddDoItem(Ft As CFeature, Pi As CPartItem, Item As CDoItem)
Dim Di As CDoItem
   
   Set Di = New CDoItem
   Di.Flag = "A"
   Call TempCollection.add(Di)
   
   If Not (Ft Is Nothing) Then
      Di.FEATURE_ID = Ft.FEATURE_ID
      Di.FEATURE_CODE = Ft.FEATURE_CODE
      Di.FEATURE_DESC = Ft.FEATURE_DESC
      Di.LOCATION_ID = -1
      Di.LOCATION_NAME = ""
      Di.PART_TYPE_NAME = ""
      Di.PART_TYPE = -1
   Else
      Di.PART_ITEM_ID = Pi.PART_ITEM_ID
      Di.PART_NO = Pi.PART_NO
      Di.PART_DESC = Pi.PART_DESC
      Di.LOCATION_ID = Item.LOCATION_ID
      Di.LOCATION_NAME = Item.LOCATION_NAME
      Di.PART_TYPE_NAME = Pi.PART_TYPE_NAME
      Di.PART_TYPE = Pi.PART_TYPE
   End If
   Di.ITEM_AMOUNT = Item.ITEM_AMOUNT
   Di.TOTAL_PRICE = Item.AC_AMOUNT + Item.UC_AMOUNT
   Di.AVG_PRICE = Item.AVG_PRICE
   Di.DEPOSIT_AMOUNT = 0
   Di.DISCOUNT_AMOUNT = 0
   Di.CONFIG_CODE = CreateConfigFlag(Ft, Pi)
   Di.ITEM_DESC = ""
   Di.FROM_PERIOD = -1
   Di.TO_PERIOD = -1
   Di.DISPLAY_ID = GetDisplayID(Ft, Pi)
   Di.LOT_MANUAL = -1
   
   Set Di = Nothing
   ParentForm.RefreshGrid
End Sub

Private Function VerifyPartItem(Pi As CPartItem) As Boolean
Dim Inp As CJobInput

   For Each Inp In Inputs
      If (Inp.Flag <> "D") And (Inp.PART_ITEM_ID = Pi.PART_ITEM_ID) Then
         VerifyPartItem = True
         Exit Function
      End If
   Next Inp
   VerifyPartItem = False
End Function

Private Sub RateFeature(Ft As CFeature, Pi As CPartItem)
Dim Ug As CJobVerify
Dim IsOK As Boolean

   Call EnableForm(Me, False)
   
   Set Ug = New CJobVerify
   Ug.Flag = "A"
   Ug.PART_ITEM_ID = Pi.PART_ITEM_ID
   Ug.PART_DESC = Pi.PART_DESC
   Ug.PART_NO = Pi.PART_NO
   Ug.NOTE = "ทดสอบหมายเหตุ"
   Call TempCollection.add(Ug)
   
   txtCode.Text = ""
   txtPrevCode.Text = Ug.PART_NO
   txtPrevDesc.Text = Ug.PART_DESC
   If VerifyPartItem(Pi) Then
      Ug.NOTE = "สำเร็จ"
      Ug.VERIFY_FLAG = "Y"
   Else
      Ug.NOTE = "ไม่พบในสูตร"
      Ug.VERIFY_FLAG = "N"
   End If
   txtPrevQuantity.Text = Ug.NOTE
   
   Set Ug = Nothing
   
   ParentForm.RefreshGrid
   Call EnableForm(Me, True)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim Pi As CPartItem
Dim PartNo As String
Dim SERIALNO As String

   If KeyAscii = 13 Then
      If Len(Trim(txtCode.Text)) <= 0 Then
         Exit Sub
      End If
      
      SERIALNO = txtCode.Text
      If InStr(1, SERIALNO, "-") > 0 Then
         PartNo = Mid(SERIALNO, 1, InStr(1, SERIALNO, "-") - 1)
      Else
         PartNo = txtCode.Text
      End If
      Set Pi = GetPartItem(m_Parts, PartNo)
      If Not (Pi Is Nothing) Then
         Call RateFeature(Nothing, Pi)
         Call UpdateSerial(SERIALNO, Pi)
         Exit Sub
      End If
      
      glbErrorLog.LocalErrorMsg = "ไม่พบรหัสวัตถุดิบที่ทำการแสกนเข้ามา"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub

Private Sub UpdateSerial(Serial As String, Pi As CPartItem)
Dim Ji As CJobInput

   For Each Ji In TempCollection2
      If (Ji.Flag <> "D") And (Ji.PART_NO = Pi.PART_NO) Then
         Ji.SERIAL_NUMBER = Serial
         If Ji.Flag <> "A" Then
            Ji.Flag = "E"
         End If
      End If
   Next Ji
End Sub

Private Sub txtPrevAvgPrice_Change()
   m_HasModify = True
End Sub
