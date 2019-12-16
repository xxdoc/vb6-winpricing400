VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobOutputEx4 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditJobOutputEx4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5955
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   10504
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlStartDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin VB.ComboBox cboLotNo 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   2085
      End
      Begin VB.ComboBox cboBinNo 
         Height          =   510
         Left            =   1800
         TabIndex        =   4
         Top             =   2100
         Width           =   2085
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   660
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   2640
         Width           =   2085
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3120
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   4080
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3600
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdAmount 
         Height          =   435
         Left            =   6045
         TabIndex        =   7
         Top             =   2640
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGoodAmount 
         Height          =   435
         Left            =   6060
         TabIndex        =   5
         Top             =   2100
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTime uctlTime1 
         Height          =   375
         Left            =   60500
         TabIndex        =   25
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   11
         Top             =   4560
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLotNoNew 
         Height          =   435
         Left            =   6060
         TabIndex        =   29
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin VB.Label lblStartDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartDate"
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblLotNo2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo2"
         Height          =   315
         Left            =   4680
         TabIndex        =   30
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   960
         TabIndex        =   28
         Top             =   1560
         Width           =   675
      End
      Begin Threed.SSCommand cmdAuto2 
         Height          =   405
         Left            =   3960
         TabIndex        =   27
         Top             =   1605
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx4.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   0
         TabIndex        =   26
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblGoodAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblGoodAmount"
         Height          =   345
         Left            =   3960
         TabIndex        =   24
         Top             =   2085
         Width           =   1935
      End
      Begin VB.Label lblBinNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBinNo"
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label lblStdAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStdAmount"
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   2640
         Width           =   1905
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   60
         TabIndex        =   21
         Top             =   3600
         Width           =   1665
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   0
         TabIndex        =   20
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   675
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2760
         TabIndex        =   12
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx4.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   13
         Top             =   5280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobOutputEx4"
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
Private m_HasModify2 As Boolean
Private m_HasModify3 As Boolean 'flag พิเศษไว้บังคับให้เปลี่ยนแปลงค่าตัวที่ต้องการ อย่าง อัตโนมัติ
Private m_Rs As ADODB.Recordset
Private m_Input_combo As Collection
Private m_Input1_combo As Collection
Public HeaderText As String
Public id As Long
Public JOB_INOUT_ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection4 As Collection
Public TempCollection2 As Collection
Public TempCollection3 As Collection
Public TempCollection5 As Collection
Public COMMIT_FLAG As String
Public StartJob As Date
Public StopJob As Date
Public PartType As Long
Private PartItemID As Long
Private m_CollLotItemWh As Collection
Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_Locations As Collection
Private m_Units As Collection
Private Lt As cLot
Public m_JobInOut As Collection
'Private tempPD As Collection
Private IWD As CInventoryWHDoc
Private LWH As CLotItemWH

Private OldLotNo As Long
Private NewLotNo As Long

Private OldHeadPackNo As Long
Private NewHeadPackNo As Long

Public typeInput As Long 'เป็นการตัดจาก bag to bulk

Private PackDate As Date

Private OldPartItemId As Long
Private NewPartItemId As Long
Public JobIdRef As Long
Public DocumentType As Long
Private LotDocId As Long
Private m_CollLotExUse As Collection

Private Sub cboBinNo_Change()
   m_HasModify = True
End Sub

Private Sub cboBinNo_Click()
   m_HasModify = True
End Sub

Private Sub cboBinNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
   KeyAscii = 0
End Sub

Private Sub cboHead_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboLockNo_Change()
   m_HasModify = True
End Sub

Private Sub cboLockNo_Click()
   m_HasModify = True
End Sub

Private Sub cboPalletNo_Change()
   m_HasModify = True
End Sub

Private Sub cboPalletNo_Click()
   m_HasModify = True
End Sub

Private Sub cboLockNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboLotNo_Change()
   m_HasModify = True
End Sub

Private Sub cboLotNo_Click()
m_HasModify = True
NewLotNo = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
End Sub

Private Sub cboLotNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
KeyAscii = 0
End Sub

Private Sub cmdAuto2_Click()
Dim No As String
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim IsOK As Boolean

  Set oMenu = New cPopupMenu
  lMenuChosen = oMenu.Popup("เพิ่ม LOT NO ใหม่", "-", "บันทึก", "-", "ลบ", "-", "LOT NO อื่นๆ")
  If lMenuChosen = 0 Then
      Exit Sub
  ElseIf lMenuChosen = 1 Then
      lblLotNo2.Enabled = True
      txtLotNoNew.Enabled = True
      txtLotNoNew.SetFocus
   ElseIf lMenuChosen = 3 Then
      If Not VerifyTextControl(lblLotNo2, txtLotNoNew, False) Then
        Exit Sub
      End If
      
      If Not VerifyDate(lblStartDate, uctlStartDate, False) Then
        Exit Sub
      End If
      
      Set Lt = New cLot
      Lt.AddEditMode = SHOW_ADD
      No = "LG" & Right(Format(Year(uctlStartDate.ShowDate) + 543, "0000"), 2) & Format(uctlStartDate.ShowDate, "mm") & Format(uctlStartDate.ShowDate, "dd")
      Lt.LOT_NO = No & Format(Val(txtLotNoNew.Text), "000")
      Lt.LOT_DATE = uctlStartDate.ShowDate

      If Not CheckUniqueNs(LOT_UNIQUE, Lt.LOT_NO, id) Then
          glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & Lt.LOT_NO & " " & MapText("อยู่ในระบบแล้ว")
          glbErrorLog.ShowUserError
          Call LoadLotFromLot(cboLotNo, Nothing, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , , 1, , 2, , Lt.LOT_NO)
          Call EnableForm(Me, True)
          Exit Sub
       End If

      Call Lt.AddEditData
      Call LoadLotIdByPartItem(cboLotNo, m_CollLotItemWh, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , PartItemID, 5, 1, 1, "I", TempCollection3, 1, Lt)
      lblLotNo2.Enabled = False
      txtLotNoNew.Enabled = False
   ElseIf lMenuChosen = 5 Then
      If Not VerifyCombo(lblLotNo, cboLotNo, False) Then
         Exit Sub
      End If

      Call EnableForm(Me, False)
      If Not glbDaily.DeleteLot(cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex)), IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call LoadLotFromLot(cboLotNo, Nothing, , uctlStartDate.ShowDate, uctlStartDate.ShowDate, , , 1, , 2)
      Call EnableForm(Me, True)

    ElseIf lMenuChosen = 7 Then
      Call LoadLotFromLot(cboLotNo, Nothing, , , , , , 1, , 2)
   End If
End Sub



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
   cmdAuto2.Picture = LoadPicture(glbParameterObj.NormalButton1)
     
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText

   Call InitNormalLabel(lblStartDate, MapText("วันที่ผลิต"))
   Call InitNormalLabel(lblType, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblProduct, MapText("เบอร์สินค้า"))
   Call InitNormalLabel(lblLotNo, MapText("Lot การผลิต"))
   Call InitNormalLabel(lblLotNo2, MapText("Lot"))
   Call InitNormalLabel(lblGoodAmount, MapText("จำนวนในโกดัง"))
   Call InitNormalLabel(lblAmount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblStdAmount, MapText("จำนวนมาตรฐาน"))
   Call InitNormalLabel(lblBinNo, MapText("เบอร์ถัง"))
   Call InitNormalLabel(lblPlace, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblSerialNo, MapText("ซีเรียล"))
   Call InitNormalLabel(lblRef, MapText("หมายเลขอ้างอิง"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))

   Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
   lblLotNo2.Enabled = False
   txtLotNoNew.Enabled = False
   
   Call txtGoodAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtStdAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call uctlProductLookup.MyTextBox.SetKeySearch("PART_NO")
   
   If DocumentType = 17 Or DocumentType = 18 Then
      txtAmount.Enabled = True
      txtStdAmount.Enabled = True
   Else
      txtAmount.Enabled = False
      txtStdAmount.Enabled = False
   End If
   
   If JobIdRef > 0 Then
      cboLotNo.Enabled = False
      cmdAuto2.Enabled = False
      cboBinNo.Enabled = False
      txtGoodAmount.Enabled = False
   End If
   
If ShowMode = SHOW_ADD Then
  uctlPartTypeLookup.Enabled = False
  uctlPlaceLookup.Enabled = False
End If

   Call InitCombo(cboLotNo)
   Call InitCombo(cboBinNo)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdAuto2, MapText("A"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
      
         If TempCollection Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      
          Dim Ma As CJobInput
         Set Ma = TempCollection.Item(id)
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, Ma.PART_TYPE_ID)
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Ma.PART_ITEM_ID)
         txtAmount.Text = Ma.TX_AMOUNT
         txtStdAmount.Text = Ma.STD_AMOUNT
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, Ma.LOCATION_ID)
         txtSerialNo.Text = Ma.SERIAL_NUMBER
         txtRef.Text = Ma.INOUT_REF
         
          If TempCollection2 Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      
         Set IWD = TempCollection2.Item(id)
         If IWD.C_LotItemsWH Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If (IWD.C_LotItemsWH.Count = 0) Or (IWD.C_LotItemsWH Is Nothing) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
          If id = 2 Then
            id = 1
          End If
         Set LWH = IWD.C_LotItemsWH.Item(id)
         
         If CountItem(LWH.C_LotDoc) <= 0 Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If LWH.C_LotDoc.Item(id) Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         Dim LTD As CLotDoc
         Set LTD = LWH.C_LotDoc.Item(id)
         
'         uctlStartDate.ShowDate = LWH.BL_START_DATE
         If LWH.BL_START_DATE > 0 Then
            uctlStartDate.ShowDate = LWH.BL_START_DATE
         Else
            uctlStartDate.ShowDate = LWH.START_DATE
            m_HasModify3 = True
         End If
         cboLotNo.ListIndex = IDToListIndex(cboLotNo, LTD.LOT_ID)
         'ทำไว้เพื่อเช็คว่า มีการเปลี่ยน lot ใหม่หรือไม่ เพื่อจะได้ให้โปรแกรมตัดสินใจได้ว่า จะสร้าง pallet ใหม่ หรือไม่
         OldLotNo = LTD.LOT_ID
         NewLotNo = LTD.LOT_ID
         
         LotDocId = LTD.LOT_DOC_ID
         
         OldHeadPackNo = LWH.HEAD_PACK_NO
         NewHeadPackNo = LWH.HEAD_PACK_NO
         
         OldPartItemId = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
         NewPartItemId = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
         ''''''''''''''''''''''''''''''
         
         cboBinNo.ListIndex = IDToListIndex(cboBinNo, LWH.BIN_NO)
         txtGoodAmount.Text = LWH.GOOD_AMOUNT
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, LWH.LOCATION_ID)
         txtNote.Text = LWH.NOTE
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         
            'ตรวจสอบว่า lot นี้ได้มีการปรับยอดไปแล้วหรือไม่
            If LTD.BALANCE_FLAG = "Y" Then
              ' MsgBox "Lot  นี้มีการปรับยอดแล้ว ไม่สามารถเปลี่ยนแปลงแก้ไขข้อมูลได้ "
               m_HasModify2 = True
               Call EnableForm(Me, True)
               Exit Sub
            End If
      Else
         If Not TempCollection5 Is Nothing Then
          If typeInput = 1 Then
       
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
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, LWH.PART_ITEM_ID)
         If CountItem(LWH.C_LotDoc) <= 0 Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If LWH.C_LotDoc.Item(id) Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
       
         
         Set LTD = LWH.C_LotDoc.Item(id)
'         Set TempCLotDoc = LTD
         
         If LWH.BL_START_DATE > 0 Then
            uctlStartDate.ShowDate = LWH.BL_START_DATE
         Else
            uctlStartDate.ShowDate = LWH.START_DATE
         End If
         
         If typeInput = 1 Then
            uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 110)
         End If
         
         cboLotNo.ListIndex = IDToListIndex(cboLotNo, LTD.LOT_ID)
         cboBinNo.ListIndex = IDToListIndex(cboBinNo, LWH.BIN_NO)

         Dim R As Long
         Dim IsOne As Boolean
         Dim SumPD As Double
         Dim PD As CPalletDoc
         For Each PD In LTD.C_PalletDoc
            SumPD = SumPD + PD.CAPACITY_AMOUNT
         Next PD

         txtGoodAmount.Text = SumPD * LWH.WEIGHT_PER_PACK
      Else
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 109)
       End If
      Else
         lblPlace.Enabled = True
         uctlPlaceLookup.Enabled = True
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
Dim I As Long
Dim tempPallet As Collection

   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not JobIdRef > 0 Then
      If Not VerifyCombo(lblLotNo, cboLotNo, False) Then
         Exit Function
      End If
      
      If Not VerifyCombo(lblBinNo, cboBinNo, False) Then
         Exit Function
      End If
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
   
   If m_HasModify2 Then
      MsgBox "Lot  นี้มีการปรับยอดแล้ว ไม่สามารถเปลี่ยนแปลงแก้ไขข้อมูลได้ "
      Exit Function
   End If
   
   If Val(txtAmount.Text) = 0 Then
        Call MsgBox(lblAmount.Caption & "ต้องไม่เท่ากับ 0 ", vbOKOnly, PROJECT_NAME)
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
   
   
   Ma.PART_TYPE_ID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   Ma.PART_DESC = uctlProductLookup.MyCombo.Text
   Ma.PART_NO = uctlProductLookup.MyTextBox.Text
   Ma.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Ma.PART_TYPE_NAME = uctlPartTypeLookup.MyCombo.Text
   Ma.TX_AMOUNT = Val(txtAmount.Text)
   Ma.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   Ma.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   Ma.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   Ma.SERIAL_NUMBER = txtSerialNo.Text
   Ma.INOUT_REF = txtRef.Text
   Ma.TX_TYPE = "I"
   Ma.STD_AMOUNT = Val(txtStdAmount.Text)
  
   
   If ShowMode = SHOW_ADD Then
      Ma.Flag = "A"
      Call TempCollection.add(Ma)
   Else
      If Ma.Flag <> "A" Then
         Ma.Flag = "E"
      End If
   End If
 If Not JobIdRef > 0 Then
   If ShowMode = SHOW_ADD Then
      Set IWD = New CInventoryWHDoc
      Set LWH = New CLotItemWH
   Else
      Set IWD = TempCollection2.Item(id)
      Set LWH = IWD.C_LotItemsWH.Item(id)
   End If
   
   LWH.BL_START_DATE = uctlStartDate.ShowDate
   LWH.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   LWH.PART_NO = uctlProductLookup.MyTextBox.Text
   LWH.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
   LWH.LOT_NO = cboLotNo.Text
   LWH.GOOD_AMOUNT = Val(txtGoodAmount.Text)
   LWH.TX_AMOUNT = Val(txtAmount.Text)
   LWH.NOTE = txtNote.Text
   LWH.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   LWH.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   LWH.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   LWH.TX_TYPE = "I" 'รับเข้า
   LWH.PACK_DATE = Now
   
   Dim LTD As CLotDoc
   Dim PD As CPalletDoc
   Set LTD = New CLotDoc
   If ShowMode = SHOW_ADD Then
         Set PD = New CPalletDoc
         PD.Flag = "A"
         PD.PALLET_DOC_NO = "1001"
         PD.CAPACITY_AMOUNT = Val(txtAmount.Text)
         PD.TX_TYPE = "I"
         PD.AddEditMode = ShowMode
         Call LTD.C_PalletDoc.add(PD)
         Set PD = Nothing
      
         LTD.Flag = "A"
         LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
         LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
         LTD.AddEditMode = ShowMode
         Call LWH.C_LotDoc.add(LTD)
         LWH.Flag = "A"
         Call IWD.C_LotItemsWH.add(LWH)
         IWD.Flag = "A"
         Call TempCollection2.add(IWD)
   Else
      LWH.Flag = "E"
      Set LTD = LWH.C_LotDoc.Item(1)
      LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
      LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
      LTD.AddEditMode = ShowMode
      LTD.Flag = "E"
      
      Set PD = LTD.C_PalletDoc(1)
         PD.CAPACITY_AMOUNT = Val(txtAmount.Text)
         PD.Flag = "E"
         PD.AddEditMode = ShowMode
      Set PD = Nothing
      Set LTD = Nothing
   End If
   End If
   SaveData = True
End Function
Function ImportInput(PartItemID As Long, TX_AMOUNT As Double)
Dim TempJob As CJob
Dim TempJobIn As CJobInput
Dim Ma As CJobInput
      'Input ส่วนผสมที่ใช้
      Set TempJob = GetObject("Cjob", m_JobInOut, Trim(str(PartItemID)))
      If Not TempJob Is Nothing Then
      For Each TempJobIn In TempJob.Inputs
        Set Ma = New CJobInput
       Ma.PART_NO = TempJobIn.PART_NO
       Ma.PART_ITEM_ID = TempJobIn.PART_ITEM_ID
       Ma.PART_TYPE_ID = TempJobIn.PART_TYPE_ID
       Ma.PART_TYPE_NAME = TempJobIn.PART_TYPE_NAME
       If Ma.PART_TYPE_ID = 26 Or Ma.PART_TYPE_ID = 29 Or Ma.PART_TYPE_ID = 30 Or Ma.PART_TYPE_ID = 31 Or Ma.PART_TYPE_ID = 47 Or Ma.PART_TYPE_ID = 48 Then
         Ma.TX_AMOUNT = (TX_AMOUNT * 2) / 100
         Ma.PART_TYPE_ID = 22
         Ma.LOCATION_ID = 117
         Ma.LOCATION_NO = ".PACK"
      Else
         Ma.TX_AMOUNT = (TX_AMOUNT * 95) / 100
         Ma.PART_TYPE_ID = 22
         Ma.LOCATION_ID = 110
         Ma.LOCATION_NO = ".BK"
       End If
       Ma.TX_TYPE = "E" 'TempJobIn.TX_TYPE
       Ma.Flag = "A"
       Call TempCollection4.add(Ma)
      Next TempJobIn
      End If
End Function

Function SumPalletAmount(Cl As Collection) As Double
   Dim PD As CPalletDoc
   For Each PD In Cl
      If Not PD.Flag = "D" Then
         SumPalletAmount = SumPalletAmount + PD.CAPACITY_AMOUNT
      End If
   Next PD
End Function
Function CheckMaxMinNamePallet(Cl As Collection)
   Dim PD As CPalletDoc
   Dim MIN As Long
   Dim MAX As Long
   Dim CMin As Long
   Dim CMax As Long
   Dim C As Long
  'FINE MIN
  MIN = 1
  MAX = 1
  C = 0
   For Each PD In Cl
      If Not PD.Flag = "D" Then
         C = C + 1
         If MIN <= Val(PD.PALLET_DOC_NO) And C = 1 Then
            MIN = Val(PD.PALLET_DOC_NO)
            CMin = PD.CAPACITY_AMOUNT
         ElseIf MIN > Val(PD.PALLET_DOC_NO) Then
            MIN = Val(PD.PALLET_DOC_NO)
            CMin = PD.CAPACITY_AMOUNT
         End If
         
         If MAX >= Val(PD.PALLET_DOC_NO) And C = 1 Then
            MAX = Val(PD.PALLET_DOC_NO)
            CMax = PD.CAPACITY_AMOUNT
         ElseIf MAX < Val(PD.PALLET_DOC_NO) Then
            MAX = Val(PD.PALLET_DOC_NO)
            CMax = PD.CAPACITY_AMOUNT
         End If
      End If
   Next PD
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
      
     Call LoadLotFromLot(cboLotNo, Nothing, , , , , , 1, , 2)
     Call LoadLocation(cboBinNo, Nothing, 2, , , , 2, "BIN")
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         Call LoadLotRefExByLotDocId(Nothing, m_CollLotExUse, -1, -1, LotDocId)
      ElseIf ShowMode = SHOW_ADD Then
         id = 0
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
         uctlStartDate.ShowDate = StartJob
         Call QueryData(True)
      End If
      
   If Not m_HasModify3 Then
      m_HasModify = False
   End If
'      m_HasModify = False
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
   Set m_Units = New Collection
   Set m_CollLotItemWh = New Collection
   Set TempCollection3 = New Collection
   Set m_JobInOut = New Collection
   Set m_CollLotExUse = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_Units = Nothing
   Set m_CollLotItemWh = Nothing
   Set TempCollection3 = Nothing
   Set Lt = Nothing
   Set m_JobInOut = Nothing
   Set m_CollLotExUse = Nothing
End Sub

Private Sub txtGoodAmount_Change()
On Error Resume Next
   m_HasModify = True
   If DocumentType = 17 Or DocumentType = 18 Or DocumentType = 13 Then
      txtAmount.Text = Val(txtGoodAmount.Text)
      txtStdAmount.Text = Val(txtGoodAmount.Text)
   End If
End Sub

Private Sub txtLoseAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLotNo_Change()
   m_HasModify = True
End Sub

Private Sub txtLotNoNew_Change()
   m_HasModify = True
End Sub

Private Sub txtLotNoNew_KeyPress(KeyAscii As Integer)
  KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPackAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletFrom_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletFrom2_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletPerUnit_Change()
   m_HasModify = True
'       txtGoodAmount.Text = (Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)
End Sub

Private Sub txtPalletPerUnit2_Change()
   m_HasModify = True
'   txtGoodAmount.Text = ((Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)) + Val(txtPalletPerUnit2.Text)
End Sub

Private Sub txtRestAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtStdAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtTimePackBegin_HasChange()
   m_HasModify = True
End Sub

Private Sub txtTimePackEnd_HasChange()
   m_HasModify = True
End Sub

Private Sub txtWeightPerPack_Change()
   m_HasModify = True
End Sub
Function getFormat(ProductType As Long, WEIGHT As Long) As Long
Dim data As Long
   If ProductType = 221 And WEIGHT = 30 Then 'ผง
      data = 48
   ElseIf ProductType = 221 And WEIGHT = 50 Then 'ผง
      data = 30
   ElseIf ProductType = 222 And WEIGHT = 10 Then 'เม็ด
      data = 60
   ElseIf ProductType = 222 And WEIGHT = 20 Then 'เม็ด
      data = 60
   ElseIf ProductType = 222 And WEIGHT = 30 Then 'เม็ด
      data = 60
   ElseIf ProductType = 222 And WEIGHT = 50 Then 'เม็ด
      data = 35
   ElseIf ProductType = 227 And WEIGHT = 30 Then  'ครัม
      data = 60
   Else
      data = 0
   End If
   getFormat = data
End Function
Private Sub uctlPackDate_HasChange()
   m_HasModify = True
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
   PartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_PartItems, Trim(str(PartItemID)))
   End If
    ''Call LoadLotIdByPartItem(cboLotNo, m_CollLotItemWh, , , , , PartItemID, 5, 1, 1, "I", TempCollection3, 1, Lt)
   m_HasModify = True

   NewPartItemId = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
End Sub

Private Sub uctlProductTypeLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlStartDate_HasChange()
   m_HasModify = True
End Sub
