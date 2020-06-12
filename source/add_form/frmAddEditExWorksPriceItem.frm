VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExWorksPriceItem 
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditExWorksPriceItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3825
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6747
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   2040
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtSocCode 
         Height          =   435
         Left            =   2520
         TabIndex        =   0
         Top             =   1140
         Width           =   4005
         _ExtentX        =   11615
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   3120
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdPrev 
            Height          =   525
            Left            =   2520
            TabIndex        =   8
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExWorksPriceItem.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   525
            Left            =   4200
            TabIndex        =   9
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExWorksPriceItem.frx":2ABC
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8520
            TabIndex        =   6
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmAddEditExWorksPriceItem.frx":2DD6
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10140
            TabIndex        =   7
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   10
         TabIndex        =   11
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPackageRate 
         Height          =   435
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   1695
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   405
         Left            =   2520
         TabIndex        =   1
         Top             =   1590
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtGP 
         Height          =   435
         Left            =   6960
         TabIndex        =   4
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
      End
      Begin VB.Label lblGP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         TabIndex        =   24
         Top             =   2520
         Width           =   1575
      End
      Begin Threed.SSCheck chkDeclareNew 
         Height          =   345
         Left            =   8520
         TabIndex        =   5
         Top             =   2640
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "chkEditPrice"
         TripleState     =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6240
         TabIndex        =   23
         Top             =   3060
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   4350
         TabIndex        =   22
         Top             =   3030
         Width           =   1065
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   10110
         TabIndex        =   21
         Top             =   3030
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   20
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblPackageRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblBath5 
         Height          =   315
         Left            =   4350
         TabIndex        =   18
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label lblBath4 
         Height          =   315
         Left            =   10050
         TabIndex        =   17
         Top             =   3900
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6630
         TabIndex        =   16
         Top             =   300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblBath1 
         Height          =   315
         Left            =   4350
         TabIndex        =   15
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   630
         TabIndex        =   14
         Top             =   2130
         Width           =   1755
      End
      Begin VB.Label lblSocCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   510
         TabIndex        =   13
         Top             =   1230
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmAddEditExWorksPriceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_TempCol As Collection
Public EX_WORKS_PRICE_ID As Long
Private RATE_TYPE As Long
Private RATE_AMOUNT As Double
Private PART_ITEM_ID As Long
Public SocPartType As Long
Public SocCode As String
Public m_FeatureTypes As Collection
Private m_TempFeatures As Collection
Public PartType As Long
Public ProductType As Long
Public ParentForm As Form

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ID_MUM As Long
Private CurrentKey As String
Public SocID As Long
Public TempCollection As Collection
Public m_ExWorkPricesItem As Collection
Public canShowGP As Boolean
Private DeclareCount As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim D As CExWorksPriceItem

   If Flag Then
      Call EnableForm(Me, False)

      Set D = TempCollection.Item(ID)
     If SocPartType = 3 Then
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, D.PART_ITEM_ID)
         CurrentKey = Trim(str(D.PART_ITEM_ID))
      End If
     txtPackageRate.Text = Val(D.PACKAGE_RATE)
     chkDeclareNew.Value = FlagToCheck(D.DECLARE_NEW_FLAG)
     txtGP.Text = (D.GP_VALUE)
     
      
      Call EnableForm(Me, True)
   End If

   Call EnableForm(Me, True)
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim EWP As CExWorksPriceItem

   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   
  If Not VerifyTextControl(lblPackageRate, txtPackageRate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   If Val(txtPackageRate.Text) <= 0 Then
      glbErrorLog.LocalErrorMsg = "ราคา " & lblPackageRate.Caption & " ต้องมีค่ามากกว่า 0"
      glbErrorLog.ShowUserError
      SaveData = True
      Exit Function
   End If
   
   Dim TempEWP As CExWorksPrice
   Dim TempEWP2 As CExWorksPrice
   Dim PartID As Long
   Dim PartNo As String
   Dim Key As String
     PartID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
     Key = Trim(str(PartID))
     Set TempEWP = GetObject("CExWorksPrice", m_ExWorkPricesItem, Key, False)
      If Not TempEWP Is Nothing Then
         If (TempEWP.EX_WORKS_PRICE_ITEM_ID <> ID_MUM) Or CurrentKey <> Key Then
            glbErrorLog.LocalErrorMsg = "มีข้อมูลของสินค้าเบอร์ " & uctlPartLookup.MyCombo.Text & " ในเอกสารชุดนี้แล้ว"
            glbErrorLog.ShowUserError
            Exit Function
         End If
      ElseIf ShowMode = SHOW_ADD Then
         Set TempEWP2 = New CExWorksPrice
         TempEWP2.Flag = "A"
         Call m_ExWorkPricesItem.add(TempEWP2, Trim(str(PartID)))
         Set TempEWP2 = Nothing
      End If
   
   If ShowMode = SHOW_ADD Then
      Set EWP = New CExWorksPriceItem
      EWP.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
      EWP.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
      EWP.PART_NO = uctlPartLookup.MyTextBox.Text
      EWP.PART_DESC = uctlPartLookup.MyCombo.Text
      EWP.PACKAGE_RATE = Val(txtPackageRate.Text)
      EWP.LAST_EDIT_FLAG = "Y" 'ถ้าเป็นการเพิ่มใหม่บังคับให้ Flag แก้ไขราคาเปิดใช้อัตโนมัติ
      EWP.DECLARE_NEW_FLAG = "Y"
      EWP.GP_VALUE = Val(txtGP.Text)
      EWP.RATE_TYPE = 1
      EWP.Flag = "A"
      
      Call TempCollection.add(EWP)
   Else
         Set EWP = TempCollection(ID)
         If Check2Flag(chkDeclareNew.Value) = "Y" Then 'เข้าแก้ไขได้ต่อเมื่อ ยังไม่เคยประกาศราคามาก่อนเท่านั้น
            EWP.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
            EWP.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
            EWP.PART_NO = uctlPartLookup.MyTextBox.Text
            EWP.PART_DESC = uctlPartLookup.MyCombo.Text
            EWP.PACKAGE_RATE = Val(txtPackageRate.Text)
            EWP.VERIFY_FLAG = "N"
            EWP.VERIFY_NAME = ""
            EWP.APPROVED_FLAG = "N"
            EWP.APPROVED_NAME = ""
            EWP.LAST_EDIT_FLAG = "Y" 'Check2Flag(chkDeclareNew.Value)
            
            EWP.RATE_TYPE = 1
            EWP.DECLARE_NEW_FLAG = Check2Flag(chkDeclareNew.Value)
         End If
         
         EWP.GP_VALUE = Val(txtGP.Text)
         If EWP.Flag <> "A" Then
            EWP.Flag = "E"
         End If
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub chkDeclareNew_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdNext_Click()
Dim D As CExWorksPriceItem
Dim Pt As CPartItem
 If Not SaveData Then
      Exit Sub
   End If
If ShowMode = SHOW_EDIT Then
   ID = GetNextID(ID, TempCollection)
   Set D = TempCollection(ID)
     If SocPartType = 3 Then
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, D.PART_ITEM_ID)
         
          ID_MUM = D.EX_WORKS_PRICE_ITEM_ID
          CurrentKey = Trim(str(D.PART_ITEM_ID))
         chkDeclareNew.Value = FlagToCheck(D.DECLARE_NEW_FLAG)
      End If
     txtPackageRate.Text = Val(D.PACKAGE_RATE)
     txtGP.Text = (D.GP_VALUE)
Else
  ID = GetNextID(ID, uctlPartLookup.MyCollection)
  Set Pt = uctlPartLookup.MyCollection(ID)
  uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Pt.PART_ITEM_ID)
   txtPackageRate.Text = ""
   txtPackageRate.SetFocus
   Call ParentForm.ShowGridItem
End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub cmdPrev_Click()
Dim D As CExWorksPriceItem
Dim Pt As CPartItem
 If Not SaveData Then
      Exit Sub
   End If
If ShowMode = SHOW_EDIT Then
   ID = GetPrevID(ID, TempCollection)
   Set D = TempCollection(ID)
     If SocPartType = 3 Then
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, D.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, D.PART_ITEM_ID)
         
         ID_MUM = D.EX_WORKS_PRICE_ITEM_ID
         CurrentKey = Trim(str(D.PART_ITEM_ID))
         chkDeclareNew.Value = FlagToCheck(D.DECLARE_NEW_FLAG)
      End If
     txtPackageRate.Text = Val(D.PACKAGE_RATE)
     txtGP.Text = (D.GP_VALUE)
Else
  ID = GetPrevID(ID, uctlPartLookup.MyCollection)
  Set Pt = uctlPartLookup.MyCollection(ID)
  uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Pt.PART_ITEM_ID)
   txtPackageRate.Text = ""
   txtPackageRate.SetFocus
   Call ParentForm.ShowGridItem
End If

End Sub

'
Private Sub Form_Activate()
Dim Sp As CSystemParam
Dim FeatureTypeID As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      If SocPartType = 3 Then
         Call LoadPartType(uctlPartTypeLookup.MyCombo, m_FeatureTypes)
         Set uctlPartTypeLookup.MyCollection = m_FeatureTypes
      End If

     
     
      txtSocCode.Text = SocCode
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      Else
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
      End If
      
      m_HasModify = False
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_TempCol = Nothing
   Set m_FeatureTypes = Nothing
   Set m_TempFeatures = Nothing
End Sub
'
'Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
'   ''Debug.Print ColIndex & " " & NewColWidth
'End Sub
'
Private Sub InitFormLayout()
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   Me.KeyPreview = True

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Call InitHeaderFooter(pnlHeader, pnlFooter)

   Call txtSocCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call InitNormalLabel(lblSocCode, MapText("แพคเกจ"))
   txtSocCode.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrev.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitNormalLabel(lblPartType, MapText("ประเภทสินค้า"))
   Call InitNormalLabel(lblPart, MapText("รหัสสินค้า"))
   Call txtPackageRate.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtGP.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
     If ProductType = 1 Then
         Call InitNormalLabel(lblPackageRate, MapText("ค่าสินค้า/ถุง"))
     ElseIf ProductType = 2 Then
         Call InitNormalLabel(lblPackageRate, MapText("ค่าสินค้า/กก."))
     Else
         Call InitNormalLabel(lblPackageRate, MapText("ค่าสินค้า/หน่วย"))
     End If
   Call InitNormalLabel(lblBath1, MapText("บาท"))
   Call InitNormalLabel(lblGP, MapText("% GP"))
   
   chkDeclareNew.Visible = False
   If ShowMode = SHOW_EDIT Then
      Call InitCheckBox(chkDeclareNew, "ประกาศราคาใหม่")
      chkDeclareNew.Visible = True
   End If
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdPrev, MapText("ก่อนหน้า"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   If Not canShowGP Then
      lblGP.Visible = False
      txtGP.Visible = False
   End If

End Sub
'
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub
'
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_TempCol = New Collection
   Set m_FeatureTypes = New Collection
   Set m_TempFeatures = New Collection
End Sub

Private Sub txtGP_Change()
   m_HasModify = True
End Sub

Private Sub txtPackageRate_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub
'
Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long

   m_HasModify = True

   If SocPartType = 3 Then
      PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
      If PartTypeID > 0 Then
         Call LoadPartItem(uctlPartLookup.MyCombo, m_TempFeatures, PartTypeID)
         Set uctlPartLookup.MyCollection = m_TempFeatures
      End If
   End If
End Sub
Private Function GetPrevID(OldID As Long, TempColl As Collection) As Long
Dim TempIndex As Long
Dim J As Long
   
   If OldID <= 1 Then
      J = 1
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดแรกแล้ว"
      glbErrorLog.ShowUserError
   Else
      J = OldID - 1
   End If
   GetPrevID = J
End Function
Private Function GetNextID(OldID As Long, TempColl As Collection) As Long
Dim TempIndex As Long
Dim J As Long
   
   If OldID >= TempColl.Count Then
      J = TempColl.Count
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
      glbErrorLog.ShowUserError
   Else
      J = OldID + 1
   End If
   
   GetNextID = J
End Function
