VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPigDoc1 
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmAddEditPigDoc1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5700
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   10054
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlHouseLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1950
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1530
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2655
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtParentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   2400
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMotherNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2850
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtBirthAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   3300
         Width           =   1455
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalWeight 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   3750
         Width           =   1455
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotal 
         Height          =   435
         Left            =   5880
         TabIndex        =   8
         Top             =   3750
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlResponseByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   4200
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   4590
         TabIndex        =   1
         Top             =   1080
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblResponseBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   25
         Top             =   4260
         Width           =   1455
      End
      Begin VB.Label lblHouse 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   7320
         TabIndex        =   23
         Top             =   3840
         Width           =   585
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   22
         Top             =   3810
         Width           =   1125
      End
      Begin VB.Label lblTotalWeight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   21
         Top             =   3870
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   3390
         TabIndex        =   20
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3390
         TabIndex        =   19
         Top             =   3390
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   18
         Top             =   1560
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2865
         TabIndex        =   10
         Top             =   4860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigDoc1.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4515
         TabIndex        =   11
         Top             =   4860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblBirthAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   16
         Top             =   3420
         Width           =   1695
      End
      Begin VB.Label lblMotherNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   2940
         Width           =   1575
      End
      Begin VB.Label lblParentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   2490
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   1140
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditPigDoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryDoc As CInventoryDoc
Private m_Houses As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String
Private m_SumUnit As Double
Dim m_OldPartItemID As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_InventoryDoc.INVENTORY_DOC_ID = ID
      m_InventoryDoc.COMMIT_FLAG = ""
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_InventoryDoc.PopulateFromRS(1, m_Rs)
      uctlDocumentDate.ShowDate = m_InventoryDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_InventoryDoc.DOCUMENT_NO
      uctlResponseByLookup.MyCombo.ListIndex = IDToListIndex(uctlResponseByLookup.MyCombo, m_InventoryDoc.EMP_ID)
      
      chkCommit.Value = FlagToCheck(m_InventoryDoc.COMMIT_FLAG)
      chkCommit.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      txtBirthAmount.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      uctlHouseLookup.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      
      Dim Ii As CLotItem
      If m_InventoryDoc.ImportExports.Count > 0 Then
         Set Ii = m_InventoryDoc.ImportExports(1)
         uctlHouseLookup.MyCombo.ListIndex = IDToListIndex(uctlHouseLookup.MyCombo, Ii.LOCATION_ID)
         txtParentNo.Text = Ii.FATHER_NO
         txtMotherNo.Text = Ii.MOTHER_NO
         txtTotalWeight.Text = Ii.TOTAL_WEIGHT
         txtBirthAmount.Text = Ii.TX_AMOUNT
         
         m_OldPartItemID = Ii.PART_ITEM_ID
      End If
      
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Pi As CPartItem
   
   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblHouse, uctlHouseLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBirthAmount, txtBirthAmount, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalWeight, txtTotalWeight, True) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtDocumentNo.Text & " " & MapText("������к�����")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Set Pi = glbDaily.DateToPartItem(uctlDocumentDate.ShowDate)
   If Pi Is Nothing Then
      glbErrorLog.LocalErrorMsg = MapText("��辺�������ء�㹪�ǧ�ѹ�Դ����к�")
      glbErrorLog.ShowUserError
      
      Exit Function
   End If
   
   If Pi.PART_ITEM_ID <= 0 Then
      glbErrorLog.LocalErrorMsg = MapText("��辺�������ء�㹪�ǧ�ѹ�Դ����к�")
      glbErrorLog.ShowUserError
      
      Exit Function
   End If
      
   If m_InventoryDoc.COMMIT_FLAG = "Y" Then
      If m_InventoryDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(m_InventoryDoc.ImportExports)
      
         If m_OldPartItemID > 0 Then
            If m_OldPartItemID <> Pi.PART_ITEM_ID Then
               glbErrorLog.LocalErrorMsg = MapText("�ѹ�������䢨е�ͧ������㹪�ǧ�ѻ�����Դ���ǡѹ")
               glbErrorLog.ShowUserError
               
               Exit Function
            End If
         End If
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryDoc.AddEditMode = ShowMode
   m_InventoryDoc.INVENTORY_DOC_ID = ID
    m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_InventoryDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryDoc.EMP_ID = uctlResponseByLookup.MyCombo.ItemData(Minus2Zero(uctlResponseByLookup.MyCombo.ListIndex))
   m_InventoryDoc.DOCUMENT_TYPE = 5
   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   
   Dim Ii As CLotItem
   If m_InventoryDoc.ImportExports.Count <= 0 Then
      Set Ii = New CLotItem
      Ii.Flag = "A"
      Call m_InventoryDoc.ImportExports.Add(Ii)
   Else
      Set Ii = m_InventoryDoc.ImportExports(1)
      If Ii.Flag <> "A" Then
         Ii.Flag = "E"
      End If
   End If
   
   Ii.INCLUDE_UNIT_PRICE = 0
   Ii.ACTUAL_UNIT_PRICE = 0
   Ii.PART_ITEM_ID = Pi.PART_ITEM_ID
   Ii.LOCATION_ID = uctlHouseLookup.MyCombo.ItemData(Minus2Zero(uctlHouseLookup.MyCombo.ListIndex))
   Ii.FATHER_NO = txtParentNo.Text
   Ii.MOTHER_NO = txtMotherNo.Text
   Ii.TX_AMOUNT = Val(txtBirthAmount.Text)
   Ii.TOTAL_WEIGHT = Val(txtTotalWeight.Text)
   Ii.CALCULATE_FLAG = "N"
   
'   Call CalculateIncludePrice
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
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
   
   Set Pi = Nothing
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

Private Sub CalculateIncludePrice()
Dim Ii As CLotItem
Dim AvgFee As Double

   If m_SumUnit > 0 Then
      AvgFee = Val(txtBirthAmount.Text) / m_SumUnit
   Else
      AvgFee = 0
   End If
   
   For Each Ii In m_InventoryDoc.ImportExports
      If Ii.Flag <> "D" Then
         Ii.INCLUDE_UNIT_PRICE = Ii.ACTUAL_UNIT_PRICE + AvgFee
      End If
   Next Ii
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadLocation(uctlHouseLookup.MyCombo, m_Houses, 1)
      Set uctlHouseLookup.MyCollection = m_Houses
      
      Call LoadEmployee(uctlResponseByLookup.MyCombo, m_Employees)
      Set uctlResponseByLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryDoc.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_InventoryDoc.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryDoc = Nothing
   Set m_Houses = Nothing
   Set m_Employees = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GetTotalPrice()
Dim Ii As CLotItem
Dim Sum As Double

   Sum = 0
   m_SumUnit = 0
   For Each Ii In m_InventoryDoc.ImportExports
      If Ii.Flag <> "D" Then
         Sum = Sum + CDbl(Format(Ii.TOTAL_ACTUAL_PRICE, "0.00"))
         m_SumUnit = m_SumUnit + Ii.TX_AMOUNT
      End If
   Next Ii
   
   txtTotalWeight.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblMotherNo, MapText("�����Ţ���"))
   Call InitNormalLabel(lblDocumentNo, MapText("�Ţ�����Դ"))
   Call InitNormalLabel(lblParentNo, MapText("�����Ţ���"))
   Call InitNormalLabel(lblBirthAmount, MapText("�ӹǹ"))
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����Դ"))
   Call InitNormalLabel(lblTotalWeight, MapText("���˹ѡ���"))
   Call InitNormalLabel(Label1, MapText("���"))
   Call InitNormalLabel(Label2, MapText("�.�."))
   Call InitNormalLabel(Label4, MapText("�.�."))
   Call InitNormalLabel(lblTotal, MapText("���˹ѡ�����"))
   Call InitNormalLabel(lblHouse, MapText("�����ç���͹�Դ"))
   Call InitNormalLabel(lblResponseBy, MapText("����Ѻ�Դ�ͺ"))
   
   Call InitCheckBox(chkCommit, MapText("�ӹǳ"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtParentNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtMotherNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBirthAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtTotalWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtTotal.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtTotal.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
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
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Houses = New Collection
   Set m_Employees = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub txtBirthAmount_Change()
   m_HasModify = True
   If Val(txtBirthAmount.Text) > 0 Then
      txtTotal.Text = Format(Val(txtTotalWeight.Text) / Val(txtBirthAmount.Text), "0.00")
   Else
      txtTotal.Text = Format(0, "0.00")
   End If
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtParentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalWeight_Change()
   m_HasModify = True
   If Val(txtBirthAmount.Text) > 0 Then
      txtTotal.Text = Format(Val(txtTotalWeight.Text) / Val(txtBirthAmount.Text), "0.00")
   Else
      txtTotal.Text = Format(0, "0.00")
   End If
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtMotherNo_Change()
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub
