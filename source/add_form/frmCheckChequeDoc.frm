VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCheckChequeDoc 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmCheckChequeDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlChequeDate 
         Height          =   405
         Left            =   1920
         TabIndex        =   1
         Top             =   1920
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   14
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5205
         Left            =   180
         TabIndex        =   7
         Top             =   2550
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9181
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
         Column(1)       =   "frmCheckChequeDoc.frx":27A2
         Column(2)       =   "frmCheckChequeDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmCheckChequeDoc.frx":290E
         FormatStyle(2)  =   "frmCheckChequeDoc.frx":2A6A
         FormatStyle(3)  =   "frmCheckChequeDoc.frx":2B1A
         FormatStyle(4)  =   "frmCheckChequeDoc.frx":2BCE
         FormatStyle(5)  =   "frmCheckChequeDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmCheckChequeDoc.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtChequeNo 
         Height          =   435
         Left            =   1920
         TabIndex        =   0
         Top             =   960
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   1920
         TabIndex        =   2
         Top             =   1440
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   767
      End
      Begin VB.Label lblChequeDocNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1470
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5040
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5280
         TabIndex        =   15
         Top             =   960
         Width           =   915
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9480
         TabIndex        =   5
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCheckChequeDoc.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9480
         TabIndex        =   6
         Top             =   1530
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCheckChequeDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCheckChequeDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   9
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCheckChequeDoc.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCheckChequeDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_ChequeDoc As CChequeDoc
Private m_TempBillingDoc As CChequeDoc
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_IvdDocType As Long
Private m_Mr As CMasterRef

Public OKClick As Boolean
Public DocumentType As CHEQUE_DOC_TYPE
Public ReceiptType As Long
Public Area As CHEQUE_DOC_TYPE
Public HeaderText As String

Private Sub cmdPasswd_Click()

End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim TempStr As String
Dim Programowner As String

   Programowner = glbParameterObj.Programowner
   
   If Area = 1 Then
      TempStr = ""
   ElseIf Area = 2 Then
      TempStr = ""
   End If
   
   
   '         frmAddEditChequeDoc.Area = Area
'         frmAddEditChequeDoc.DocumentType = DocumentType
'         frmAddEditChequeDoc.ReceiptType = lMenuChosen
'          frmAddEditChequeDoc.HeaderText = MapText("����������㺵�Ǩ�ͺ����Ѻ / �׹ ��" & TempStr)
'         frmAddEditChequeDoc.ShowMode = SHOW_ADD
'         Load frmAddEditChequeDoc
'          frmAddEditChequeDoc.Show 1
'
'         OKClick = frmAddEditChequeDoc.OKClick
'
'         Unload frmAddEditChequeDoc
'         Set frmAddEditChequeDoc = Nothing


   frmAddEditReceiptChequeDoc.DocumentType = Area
   frmAddEditReceiptChequeDoc.HeaderText = MapText("����������" & ChequeDocType2Text(Area))
   frmAddEditReceiptChequeDoc.ShowMode = SHOW_ADD
   Load frmAddEditReceiptChequeDoc
   frmAddEditReceiptChequeDoc.Show 1
   
   OKClick = frmAddEditReceiptChequeDoc.OKClick

   Unload frmAddEditReceiptChequeDoc
   Set frmAddEditReceiptChequeDoc = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtChequeNo.Text = ""
'   txtChequeNoEx.Text = ""
   txtCustomerCode.Text = ""
   
   uctlChequeDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim PaymentID As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
'   PaymentID = GridEX1.Value(8)
    
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
    id = GridEX1.Value(1)
   
   Call EnableForm(Me, False)
     If Not glbDaily.DeleteChequeDoc(id, IsOK, True, glbErrorLog) Then
      m_ChequeDoc.CHEQUE_DOC_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)

End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
Dim TempStr As String
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   id = Val(GridEX1.Value(1))
   
   If Area = 1 Then
      TempStr = ""
   ElseIf Area = 2 Then
      TempStr = ""
   End If
   
   frmAddEditReceiptChequeDoc.id = id
   frmAddEditReceiptChequeDoc.Area = Area
   frmAddEditReceiptChequeDoc.HeaderText = MapText("��䢢�����" & ChequeDocType2Text(Area))
  frmAddEditReceiptChequeDoc.ShowMode = SHOW_EDIT
   Load frmAddEditReceiptChequeDoc
   frmAddEditReceiptChequeDoc.Show 1

   OKClick = frmAddEditReceiptChequeDoc.OKClick

   Unload frmAddEditReceiptChequeDoc
   Set frmAddEditReceiptChequeDoc = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitChequeDocOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      uctlChequeDate.ShowDate = Now
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
       m_ChequeDoc.CHEQUE_DOC_ID = -1
       m_ChequeDoc.CHEQUE_DOC_NO = PatchWildCard(txtChequeNo.Text)
       m_ChequeDoc.CUSTOMER_CODE = PatchWildCard(txtCustomerCode.Text)
       m_ChequeDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
       m_ChequeDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
       m_ChequeDoc.FROM_DATE = uctlChequeDate.ShowDate

      
      
      If Not glbDaily.QueryChequeDoc(m_ChequeDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 3000
   Col.Caption = MapText("�Ţ�����")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2160
  Col.Caption = MapText("�ѹ����͡���")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2160
    Col.Caption = MapText("�����١���")
    
       Set Col = GridEX1.Columns.add '5
   Col.Width = 5000
    Col.Caption = MapText("�����١���")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2160
   Col.Caption = MapText("�礼�ҹ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("�ѹ����͡���"))
   Call InitNormalLabel(lblCustomerCode, MapText("�����١���"))
   Call InitNormalLabel(lblChequeDocNo, MapText("�Ţ�����"))
'   Call InitNormalLabel(lblChequeNoEx, MapText("�Ţ�����"))
   Call InitNormalLabel(lblOrderBy, MapText("���§���"))
   Call InitNormalLabel(lblOrderType, MapText("���§�ҡ"))
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Enabled = False
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdClear, MapText("������ (F4)"))
   
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   m_HasActivate = False
   
   Set m_ChequeDoc = New CChequeDoc
   Set m_TempBillingDoc = New CChequeDoc
   Set m_Rs = New ADODB.Recordset
   Set m_Mr = New CMasterRef
      
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Mr = Nothing
   Set m_ChequeDoc = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim oMenu As cPopupMenu
'Dim lMenuChosen As Long
'Dim TempID1 As Long
'Dim Cd As CChequeDoc
'Dim IsOK As Boolean
'Dim OKClick As Boolean
'
'   If GridEX1.ItemCount <= 0 Then
'         Exit Sub
'   End If
'
'   TempID1 = GridEX1.Value(1)
'   If Button = 2 Then
'      Set oMenu = New cPopupMenu
'        lMenuChosen = oMenu.Popup("�礼�ҹ", "-", "������ҹ")
'      If lMenuChosen = 0 Then
'         Exit Sub
'      End If
'      Set oMenu = Nothing
'   Else
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, False)
'   If lMenuChosen = 1 Then
'
'     Call EnableForm(Me, False)
'      Call glbDaily.StartTransaction
'      Set Cd = New CChequeDoc
'      Cd.AddEditMode = SHOW_EDIT
'      Cd.CHEQUE_DOC_ID = TempID1
'      Cd.PASSCHEQUE_FLAG = "Y"
'      Cd.BADCHEQUE_FLAG = "N"
'      Cd.PASSCHEQUE_DATE = Now
'      Call Cd.UpdateChequeStatus
'      Call glbDaily.CommitTransaction
'      Call QueryData(True)
'      Set oMenu = Nothing
'      Set Cd = Nothing
'   ElseIf lMenuChosen = 3 Then
'
'       Call EnableForm(Me, False)
'      Call glbDaily.StartTransaction
'      Set Cd = New CChequeDoc
'      Cd.AddEditMode = SHOW_EDIT
'      Cd.CHEQUE_DOC_ID = TempID1
'      Cd.PASSCHEQUE_FLAG = "N"
'      Cd.BADCHEQUE_FLAG = "Y"
'      Cd.BADCHEQUE_DATE = Now
'      Call Cd.UpdateChequeStatus
'      Call glbDaily.CommitTransaction
'      Call QueryData(True)
'      Set oMenu = Nothing
'      Set Cd = Nothing
'
'   End If
'
'   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If DocumentType = CASH_PITTYCASH Then
   Else
      RowBuffer.RowStyle = RowBuffer.Value(4)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempBillingDoc.PopulateFromRS(1, m_Rs)

   Values(1) = m_TempBillingDoc.CHEQUE_DOC_ID
   Values(2) = m_TempBillingDoc.CHEQUE_DOC_NO
   Values(3) = DateToStringExtEx2(m_TempBillingDoc.CHEQUE_DOC_DATE)
   Values(4) = m_TempBillingDoc.CUSTOMER_CODE
    Values(5) = m_TempBillingDoc.LONG_NAME
    Values(6) = m_TempBillingDoc.PASSCHEQUE_FLAG

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
