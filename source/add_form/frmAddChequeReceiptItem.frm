VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddReceiptDocItem 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddChequeReceiptItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   3360
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5865
         Left            =   150
         TabIndex        =   3
         Top             =   1890
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   10345
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
         Column(1)       =   "frmAddChequeReceiptItem.frx":27A2
         Column(2)       =   "frmAddChequeReceiptItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddChequeReceiptItem.frx":290E
         FormatStyle(2)  =   "frmAddChequeReceiptItem.frx":2A6A
         FormatStyle(3)  =   "frmAddChequeReceiptItem.frx":2B1A
         FormatStyle(4)  =   "frmAddChequeReceiptItem.frx":2BCE
         FormatStyle(5)  =   "frmAddChequeReceiptItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddChequeReceiptItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   3360
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   5865
         Left            =   6540
         TabIndex        =   6
         Top             =   1890
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   10345
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
         Column(1)       =   "frmAddChequeReceiptItem.frx":2F36
         Column(2)       =   "frmAddChequeReceiptItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddChequeReceiptItem.frx":30A2
         FormatStyle(2)  =   "frmAddChequeReceiptItem.frx":31FE
         FormatStyle(3)  =   "frmAddChequeReceiptItem.frx":32AE
         FormatStyle(4)  =   "frmAddChequeReceiptItem.frx":3362
         FormatStyle(5)  =   "frmAddChequeReceiptItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddChequeReceiptItem.frx":34F2
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   5648
         TabIndex        =   5
         Top             =   5040
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeReceiptItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5648
         TabIndex        =   4
         Top             =   4470
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeReceiptItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   8280
         TabIndex        =   2
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeReceiptItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   12
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   1320
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   7
         Top             =   7860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeReceiptItem.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5970
         TabIndex        =   8
         Top             =   7860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddReceiptDocItem"
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
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection

Public AccountID As Long
Public ReceiptType As Long
Public CustomerID As Long
Public Area As Long

'Private Sub PopulateDestColl()
'Dim Ri As CReceiptItem
'Dim D As CBillingDoc
'
'   For Each Ri In TempCollection
'      Set D = New CBillingDoc
'
'      If Ri.Flag <> "D" Then
'         D.BILLING_DOC_ID = Ri.DO_ID
'         D.DOCUMENT_DATE = Ri.DOCUMENT_DATE
'         D.DOCUMENT_NO = Ri.DOCUMENT_NO
'         D.TEMP_PAID_AMOUNT = Ri.PAID_AMOUNT
'         Call m_TempCol2.add(D)
'      End If
'
'      Set D = Nothing
'   Next Ri
'End Sub
Private Sub PopulateDestColl()
'Dim Ri As CReceiptItem
Dim Ri As CReceiptChequeDoc
Dim D As CBillingDoc

   For Each Ri In TempCollection
      Set D = New CBillingDoc
      
      If Ri.Flag <> "D" Then
         D.BILLING_DOC_ID = Ri.BILLING_DOC_ID
         D.DOCUMENT_DATE = Ri.RECEIPT_CHEQUE_DOC_DATE
         D.DOCUMENT_NO = Ri.RECEIPT_CHEQUE_DOC_NO
         D.TEMP_PAID_AMOUNT = Ri.PAID_AMOUNT
         Call m_TempCol2.add(D)
      End If
      
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CBillingDoc
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.BILLING_DOC_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim Bd As CBillingDoc
Dim X As Double

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set Bd = New CBillingDoc
      Call Bd.PopulateFromRS(104, Rs)
      If Bd.DOCUMENT_NO = "01-00612-13" Then
         Debug.Print (Bd.DOCUMENT_NO)
      End If
      
      X = Val(Format(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.SUM_DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT) - Bd.PAID_AMOUNT - Bd.CASH_DISCOUNT - Bd.SUM_PAID_AMOUNT2, "0.00"))
      If X <> 0 And X > 0 Then       '�Դ��ҧ�ҡ�����ʹ˹��                            ���� ⴹ�Ѵ������������Ѻ����(�ѹ�Ѵ ��� �ҡ�Ѵ�����ѧ�еԴź )                '��������
         If Not IsIn(m_TempCol2, Bd.BILLING_DOC_ID) Then
            Call TempCol.add(Bd)
         End If
      End If
      
      Set Bd = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingDoc.COMMIT_FLAG = ""
      m_BillingDoc.FROM_DATE = uctlFromDate.ShowDate
      m_BillingDoc.TO_DATE = uctlDocumentDate.ShowDate
      m_BillingDoc.DOCUMENT_TYPE = 1
      m_BillingDoc.CUSTOMER_ID = CustomerID
'      If Area = 1 Then
'         m_BillingDoc.ACCOUNT_ID = AccountID
'         If ReceiptType = 3 Then  '���ҧ�ҡ Do
'            m_BillingDoc.DOCUMENT_TYPE = 1
'         ElseIf ReceiptType = 5 Then  '���ҧ�ҡ Invoice
'            m_BillingDoc.DOCUMENT_TYPE = 5
'         End If
'      ElseIf Area = 2 Then
'         m_BillingDoc.SUPPLIER_ID = AccountID
'         If ReceiptType = 3 Then  '���ҧ�ҡ Do
'            'm_BillingDoc.DOCUMENT_TYPE = 7
'            If cboDocumentType.ItemData(Minus2Zero(cboDocumentType.ListIndex)) > 0 Then
'               m_BillingDoc.DOCUMENT_TYPE = cboDocumentType.ItemData(Minus2Zero(cboDocumentType.ListIndex))
'            Else
'               m_BillingDoc.DOCUMENT_TYPE_SET = "(100,101,102,103)"
'            End If
'         ElseIf ReceiptType = 5 Then  '���ҧ�ҡ Invoice
'            m_BillingDoc.DOCUMENT_TYPE = 11
'         End If
'      End If
      m_BillingDoc.OrderBy = 10
      m_BillingDoc.OrderType = 1
'       m_BillingDoc.CUSTOMER_CODE =
'      If Area = 1 Then
  m_BillingDoc.ItemSumFlag = True
'      ElseIf Area = 2 Then
'         m_BillingDoc.ItemSumFlagSup = True
'      End If
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call PopulateTempColl
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
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

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CBillingDoc

   If ID > 0 Then
      Set L = TempCol1(ID)
      
      frmAddEditReceiptDocAmount.HeaderText = "�����Թ�١���"
      Set frmAddEditReceiptDocAmount.BillingDoc = L
      frmAddEditReceiptDocAmount.ShowMode = SHOW_EDIT
      Load frmAddEditReceiptDocAmount
      frmAddEditReceiptDocAmount.Show 1
      
      OKClick = frmAddEditReceiptDocAmount.OKClick
      
      Unload frmAddEditReceiptDocAmount
      Set frmAddEditReceiptDocAmount = Nothing
      
      If OKClick Then
         L.Flag = "A"
         Call TempCol2.add(L)
         TempCol1.Remove (ID)
      End If
   End If
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long
Dim D As CBillingDoc

'(D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT) - D.PAID_AMOUNT)
   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
      TempCol1(j).PAID_TYPE = 1
      Set D = TempCol1(j)
      TempCol1(j).TEMP_PAID_AMOUNT = D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT) - D.PAID_AMOUNT
      Call TempCol2.add(TempCol1(j))
   Next j
   
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long

   m_HasModify = True
   
   TempID = GridEX1.Row
   Call CopyItem(m_TempCol1, m_TempCol2, TempID)

   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)
   
   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Public Sub PopulateTempColl()
Dim D As CBillingDoc
Dim Ri As CReceiptChequeDoc
'Dim Ri As CReceiptItem

   For Each D In m_TempCol2
      Set Ri = New CReceiptChequeDoc

      If D.Flag = "A" Then
         Ri.Flag = "A"
         Ri.BILLING_DOC_ID = D.BILLING_DOC_ID
         Ri.RECEIPT_CHEQUE_DOC_DATE = D.DOCUMENT_DATE
         Ri.RECEIPT_CHEQUE_DOC_NO = D.DOCUMENT_NO
         Ri.RECEIPT_ITEM_AMOUNT = D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT)
         Ri.CASH_DISCOUNT = D.CASH_DISCOUNT
         If D.PAID_TYPE = 1 Then '�������ǧ�Թ
            Ri.PAID_AMOUNT = (D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT) - D.PAID_AMOUNT)
         Else
            Ri.PAID_AMOUNT = D.TEMP_PAID_AMOUNT
         End If

'         Call GenerateCnDnItem(Ri.DO_ID)
         Call TempCollection.add(Ri)
      End If

      Set Ri = Nothing
   Next D
End Sub
'Public Sub PopulateTempColl()
'Dim D As CBillingDoc
'Dim RCD As CReceiptChequeDoc
'
'   For Each D In m_TempCol2
'      Set RCD = New CReceiptChequeDoc
'
'      If D.Flag = "A" Then
'         RCD.Flag = "A"
'         RCD.BILLING_DOC_ID = D.BILLING_DOC_ID
'         RCD.RECEIPT_CHEQUE_DOC_DATE = D.DOCUMENT_DATE
'         RCD.RECEIPT_CHEQUE_DOC_NO = D.DOCUMENT_NO
'        RCD.RECEIPT_ITEM_AMOUNT = D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT)
'         RCD.CASH_DISCOUNT = D.CASH_DISCOUNT
'         If D.PAID_TYPE = 1 Then '�������ǧ�Թ
'            RCD.PAID_AMOUNT = (D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT) - D.PAID_AMOUNT)
'         Else
'            RCD.PAID_AMOUNT = D.TEMP_PAID_AMOUNT
'         End If
'
''         Call GenerateCnDnItem(Ri.DO_ID)
'         Call TempCollection.add(RCD)
'      End If
'
'      Set RCD = Nothing
'   Next D
'End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      Call InitDocumentTypeSup(cboDocumentType)
      
      Call EnableForm(Me, False)
      Call PopulateDestColl
'      If Area = 1 Then
         If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
            m_BillingDoc.QueryFlag = 1
            Call QueryData(True)
         ElseIf ShowMode = SHOW_ADD Then
            m_BillingDoc.QueryFlag = 0
            Call QueryData(True)
         End If
'      End If
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_Employees = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 1650
   Col.Caption = MapText("�ѹ����͡���")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1830
   Col.Caption = MapText("�����Ţ�͡���")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1665
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ʹ˹��")
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX2.Columns.add '3
   Col.Width = 1650
   Col.Caption = MapText("�ѹ����͡���")

   Set Col = GridEX2.Columns.add '4
   Col.Width = 1830
   Col.Caption = MapText("�����Ţ�͡���")

   Set Col = GridEX2.Columns.add '5
   Col.Width = 1665
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ʹ����")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_BillingDoc.ImportExports
'      If II.Flag <> "D" Then
'         Sum = Sum + CDbl(Format(II.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(II.EXPORT_AMOUNT, "0.00"))
'      End If
'   Next II
''
''   txtDeliveryFee.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
'   Call InitNormalLabel(lblDocumentType, MapText("�ҡ�͡���"))
   
   Call InitNormalLabel(lblDocumentDate, MapText("�֧�ѹ���"))
   Call InitNormalLabel(lblFromDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdSearch, MapText("����"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   
'   Call InitCombo(cboDocumentType)
   
   Call InitGrid1
   Call InitGrid2
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
   Set m_Employees = New Collection
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim X As Double

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBillingDoc
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   X = (CR.DO_TOTAL_PRICE + CR.REVENUE_TOTAL_PRICE - CR.SUM_DISCOUNT_AMOUNT)
   'X = X + (CR.DEBIT_AMOUNT - CR.CREDIT_AMOUNT) - CR.CASH_DISCOUNT - CR.SUM_PAID_AMOUNT2
   X = X + (CR.DEBIT_AMOUNT - CR.CREDIT_AMOUNT) - CR.PAID_AMOUNT - CR.CASH_DISCOUNT - CR.SUM_PAID_AMOUNT2
   Values(5) = FormatNumber(X)
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBillingDoc
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Values(5) = FormatNumber(CR.TEMP_PAID_AMOUNT)
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
