VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddRQItem 
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15690
   Icon            =   "frmAddRQItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   15690
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   9825
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   17330
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtDoNo 
         Height          =   405
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
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
         Height          =   6855
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   12091
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
         Column(1)       =   "frmAddRQItem.frx":27A2
         Column(2)       =   "frmAddRQItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddRQItem.frx":290E
         FormatStyle(2)  =   "frmAddRQItem.frx":2A6A
         FormatStyle(3)  =   "frmAddRQItem.frx":2B1A
         FormatStyle(4)  =   "frmAddRQItem.frx":2BCE
         FormatStyle(5)  =   "frmAddRQItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddRQItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   16005
         _ExtentX        =   28231
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   6855
         Left            =   8280
         TabIndex        =   6
         Top             =   2160
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   12091
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
         Column(1)       =   "frmAddRQItem.frx":2F36
         Column(2)       =   "frmAddRQItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddRQItem.frx":30A2
         FormatStyle(2)  =   "frmAddRQItem.frx":31FE
         FormatStyle(3)  =   "frmAddRQItem.frx":32AE
         FormatStyle(4)  =   "frmAddRQItem.frx":3362
         FormatStyle(5)  =   "frmAddRQItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddRQItem.frx":34F2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   7260
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   7260
         TabIndex        =   1
         Top             =   1350
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11160
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddRQItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblDoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   900
         Width           =   1155
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   7560
         TabIndex        =   5
         Top             =   6000
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddRQItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   7560
         TabIndex        =   4
         Top             =   5400
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddRQItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11160
         TabIndex        =   2
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddRQItem.frx":4018
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5970
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
         Left            =   5970
         TabIndex        =   11
         Top             =   1440
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6120
         TabIndex        =   7
         Top             =   9120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddRQItem.frx":4332
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7920
         TabIndex        =   8
         Top             =   9120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddRQItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_RQ As CInventoryDoc

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection
Public T_CBillingDoc As CBillingDoc
Public BILLING_DOC_ID As Long
Public SumLotAmount As Double
Public temp_So As Collection
Public temp_SumDO As Collection

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Public m_TempCol2 As Collection

Public AccountID As Long
Public LocationID As Long
Public Area As Long
Public DOCUMENT_TYPE As Long
Public DocumentNo As String
Public DocumentDate As Date
Public CustomerID As Long
Public CustomerCode As String
Public TruckNo As String
Public NOTE As String


Private Sub PopulateDestColl()
Dim Ri As CDoItem
Dim D As CDoItem

   If TempCollection Is Nothing Then
      Exit Sub
   End If
   
   For Each Ri In TempCollection
      Set D = New CDoItem
      If Ri.Flag <> "D" Then
         Call D.CopyObject(1, Ri)
         Call m_TempCol2.add(D)
      End If
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CDoItem
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.PO_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection, Optional TempCol2 As Collection)
'Dim BD As CBillingDoc
'Dim Temp_IWD As CInventoryWHDoc
'Set m_TempCol1 = Nothing
'Set m_TempCol1 = New Collection
'   Rs.MoveFirst
'   While Not Rs.EOF
'      Set BD = New CBillingDoc
'      If Area = 4 Or Area = 5 Or Area = 6 Then
'         Call BD.PopulateFromRS(114, Rs)
'         Set Temp_IWD = GetObject("CInventoryWHDoc", TempCol2, str(BD.BILLING_DOC_ID), False)
'         If Not Temp_IWD Is Nothing Then
'            BD.TRUCK_NO = Temp_IWD.TRUCK_NO
'            BD.LOAD_GOODS_NO = Temp_IWD.LOAD_GOODS_NO
'            BD.DO_NO = Temp_IWD.DO_NO
'            BD.LOAD_FLAG = Temp_IWD.LOAD_FLAG
'            BD.NOTE = Temp_IWD.NOTE
'            BD.INVENTORY_WH_DOC_ID = Temp_IWD.INVENTORY_WH_DOC_ID
'
'            Call TempCol.add(BD)
'         End If
'      Else
'         Call BD.PopulateFromRS(113, Rs)
'         Call TempCol.add(BD)
'      End If
'      Set BD = Nothing
'      Rs.MoveNext
'   Wend
'Dim BD As CBillingDoc
''Dim Temp_IWD As CInventoryWHDoc
'Dim Temp_RQ As CInventoryDoc
'Set m_TempCol1 = Nothing
'Set m_TempCol1 = New Collection
'   Rs.MoveFirst
'   While Not Rs.EOF
'      Set BD = New CBillingDoc
'      If Area = 4 Or Area = 5 Or Area = 6 Then
'         Call BD.PopulateFromRS(114, Rs)
'         Set Temp_RQ = GetObject("CInventoryDoc", TempCol2, str(BD.BILLING_DOC_ID), False)
'         If Not Temp_RQ Is Nothing Then
'            BD.TRUCK_NO = Temp_IWD.TRUCK_NO
'            BD.LOAD_GOODS_NO = Temp_IWD.LOAD_GOODS_NO
'            BD.DO_NO = Temp_IWD.DO_NO
'            BD.LOAD_FLAG = Temp_IWD.LOAD_FLAG
'            BD.NOTE = Temp_IWD.NOTE
'            BD.INVENTORY_WH_DOC_ID = Temp_IWD.INVENTORY_WH_DOC_ID
'
'            Call TempCol.add(BD)
'         End If
'      Else
''         Call BD.PopulateFromRS(113, Rs)
''         Call TempCol.add(BD)
'      End If
'      Set BD = Nothing
'      Rs.MoveNext
'   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim TempColl As Collection

Dim tSo As CSaleOrder
Dim tDo As CDoItem
Dim Total As Double

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_RQ.INVENTORY_DOC_ID = -1
      m_RQ.INVENTORY_WH_DOC_ID = -1
      m_RQ.COMMIT_FLAG = ""
      m_RQ.DOCUMENT_NO = DocumentNo
      m_RQ.DOCUMENT_DATE = -1
      m_RQ.FROM_DATE = uctlFromDate.ShowDate
      m_RQ.TO_DATE = uctlToDate.ShowDate
      m_RQ.CUSTOMER_ID = CustomerID
      m_RQ.DOCUMENT_TYPE = DOCUMENT_TYPE 'ใบ So
      
         Call m_RQ.QueryData(m_Rs, ItemCount, 3)
         Set m_TempCol1 = New Collection
         While Not m_Rs.EOF
            Set m_RQ = Nothing
            Set m_RQ = New CInventoryDoc
            Call m_RQ.PopulateFromRS(3, m_Rs)
            
            Set temp_SumDO = Nothing
            Set temp_SumDO = New Collection
            Set temp_So = Nothing
            Set temp_So = New Collection
            
            Call LoadSupDoItem(Nothing, temp_SumDO, -1, -1, m_RQ.BILLING_DOC_ID)
            Call LoadSaleOrder(Nothing, temp_So, -1, -1, m_RQ.INVENTORY_WH_DOC_ID)
            
            Total = 0
            For Each tSo In temp_So
               Set tDo = GetObject("CDoItem", temp_SumDO, Trim(str(tSo.BILLING_DOC_SO_ID)) & "-" & Trim(str(tSo.PART_ITEM_ID)), False)
               If Not tDo Is Nothing Then
                  tSo.PACK_AMOUNT = tSo.PACK_AMOUNT - tDo.PACK_AMOUNT
               End If
               Total = Total + tSo.PACK_AMOUNT
            Next tSo
   
            If Total > 0 Then
               Call m_TempCol1.add(m_RQ)
            End If
            m_Rs.MoveNext
        Wend
   End If

   If ItemCount > 0 Then
'      Call GenerateSourceItem(m_Rs, m_TempCol1, TempColl) ' เดี๋ยวมาทำต่อ
      GridEX1.ItemCount = CountItem(m_TempCol1)
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
      Call MsgBox("ไม่มีข้อมูลใบฝากขาย กรุณาตรวจสอบเงื่อนไขอีกครั้ง", vbOKOnly)
   End If
   
   If Not (m_TempCol2 Is Nothing) Then 'TempCollection
      GridEX2.ItemCount = CountItem(m_TempCol2) 'TempCollection.Count
      GridEX2.Rebind
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
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   Call PopulateTempColl(SumLotAmount)

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

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub


Private Sub cmdClear_Click()
   txtDoNo.Text = ""
End Sub

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
      uctlFromDate.ShowDate = uctlFromDate.ShowDate
      uctlToDate.ShowDate = uctlToDate.ShowDate

      DocumentNo = txtDoNo.Text
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_RQ.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_RQ.QueryFlag = 0
         Call QueryData(True)
      End If
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long, Optional DocType As Long = -1)
Dim L As CInventoryDoc
Dim OKClick As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Poi As CSaleOrder
Dim Poi2 As CDoItem
Dim Di As CDoItem
Dim IsOK As Boolean
Dim TempDoc As String
Dim TempTruck As String
Dim c_LIW As Collection
Dim temp_LIW As CLotItemWH
Dim PackAmount As Double
Dim temp_DI As CDoItem

   Set TempRs = New ADODB.Recordset
   If ID > 0 Then
      Set L = TempCol1(ID)
      If DocType = 19 Then
               'copy to Billing_Doc
           T_CBillingDoc.NOTE = L.TRUCK_NO
           T_CBillingDoc.REF = L.LOAD_GOODS_NO
           T_CBillingDoc.REFERENCE = L.DOCUMENT_NO
           T_CBillingDoc.PAYMENT_DESC = L.NOTE
           T_CBillingDoc.TEMP_DO_NO = L.DO_NO
           T_CBillingDoc.INVENTORY_WH_DOC_ID = L.INVENTORY_WH_DOC_ID
           T_CBillingDoc.LOAD_FLAG = L.LOAD_FLAG
           T_CBillingDoc.B_SUCCESS_FLAG = L.B_SUCCESS_FLAG
           T_CBillingDoc.BILLING_DOC_SO_ID = L.BILLING_DOC_ID
           T_CBillingDoc.DELIVERY_CUS_ITEM_ID = L.DELIVERY_CUS_ITEM_ID
           T_CBillingDoc.PRICE_THINK_TYPE = L.PRICE_THINK_TYPE
           T_CBillingDoc.USER_APPLOVE_PRICE_THINK = L.USER_APPLOVE_PRICE_THINK
           T_CBillingDoc.INVENTORY_DOC_TRN_ID = L.INVENTORY_DOC_ID
      'end Copy
         For Each Poi In temp_So
            If Poi.PACK_AMOUNT > 0 Then
               Set Di = New CDoItem
               Call Di.CopyObjectFromSo(1, Poi)
               Di.Flag = "A"
               Di.PO_ID = Poi.DO_ID
               Di.PO_NO = Poi.DOCUMENT_NO
               Di.BILLING_DOC_ID = Poi.DO_ID
               If Di.PART_ITEM_ID = -1 Then
                  Di.ITEM_AMOUNT = Di.PACK_AMOUNT * Di.WEIGHT_PER_PACK_SO
               Else
                  Di.ITEM_AMOUNT = Di.PACK_AMOUNT * Di.WEIGHT_PER_PACK
               End If
               Di.TX_AMOUNT = Di.ITEM_AMOUNT
               Di.TOTAL_PRICE = Di.PACK_AMOUNT * Di.PRICE_PER_PACK
            
                  If Di.PART_ITEM_ID > -1 Then 'เลือกเฉพาะที่เป็นสินค้าเท่านั้น
                    Set temp_DI = GetObject("CDoItem", TempCol2, Trim(str(Di.PO_ID) & "-" & str(Di.PART_ITEM_ID)), False)
                    If Not temp_DI Is Nothing Then
                        Call temp_DI.CopyObjectFromSo(1, Poi)
                        If temp_DI.Flag <> "A" Then
                           temp_DI.Flag = "E"
                        End If
                        temp_DI.PO_ID = Poi.DO_ID
                        temp_DI.PO_NO = Poi.DOCUMENT_NO
                        temp_DI.BILLING_DOC_ID = Poi.DO_ID
                    Else
                         Call TempCol2.add(Di, Trim(str(Di.PO_ID) & "-" & str(Di.PART_ITEM_ID)))
                     End If
               Else
                  Call TempCol2.add(Di)
               End If 'end  If Area = 1 Then
               Set Di = Nothing
            End If
         Next Poi
      End If
      If temp_So.Count > 0 Then
         Call TempCol1.Remove(ID)
      Else
         glbErrorLog.LocalErrorMsg = "เอกสารใบนี้ประเภทการขายไม่ถูกต้อง"
         glbErrorLog.ShowUserError
      End If
   End If
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
End Sub
Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim J As Long

   For J = 1 To TempCol1.Count
      TempCol1(J).Flag = "A"
      TempCol1(J).IncludeFlag = True
      Call TempCol2.add(TempCol1(J))
   Next J
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long
Dim tSo As CSaleOrder
Dim tDo As CDoItem

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   m_HasModify = True
   
   TempID = GridEX1.row
   Call LoadSupDoItem(Nothing, temp_SumDO, -1, -1, GridEX1.Value(7))
   Call LoadSaleOrder(Nothing, temp_So, -1, -1, GridEX1.Value(6))
   
   For Each tSo In temp_So
         Set tDo = GetObject("CDoItem", temp_SumDO, Trim(str(tSo.BILLING_DOC_SO_ID)) & "-" & Trim(str(tSo.PART_ITEM_ID)), False)
         If Not tDo Is Nothing Then
            tSo.PACK_AMOUNT = tSo.PACK_AMOUNT - tDo.PACK_AMOUNT
         End If
   Next tSo
   
   Call CopyItem(m_TempCol1, m_TempCol2, TempID, 19)

   GridEX1.ItemCount = CountItem(m_TempCol1)
   GridEX1.Rebind

   GridEX2.ItemCount = CountItem(m_TempCol2)
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
Public Sub PopulateTempColl2(Tempsum As Double)
Dim D As CDoItem
Dim Ri As CLotItemWH
Dim Sum As Double
Dim temp_WH As CLotItemWH

   Sum = 0
   For Each D In m_TempCol2
      Set Ri = New CLotItemWH
      If (D.Flag = "A") Then
         Call Ri.CopyObjectFromDoItem(1, D)
         Ri.Flag = "A"
         Ri.AddEditMode = SHOW_ADD
         Ri.TX_TYPE = "E"
         Sum = Sum + Ri.PACK_AMOUNT
         Call TempCollection.add(Ri)
      ElseIf D.Flag = "E" Or D.Flag = "D" Then
         For Each Ri In TempCollection
            If Ri.BILLING_DOC_ID = D.BILLING_DOC_ID And Ri.PART_ITEM_ID = D.PART_ITEM_ID Then
                     Call Ri.CopyObjectFromDoItem(1, D)
                    Ri.Flag = D.Flag '"E"
            End If
         Next Ri
      End If
      Set Ri = Nothing
   Next D
   
   Tempsum = Sum
End Sub
Public Sub PopulateTempColl(Tempsum As Double)
Dim D As CDoItem
Dim Ri As CDoItem
Dim Sum As Double
Dim D2 As CDoItem
Dim LIW As CLotItemWH

   Sum = 0
   For Each D In m_TempCol2
      Set Ri = New CDoItem
      If (D.Flag = "A") Then
         Call Ri.CopyObject(1, D)
         Ri.PO_ID = D.PO_ID
         Ri.PO_NO = D.PO_NO
         Ri.Flag = "A"
         Call TempCollection.add(Ri)
      End If
      Set Ri = Nothing
   Next D
   
   Tempsum = Sum
End Sub

Private Sub Form_Activate()
Dim firstDate As Date
Dim lastDate As Date
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)

      Call GetFirstLastDate(Now, firstDate, lastDate)
      uctlFromDate.ShowDate = firstDate
      uctlToDate.ShowDate = DocumentDate
      txtDoNo.Text = DocumentNo

      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_RQ.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_RQ.QueryFlag = 0
         Call QueryData(True)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
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
   
   Set m_RQ = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing '
   Set temp_So = Nothing
   Set temp_SumDO = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdSelect_Click
End Sub

Private Sub GridEX2_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 1575
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1900
   Col.Caption = MapText("เลขที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 4000
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 0
   Col.Caption = MapText("INVENTORY_WH_DOC_ID")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 0
   Col.Caption = MapText("BILLING_DOC_SO_ID")
   
'   If Area = 4 Or Area = 5 Or Area = 6 Then
'      Set Col = GridEX1.Columns.add '6
'      Col.Width = 2100
'      Col.Caption = MapText("เลขใบขึ้นอาหาร")
'
'      Set Col = GridEX1.Columns.add '7
'      Col.Width = 0
'      Col.Caption = MapText("LOAD GOODS ID")
'  Else
'      Set Col = GridEX1.Columns.add '6
'      Col.Width = 0
'      Col.Caption = MapText("")
'   End If
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.NAME = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation

   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX2.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("รหัสสินค้า")
   
'   Set Col = GridEX2.Columns.add '4
'   Col.Width = 2100
'   Col.Caption = MapText("รายละเอียด")
   
   Set Col = GridEX2.Columns.add '4
   Col.Width = 1700
   Col.Caption = MapText("เลขที่ใบฝากขาย")
   
   Set Col = GridEX2.Columns.add '5
   Col.Width = 2000
   Col.Caption = MapText("เลขที่ Sale Order")
   
   Set Col = GridEX2.Columns.add '6
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")

'  If Area = 1 Then
'       Set Col = GridEX2.Columns.add '3
'      Col.Width = 2100
'      Col.Caption = MapText("รายละเอียด")
'
'      Set Col = GridEX2.Columns.add '3
'      Col.Width = 2100
'      Col.Caption = MapText("เลขที่เอกสาร")
'   Else
'      Set Col = GridEX2.Columns.add '3
'      Col.Width = 2100
'      Col.Caption = MapText("รายละเอียด")
'
'      Set Col = GridEX2.Columns.add '3
'      Col.Width = 0
'      Col.Caption = MapText("")
'   End If
'
'   If Area = 1 Or Area = 4 Or Area = 5 Or Area = 6 Then
'      Set Col = GridEX2.Columns.add '4
'      Col.Width = 1000
'      Col.TextAlignment = jgexAlignRight
'      Col.Caption = MapText("จำนวน")
'   End If

End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_PO.ImportExports
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
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblDoNo, MapText("เลขที่ใบฝากขาย"))
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   
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
   Set m_RQ = New CInventoryDoc
   Set m_TempCol1 = New Collection
   Set temp_So = New Collection
   Set temp_SumDO = New Collection
   Set m_TempCol2 = New Collection
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CInventoryDoc
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = CR.INVENTORY_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Values(5) = CR.CUSTOMER_NAME
   Values(6) = CR.INVENTORY_WH_DOC_ID
   Values(7) = CR.BILLING_DOC_ID
     
'   If Area = 4 Or Area = 5 Or Area = 6 Then
'      Values(6) = CR.LOAD_GOODS_NO
'      Values(7) = CR.INVENTORY_WH_DOC_ID
'   Else
'       Values(6) = ""
'   End If
'   Values(7) = CR.TOTAL_AMOUNT

   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX2_DblClick()

  frmAddEditAmount.ID = GridEX2.Value(2)
  frmAddEditAmount.ShowMode = SHOW_EDIT
  Set frmAddEditAmount.TempCollection = m_TempCol2
'  frmAddEditAmount.PART_NO = GridEX2.Value(3)
'  frmAddEditAmount.PACK_AMOUNT = GridEX2.Value(6)
   Load frmAddEditAmount
   frmAddEditAmount.Show 1

   Unload frmAddEditAmount
   Set frmAddEditAmount = Nothing
   
   GridEX2.ItemCount = CountItem(m_TempCol2)
   GridEX2.Rebind
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CDoItem
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.PART_ITEM_ID
   Values(2) = RealIndex
   If CR.PART_ITEM_ID > 0 Then
      Values(3) = CR.PART_NO
   ElseIf CR.PART_ITEM_ID = -1 Then
      Values(3) = CR.FEATURE_CODE
   Else
      Values(3) = ""
   End If
'  Values(4) = CR.ShowDescText
  Values(4) = CR.PO_NO
  Values(5) = CR.BILLING_DOC_SO_NO
  Values(6) = CR.PACK_AMOUNT
      
'   If Area = 1 Then
'      Values(4) = CR.ShowDescText
'      Values(5) = CR.PO_NO ' CR.ShowDescText ' CR.PART_NO '
'   Else
'      Values(4) = CR.ShowDescText
'       Values(5) = ""
'   End If
'   If DOCUMENT_TYPE = 2000 Then
'      Values(6) = CR.PACK_AMOUNT
'   ElseIf DOCUMENT_TYPE = 2001 Then
'      Values(6) = CR.ITEM_AMOUNT
'   Else
'      Values(6) = CR.PACK_AMOUNT
'   End If
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
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

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
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

Private Sub SSCommand2_Click()

End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub

