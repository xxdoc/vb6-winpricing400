VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPlanPartOutFormulaItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "frmPlanPartOutFormulaItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8955
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   15796
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   4
         Top             =   0
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtItemAmount 
         Height          =   435
         Left            =   2700
         TabIndex        =   0
         Top             =   870
         Width           =   5415
         _ExtentX        =   20558
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   5385
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9499
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
         Column(1)       =   "frmPlanPartOutFormulaItem.frx":27A2
         Column(2)       =   "frmPlanPartOutFormulaItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPlanPartOutFormulaItem.frx":290E
         FormatStyle(2)  =   "frmPlanPartOutFormulaItem.frx":2A6A
         FormatStyle(3)  =   "frmPlanPartOutFormulaItem.frx":2B1A
         FormatStyle(4)  =   "frmPlanPartOutFormulaItem.frx":2BCE
         FormatStyle(5)  =   "frmPlanPartOutFormulaItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmPlanPartOutFormulaItem.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtProduct 
         Height          =   435
         Left            =   2730
         TabIndex        =   7
         Top             =   2160
         Width           =   5415
         _ExtentX        =   20558
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtProductAmount 
         Height          =   435
         Left            =   2730
         TabIndex        =   9
         Top             =   2640
         Width           =   2415
         _ExtentX        =   20558
         _ExtentY        =   767
      End
      Begin VB.Label lblProductAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1080
         TabIndex        =   10
         Top             =   2730
         Width           =   1575
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1080
         TabIndex        =   8
         Top             =   2250
         Width           =   1575
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1050
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5250
         TabIndex        =   2
         Top             =   1410
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3600
         TabIndex        =   1
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPartOutFormulaItem.frx":2F36
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmPlanPartOutFormulaItem"
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
Public BatchAmount As Double

Private m_Rs As ADODB.Recordset
Private m_Formula As CFormula

Public TempCollectionPlanIn As Collection
Public TempCollectionPlanOut As Collection

Private Sub cmdOK_Click()
Dim TempPlanOut As CFormulaItem
Dim TempPlanIn As CFormulaItem
Dim Fi As CFormulaItem
   
   If Val(txtItemAmount.Text) > 0 Then
      OKClick = True
      BatchAmount = Val(txtItemAmount.Text)
         
      For Each Fi In m_Formula.Inputs
         If Fi.Flag <> "D" Then
            Set TempPlanOut = GetObject("CFormulaItem", TempCollectionPlanOut, Trim(Str(Fi.PART_ITEM_ID)), False)
            If TempPlanOut Is Nothing Then
               Set TempPlanOut = New CFormulaItem
               TempPlanOut.PART_ITEM_ID = Fi.PART_ITEM_ID
               TempPlanOut.PART_NO = Fi.PART_NO
               TempPlanOut.PART_ITEM_NAME = Fi.PART_ITEM_NAME
               TempPlanOut.REAL_AMOUNT = Fi.REAL_AMOUNT * BatchAmount
               
               Call TempCollectionPlanOut.add(TempPlanOut, Trim(Str(TempPlanOut.PART_ITEM_ID)))
            Else
               TempPlanOut.REAL_AMOUNT = TempPlanOut.REAL_AMOUNT + (Fi.REAL_AMOUNT * BatchAmount)
            End If
         End If
      Next Fi
      Set TempPlanOut = Nothing
      
      Set TempPlanIn = GetObject("CFormulaItem", TempCollectionPlanIn, Trim(Str(m_Formula.PART_ITEM_ID)), False)
      If TempPlanIn Is Nothing Then
         Set TempPlanIn = New CFormulaItem
         TempPlanIn.PART_ITEM_ID = m_Formula.PART_ITEM_ID
         TempPlanIn.PART_NO = m_Formula.PART_NO
         TempPlanIn.PART_ITEM_NAME = m_Formula.PART_ITEM_NAME
         TempPlanIn.REAL_AMOUNT = CalculateTotalRatio * BatchAmount
         
         Call TempCollectionPlanIn.add(TempPlanIn, Trim(Str(TempPlanIn.PART_ITEM_ID)))
      Else
         TempPlanIn.REAL_AMOUNT = TempPlanIn.REAL_AMOUNT + (CalculateTotalRatio * BatchAmount)
      End If
      Set TempPlanIn = Nothing
   Else
      txtItemAmount.SetFocus
   End If
   Unload Me
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblItemAmount, MapText("จำนวน BATCH"))
   
   Call InitNormalLabel(lblProduct, MapText("ผลิตภัณฑ์"))
   Call InitNormalLabel(lblProductAmount, MapText("ยอดผลิตภัณฑ์"))
   
   Call txtItemAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
   txtItemAmount.Text = "1"
   
   txtProduct.Enabled = False
   txtProductAmount.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
   
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
   m_HasActivate = False
   
   Set m_Formula = New CFormula
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.itemcount = 0
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
   
   Set Col = GridEX2.Columns.add '2
   Col.Width = 1500
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX2.Columns.add '3
   Col.Width = 3000
   Col.Caption = MapText("ชื่อวัตถุดิบ")
   
   Set Col = GridEX2.Columns.add '4
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนัก (Kg)")
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Formula.FORMULA_ID = ID
      m_Formula.QueryFlag = 1
      If Not glbProduction.QueryFormula(m_Formula, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
     End If
   End If
   
   Call m_Formula.PopulateFromRS(1, m_Rs)
   
   txtProduct.Text = m_Formula.PART_ITEM_NAME & "(" & m_Formula.PART_NO & ")"
   txtProductAmount.Text = CalculateTotalRatio * Val(txtItemAmount.Text)
   
   GridEX2.itemcount = m_Formula.Inputs.Count
   GridEX2.Rebind
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Formula = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub txtItemAmount_Change()
   If Not m_Formula Is Nothing Then
      txtProductAmount.Text = CalculateTotalRatio * Val(txtItemAmount.Text)

      GridEX2.itemcount = m_Formula.Inputs.Count
      GridEX2.Rebind
   End If
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Formula.Inputs Is Nothing Then
      Exit Sub
   End If

     If RowIndex <= 0 Then
         Exit Sub
      End If

   Dim Ci As CFormulaItem
   If m_Formula.Inputs.Count <= 0 Then
      Exit Sub
   End If
   Set Ci = GetItem(m_Formula.Inputs, RowIndex, RealIndex)
   If Ci Is Nothing Then
      Exit Sub
   End If

   Values(1) = Ci.FORMULA_ITEM_ID
   Values(2) = RealIndex
   Values(3) = Ci.PART_NO
   Values(4) = Ci.PART_ITEM_NAME
   Values(5) = FormatNumber(Ci.REAL_AMOUNT * Val(txtItemAmount.Text), 3)

   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Function CalculateTotalRatio() As Double
Dim D As CFormulaItem
Dim Sum3 As Double

   Sum3 = 0
   For Each D In m_Formula.Inputs
      If D.Flag <> "D" Then
         Sum3 = Sum3 + D.REAL_AMOUNT
      End If
   Next D
   
   CalculateTotalRatio = Sum3
End Function


