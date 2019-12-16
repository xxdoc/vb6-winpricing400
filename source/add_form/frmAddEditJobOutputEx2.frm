VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditJobOutputEx2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditJobOutputEx2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   9675
      Left            =   0
      TabIndex        =   34
      Top             =   600
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   17066
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlStartDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboHead 
         Height          =   510
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   800
      End
      Begin VB.ComboBox cboLotNo 
         Height          =   510
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   2085
      End
      Begin prjFarmManagement.uctlTime txtTimePackBegin 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   6720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlDate uctlPackDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   6240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin VB.ComboBox cboLockNo 
         Height          =   510
         Left            =   5445
         TabIndex        =   8
         Top             =   2820
         Width           =   1725
      End
      Begin VB.ComboBox cboBinNo 
         Height          =   510
         Left            =   1800
         TabIndex        =   7
         Top             =   2820
         Width           =   2085
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   660
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   17
         Top             =   5280
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   24
         Top             =   7200
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRef 
         Height          =   435
         Left            =   5040
         TabIndex        =   28
         Top             =   8640
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSerialNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   27
         Top             =   8640
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtStdAmount 
         Height          =   435
         Left            =   5445
         TabIndex        =   20
         Top             =   5760
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightPerPack 
         Height          =   435
         Left            =   1800
         TabIndex        =   14
         Top             =   4320
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductTypeLookup 
         Height          =   435
         Left            =   3840
         TabIndex        =   3
         Top             =   1080
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtGoodAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   15
         Top             =   4800
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLoseAmount 
         Height          =   435
         Left            =   5445
         TabIndex        =   16
         Top             =   4800
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRestAmount 
         Height          =   435
         Left            =   5445
         TabIndex        =   18
         Top             =   5280
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   1800
         TabIndex        =   19
         Top             =   5760
         Width           =   1725
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTime uctlTime1 
         Height          =   375
         Left            =   60500
         TabIndex        =   54
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTime txtTimePackEnd 
         Height          =   375
         Left            =   6045
         TabIndex        =   23
         Top             =   6720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1800
         TabIndex        =   29
         Top             =   9120
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletFrom 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3360
         Width           =   765
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletTo 
         Height          =   435
         Left            =   3240
         TabIndex        =   10
         Top             =   3360
         Width           =   765
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletPerUnit 
         Height          =   435
         Left            =   5880
         TabIndex        =   11
         Top             =   3360
         Width           =   885
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletFrom2 
         Height          =   435
         Left            =   1800
         TabIndex        =   12
         Top             =   3840
         Width           =   765
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPalletPerUnit2 
         Height          =   435
         Left            =   4200
         TabIndex        =   13
         Top             =   3840
         Width           =   885
         _ExtentX        =   4524
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLotNoNew 
         Height          =   435
         Left            =   5460
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookupRest 
         Height          =   435
         Left            =   1800
         TabIndex        =   25
         Top             =   7680
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPlaceLookupLose 
         Height          =   435
         Left            =   1800
         TabIndex        =   26
         Top             =   8160
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblPlaceLose 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlaceLose"
         Height          =   315
         Left            =   0
         TabIndex        =   69
         Top             =   8160
         Width           =   1695
      End
      Begin VB.Label lblPlaceRest 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlaceRest"
         Height          =   315
         Left            =   0
         TabIndex        =   68
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Label lblStartDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartDate"
         Height          =   315
         Left            =   480
         TabIndex        =   66
         Top             =   1680
         Width           =   1155
      End
      Begin Threed.SSCommand cmdEditUnitPerPallet 
         Height          =   405
         Left            =   6720
         TabIndex        =   32
         Top             =   3840
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx2.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblHead 
         Alignment       =   1  'Right Justify
         Caption         =   "lblHead"
         Height          =   345
         Left            =   360
         TabIndex        =   65
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblLotNo2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo2"
         Height          =   315
         Left            =   4440
         TabIndex        =   64
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblLotNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblJobNo"
         Height          =   315
         Left            =   480
         TabIndex        =   63
         Top             =   2280
         Width           =   1155
      End
      Begin Threed.SSCommand cmdAuto2 
         Height          =   405
         Left            =   3960
         TabIndex        =   67
         Top             =   2325
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx2.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUnit1"
         Height          =   345
         Left            =   5160
         TabIndex        =   62
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label lblPalletPerUnit2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletPerUnit"
         Height          =   345
         Left            =   2520
         TabIndex        =   61
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblPalletFrom2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletFrom"
         Height          =   345
         Left            =   0
         TabIndex        =   60
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblUnit1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblUnit1"
         Height          =   345
         Left            =   6720
         TabIndex        =   59
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label lblPalletPerUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletPerUnit"
         Height          =   345
         Left            =   4080
         TabIndex        =   58
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblPalletTo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletTo"
         Height          =   345
         Left            =   2520
         TabIndex        =   57
         Top             =   3360
         Width           =   615
      End
      Begin Threed.SSCommand cmdPalletNo 
         Height          =   405
         Left            =   5445
         TabIndex        =   56
         Top             =   4320
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx2.frx":0EFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   0
         TabIndex        =   55
         Top             =   9120
         Width           =   1695
      End
      Begin VB.Label lblTimePackEnd 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTimePackEnd"
         Height          =   375
         Left            =   3960
         TabIndex        =   44
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label lblTimePackBegin 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTimePackBegin"
         Height          =   375
         Left            =   -240
         TabIndex        =   46
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label lblPackDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   -240
         TabIndex        =   53
         Top             =   6240
         Width           =   1905
      End
      Begin VB.Label lblRestAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRestAmount"
         Height          =   375
         Left            =   3600
         TabIndex        =   52
         Top             =   5280
         Width           =   1665
      End
      Begin VB.Label lblLoseAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLoseAmount"
         Height          =   345
         Left            =   4320
         TabIndex        =   51
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label lblGoodAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblGoodAmount"
         Height          =   345
         Left            =   0
         TabIndex        =   50
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label lblBinNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBinNo"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   2760
         Width           =   1545
      End
      Begin VB.Label lblPalletFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPalletFrom"
         Height          =   345
         Left            =   0
         TabIndex        =   48
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblLockNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLockNo"
         Height          =   345
         Left            =   3600
         TabIndex        =   47
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblProductType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProductType"
         Height          =   315
         Left            =   2760
         TabIndex        =   45
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblWeightPerPack 
         Alignment       =   1  'Right Justify
         Caption         =   "lblWeightPerPack"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   4320
         Width           =   1545
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPackAmount"
         Height          =   375
         Left            =   -240
         TabIndex        =   42
         Top             =   5760
         Width           =   1905
      End
      Begin VB.Label lblStdAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStdAmount"
         Height          =   375
         Left            =   3360
         TabIndex        =   41
         Top             =   5760
         Width           =   1905
      End
      Begin VB.Label lblSerialNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSerialNo"
         Height          =   375
         Left            =   60
         TabIndex        =   40
         Top             =   8640
         Width           =   1665
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "lblRef"
         Height          =   345
         Left            =   3960
         TabIndex        =   39
         Top             =   8640
         Width           =   975
      End
      Begin VB.Label lblPlace 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPlace"
         Height          =   315
         Left            =   0
         TabIndex        =   38
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "lblAmount"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   5280
         Width           =   1545
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblType"
         Height          =   315
         Left            =   270
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProduct"
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   675
         Width           =   1455
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   7320
         TabIndex        =   30
         Top             =   8400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJobOutputEx2.frx":1218
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7320
         TabIndex        =   31
         Top             =   9000
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditJobOutputEx2"
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
Public ID As Long
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
Private m_LocationsRest As Collection
Private m_LocationsLose As Collection
Private m_Units As Collection
Private Lt As cLot
Public m_JobInOut As Collection
'Private tempPD As Collection
Private IWD As CInventoryWHDoc
Private LWH As CLotItemWH

Private OldLotNo As Long
Private NewLotNo As Long
Private LotDocId As Long
Private OldWeightPerPack As Long
Private OldTxAmount As Double

Private OldHeadPackNo As Long
Private NewHeadPackNo As Long

Private OldPartItemId As Long
Private NewPartItemId As Long

Public DOCUMENT_TYPE As Long
Public typeInput As Long 'เป็นการตัดแตก หรือ ถ่ายถุงธรรมดา

Public TempPDEdit As Collection
Public m_CollLotExUse As Collection
Public m_CollPalletInLot As Collection
Public TempCLotDoc As CLotDoc

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

Private Sub cboHead_Change()
m_HasModify = True
End Sub

Private Sub cboHead_Click()
   If cboHead.ListIndex = 1 Then 'head pack 1
      Call LoadLocation(cboBinNo, Nothing, 2, , -2, , , "BIN")
   ElseIf cboHead.ListIndex = 2 Then 'head pack 2
      Call LoadLocation(cboBinNo, Nothing, 2, , -3, , , "BIN")
   End If
   m_HasModify = True
   NewHeadPackNo = cboHead.ItemData(Minus2Zero(cboHead.ListIndex))

   If ShowMode = SHOW_ADD And typeInput = -1 Then 'ถ้าเป็นการเพิ่มแบบปรกติ
      Call RefreshPallet
   ElseIf ShowMode = SHOW_ADD And (typeInput = 1 Or typeInput = 3 Or typeInput = 5) Then
     Call RefreshPallet
     Call RefreshPallet2
   End If

End Sub

Private Sub cboHead_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
KeyAscii = 0
End Sub

Private Sub cboLotNo_Change()
   m_HasModify = True
End Sub
Function EditPallet(PartItemID As Long, LotId As Long, HeadPackNo As Long, DocumentType As Long)
Dim m_CollPallet As Collection
Dim TempI As Collection
Dim TempE As Collection
Dim PD_I As CPalletDoc
Dim PD_E As CPalletDoc

Dim PrevKey As Long
Dim I As Long

      Set m_CollPallet = New Collection
      Set TempI = New Collection
      Set TempE = New Collection
      
   Call LoadPalletDoc(Nothing, TempI, LotId, 7, , 9, "I", IWD.INVENTORY_WH_DOC_ID, , , OldPartItemId, DocumentType)
   For Each PD_I In TempI
      PD_I.PALLET_DOC_NO_OLD = PD_I.PALLET_DOC_NO
      PD_I.PALLET_DOC_NO = Val(txtPalletFrom.Text) + I
      I = I + 1
      
      If I = 1 Then
         txtPalletFrom.Text = PD_I.PALLET_DOC_NO
         txtPalletTo.Text = PD_I.PALLET_DOC_NO
      Else
         txtPalletTo.Text = PD_I.PALLET_DOC_NO
      End If

      Call TempPDEdit.add(PD_I, str(PD_I.PALLET_DOC_NO_OLD) & "-" & str(PD_I.LOT_DOC_ID) & "-" & PD_I.TX_TYPE)
      If PrevKey <> PD_I.LOT_DOC_ID Then
         PrevKey = PD_I.LOT_DOC_ID
         Call LoadPalletDoc(Nothing, TempE, LotId, 1, , 9, "E", , , , OldPartItemId, , PD_I.LOT_DOC_ID)
         For Each PD_E In TempE
            PD_E.PALLET_DOC_NO_OLD = PD_E.PALLET_DOC_NO
            PD_E.PART_ITEM_ID = NewPartItemId
             Call TempPDEdit.add(PD_E, str(PD_E.PALLET_DOC_NO) & "-" & str(PD_E.LOT_DOC_ID_REF) & "-" & PD_E.TX_TYPE)
         Next PD_E
      End If
   Next PD_I
   
   For Each PD_I In TempI
    Set PD_E = GetObject("CPalletDoc", TempPDEdit, str(PD_I.PALLET_DOC_NO_OLD) & "-" & str(PD_I.LOT_DOC_ID) & "-" & "E", False)
      If Not (PD_E Is Nothing) Then
              PD_E.PALLET_DOC_NO = PD_I.PALLET_DOC_NO
      End If
   Next PD_I
End Function
Function RefreshPallet()
Dim TempI As Collection
'Dim TempC As Collection
Dim LotId As Long
Dim HeadPackNo As Long
Dim PartItemID As Long
Dim LocationID As Long
Dim PD As CPalletDoc
   m_HasModify = True
   m_HasModify2 = False
   Set TempI = New Collection
   LotId = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
   NewLotNo = LotId
   HeadPackNo = cboHead.ItemData(Minus2Zero(cboHead.ListIndex))
   LocationID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   PartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   Call LoadPalletDoc(Nothing, TempI, LotId, 4, , 6, "I", , , HeadPackNo, PartItemID, , , LocationID)
   Set PD = GetObject("CPalletDoc", TempI, Trim(str(NewLotNo) & "-" & "I" & "-" & HeadPackNo), False) 'Set PD = GetObject("CPalletDoc", TempI, Trim(str(LotId) & "-" & "I" & "-" & HeadPackNo), False)
    If Not (PD Is Nothing) Then
      If (Not NewLotNo = OldLotNo) Or (Not NewHeadPackNo = OldHeadPackNo) Or (Not NewPartItemId = OldPartItemId) Then
         If ShowMode = SHOW_EDIT Then
            If LWH.FULL_PALLET_FROM > Val(PD.PALLET_DOC_NO) Then 'ถ้า palllet เริ่มต้นของ lot เดิม มากกว่า palllet เริ่มต้นของ lot ใหม่ ก็ให้สามารถแก้ไขข้อมูลได้โดยไม่ต้องเปลี่ยนแปลงชื่อ pallet เดิม
               m_HasModify = True
               Exit Function
            End If
         End If
      
         txtPalletFrom.Enabled = False
         txtPalletFrom2.Enabled = False
         txtPalletFrom.Text = Val(PD.PALLET_DOC_NO) + 1
         txtPalletTo.Text = Val(txtPalletFrom.Text) + 1
         txtPalletTo.Enabled = True
         txtPalletPerUnit2.Enabled = True
         If ShowMode = SHOW_EDIT And OldLotNo > 0 Then
               MsgBox "มีข้อมูลพาเลทที่มีชื่อเหมือนกันในล๊อตนี้อยู่แล้วระบบจะทำการสร้างพาเลทให้ใหม่"
               If NewLotNo > 0 Then
                  Call EditPallet(PartItemID, NewLotNo, NewHeadPackNo, 14)
               End If
         End If
      End If
   Else
      If ShowMode = SHOW_ADD Then
         txtPalletFrom.Text = 1
         txtPalletTo.Text = Val(txtPalletFrom.Text) + 1
         txtPalletTo.Enabled = True
         txtPalletPerUnit2.Enabled = True
      End If
   End If
   If (NewLotNo = OldLotNo) And (NewHeadPackNo = OldHeadPackNo) And (NewPartItemId = OldPartItemId) Then
      If ShowMode = SHOW_EDIT Then  'ถ้ากลับมาเลือก pallet เดิม หรือเป็นการย้ายล๊อตแล้ว pallet ไม่ซ้ำกัน
            txtPalletFrom.Enabled = False
            txtPalletFrom2.Enabled = False
            txtPalletFrom.Text = LWH.FULL_PALLET_FROM
            txtPalletTo.Text = LWH.FULL_PALLET_TO
            txtPalletPerUnit.Text = LWH.FULL_UNIT_PER_PALLET
            txtPalletFrom2.Text = LWH.SCRAP_PALLET
            txtPalletPerUnit2.Text = LWH.SCRAP_UNIT_PER_PALLET
            txtPalletTo.Enabled = False
            txtPalletPerUnit2.Enabled = False
         End If
      End If
   If LotId = 0 Then
      txtPalletFrom.Text = ""
      txtPalletTo.Text = ""
      txtPalletFrom2.Text = ""
      txtPalletPerUnit.Text = ""
      txtPalletPerUnit2.Text = ""
  End If
'  txtGoodAmount.Text = (Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)
  txtGoodAmount.Text = ((Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)) + Val(txtPalletPerUnit2.Text)
  
End Function
Function RefreshPallet2()
      Dim PpU As Double
      Dim Tiwd As CInventoryWHDoc
      Dim Tltd As CLotItemWH
      Dim Td1 As Double
      Dim Td2 As Double
      Set Tiwd = TempCollection5.Item(1)
      Set Tltd = Tiwd.C_LotItemsWH.Item(1)
         
      If Val(txtWeightPerPack.Text) <> Tltd.WEIGHT_PER_PACK Then   'กรณีถ้าน้ำหนัก Input และ Putput ไม่เท่ากัน
         PpU = getFormat(uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)), Val(txtWeightPerPack.Text))
         Td1 = MyDiff(Tltd.TX_AMOUNT * Tltd.WEIGHT_PER_PACK, Val(txtWeightPerPack.Text)) \ PpU 'หารเอาจำนวนเต็ม
         Td2 = MyDiff(Tltd.TX_AMOUNT * Tltd.WEIGHT_PER_PACK, Val(txtWeightPerPack.Text)) / PpU
         
         If Td1 = 0 Then 'ถ้า Pallet แรก ไม่เต็ม let
             txtPalletTo.Text = Val(txtPalletFrom.Text)
             txtPalletPerUnit.Text = MyDiff(Tltd.TX_AMOUNT * Tltd.WEIGHT_PER_PACK, Val(txtWeightPerPack.Text))
         Else
            txtPalletTo.Text = Td1 + Val(txtPalletFrom.Text) - 1
            txtPalletPerUnit.Text = PpU
         End If
         
         If Td2 >= 1 And Td2 > Td1 Then
            txtPalletFrom2.Text = Val(txtPalletTo.Text) + 1
            txtPalletPerUnit2.Text = MyDiff(Tltd.TX_AMOUNT * Tltd.WEIGHT_PER_PACK, Val(txtWeightPerPack.Text)) - (PpU * Td1)
         Else
             txtPalletFrom2.Text = Val(txtPalletTo.Text) + 1
            txtPalletPerUnit2.Text = "0"
         End If
        txtGoodAmount.Text = MyDiff(Tltd.TX_AMOUNT * Tltd.WEIGHT_PER_PACK, Val(txtWeightPerPack.Text))
      End If
End Function
Private Sub cboLotNo_Click()
  If Not VerifyCombo(lblPlace, uctlPlaceLookup.MyCombo, False) Then
      Exit Sub
   End If
   
m_HasModify = True
NewLotNo = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
   
    If ShowMode = SHOW_ADD And typeInput = -1 Then 'ถ้าเป็นการเพิ่มแบบปรกติ
      Call RefreshPallet
   ElseIf ShowMode = SHOW_ADD And (typeInput = 1 Or typeInput = 3 Or typeInput = 5) Then
     Call RefreshPallet
     Call RefreshPallet2
   End If
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

      If Not CheckUniqueNs(LOT_UNIQUE, Lt.LOT_NO, ID) Then
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

Private Sub cmdEditUnitPerPallet_Click()
   txtPalletPerUnit.Enabled = True
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
   Call InitNormalLabel(lblProductType, MapText("ชนิดสินค้า"))
   Call InitNormalLabel(lblLotNo, MapText("Lot การผลิต"))
   Call InitNormalLabel(lblLotNo2, MapText("Lot"))
   Call InitNormalLabel(lblGoodAmount, MapText("จำนวนในโกดัง"))
   Call InitNormalLabel(lblLoseAmount, MapText("ของเสีย"))
   Call InitNormalLabel(lblWeightPerPack, MapText("ขนาดถุง"))
   Call InitNormalLabel(lblAmount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblStdAmount, MapText("จำนวนมาตรฐาน"))
   Call InitNormalLabel(lblRestAmount, MapText("จำนวนเศษ (กก.)"))
   Call InitNormalLabel(lblBinNo, MapText("เบอร์ถัง"))
   Call InitNormalLabel(lblPalletFrom, MapText("วางพาเลทที่"))
   Call InitNormalLabel(lblPalletTo, MapText("ถึง"))
   Call InitNormalLabel(lblPalletPerUnit, MapText("จำนวน/พาเลท"))
   Call InitNormalLabel(lblUnit1, MapText("ถุง"))
   Call InitNormalLabel(lblPalletFrom2, MapText("เศษที่เหลือ"))
   Call InitNormalLabel(lblPalletPerUnit2, MapText("จำนวน/พาเลท"))
   Call InitNormalLabel(lblUnit2, MapText("ถุง"))
   Call InitNormalLabel(lblHead, MapText("หัวแพ็ค"))
   
   
   Call InitNormalLabel(lblLockNo, MapText("ล๊อค"))
'   Call InitNormalLabel(lblStartDate, MapText("วันที่ผลิต"))
   Call InitNormalLabel(lblPackDate, MapText("วันที่บรรจุ"))
   Call InitNormalLabel(lblTimePackBegin, MapText("เวลาเริ่มบรรจุ"))
   Call InitNormalLabel(lblTimePackEnd, MapText("เวลาหลังบรรจุ"))
   Call InitNormalLabel(lblPlace, MapText("ที่เก็บอาหารดี"))
   Call InitNormalLabel(lblPlaceRest, MapText("ที่เก็บอาหารเศษ"))
   Call InitNormalLabel(lblPlaceLose, MapText("ที่เก็บอาหารเสีย"))
   Call InitNormalLabel(lblPackAmount, MapText("จำนวนบรรจุ (ถุง)"))
   Call InitNormalLabel(lblSerialNo, MapText("ซีเรียล"))
   Call InitNormalLabel(lblRef, MapText("หมายเลขอ้างอิง"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))

   Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
   lblLotNo2.Enabled = False
   txtLotNoNew.Enabled = False
   Call txtPalletFrom.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPalletTo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPalletPerUnit.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPalletFrom2.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPalletPerUnit2.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call txtGoodAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtGoodAmount.Enabled = False
   Call txtLoseAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtWeightPerPack.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtWeightPerPack.Enabled = False
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtAmount.Enabled = False
   Call txtStdAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtStdAmount.Enabled = False
   Call txtRestAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPackAmount.Enabled = False
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtSerialNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call uctlProductLookup.MyTextBox.SetKeySearch("PART_NO")
   
  If ShowMode = SHOW_EDIT Then
      txtPalletFrom.Enabled = False
      txtPalletTo.Enabled = False
      txtPalletPerUnit.Enabled = False
      txtPalletFrom2.Enabled = False
      txtPalletPerUnit2.Enabled = False
      uctlPlaceLookup.Enabled = False
'      uctlPlaceLookupRest.Enabled = False
'      uctlPlaceLookupLose.Enabled = False
      cmdEditUnitPerPallet.Enabled = False
   Else
      txtPalletFrom.Enabled = False
      txtPalletTo.Enabled = True
      txtPalletPerUnit.Enabled = False
      txtPalletFrom2.Enabled = False
      txtPalletPerUnit2.Enabled = True
      uctlPlaceLookup.Enabled = False
      uctlPlaceLookupRest.Enabled = False
      uctlPlaceLookupLose.Enabled = False
  End If
  

   Call InitCombo(cboLotNo)
   Call InitCombo(cboBinNo)
   Call InitCombo(cboLockNo)
   Call InitCombo(cboHead)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPalletNo.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditUnitPerPallet.Picture = LoadPicture(glbParameterObj.NormalButton1)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdEditUnitPerPallet, MapText("E"))
   Call InitMainButton(cmdPalletNo, MapText("รายละเอียดพาเลท"))
   Call InitMainButton(cmdAuto2, MapText("A"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

Dim Ma As CJobInput
Dim LTD As CLotDoc
Dim PD As CPalletDoc

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
      
         If TempCollection Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      
         
         Set Ma = TempCollection.Item(ID)
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
      
         Set IWD = TempCollection2.Item(ID)
         If IWD.C_LotItemsWH Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
               
         If (IWD.C_LotItemsWH.Count = 0) Or (IWD.C_LotItemsWH Is Nothing) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
          If ID = 2 Then
            ID = 1
          End If
         Set LWH = IWD.C_LotItemsWH.Item(ID)
         
         If CountItem(LWH.C_LotDoc) <= 0 Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If LWH.C_LotDoc.Item(ID) Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         
         Set LTD = LWH.C_LotDoc.Item(ID)

         If LWH.BL_START_DATE > 0 Then
            uctlStartDate.ShowDate = LWH.BL_START_DATE
         Else
            uctlStartDate.ShowDate = LWH.START_DATE
            m_HasModify3 = True
         End If
         uctlProductTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlProductTypeLookup.MyCombo, LWH.PRODUCT_TYPE_ID)
         cboLotNo.ListIndex = IDToListIndex(cboLotNo, LTD.LOT_ID)
         'ทำไว้เพื่อเช็คว่า มีการเปลี่ยน lot ใหม่หรือไม่ เพื่อจะได้ให้โปรแกรมตัดสินใจได้ว่า จะสร้าง pallet ใหม่ หรือไม่
         OldLotNo = LTD.LOT_ID
         NewLotNo = LTD.LOT_ID
         
         OldHeadPackNo = LWH.HEAD_PACK_NO
         NewHeadPackNo = LWH.HEAD_PACK_NO
         
         OldPartItemId = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
         NewPartItemId = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
         ''''''''''''''''''''''''''''''
         cboLockNo.ListIndex = IDToListIndex(cboLockNo, LWH.LOCK_NO)
         cboHead.ListIndex = IDToListIndex(cboHead, LWH.HEAD_PACK_NO)
         
         If cboHead.ListIndex = 1 Then 'head pack 1
            Call LoadLocation(cboBinNo, Nothing, 2, , -2, , , "BIN")
         ElseIf cboHead.ListIndex = 2 Then 'head pack 2
            Call LoadLocation(cboBinNo, Nothing, 2, , -3, , , "BIN")
         End If
         
         cboBinNo.ListIndex = IDToListIndex(cboBinNo, LWH.BIN_NO)
         txtGoodAmount.Text = LWH.GOOD_AMOUNT
         txtLoseAmount.Text = LWH.LOSE_AMOUNT
         txtWeightPerPack.Text = LWH.WEIGHT_PER_PACK
         OldWeightPerPack = LWH.WEIGHT_PER_PACK
         txtRestAmount.Text = LWH.REST_AMOUNT
         txtPackAmount.Text = LWH.PACK_AMOUNT
         uctlPackDate.ShowDate = LWH.PACK_DATE
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, LWH.LOCATION_ID)
         uctlPlaceLookupRest.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupRest.MyCombo, LWH.LOCATION_REST_ID)
         uctlPlaceLookupLose.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupLose.MyCombo, LWH.LOCATION_LOSE_ID)
         uctlPackDate.ShowDate = LWH.PACK_DATE
         txtTimePackBegin.HR = HOUR(LWH.TIME_PACK_BEGIN)
         txtTimePackBegin.MI = Minute(LWH.TIME_PACK_BEGIN)
         txtTimePackEnd.HR = HOUR(LWH.TIME_PACK_END)
         txtTimePackEnd.MI = Minute(LWH.TIME_PACK_END)
         txtNote.Text = LWH.NOTE
         'เป็นการ Gen ตอนสร้าง เพราะฉะนั้นตอนแก้ไขก็ไม่ต้องให้แสดง
         txtPalletFrom.Text = LWH.FULL_PALLET_FROM
         txtPalletTo.Text = LWH.FULL_PALLET_TO
         txtPalletPerUnit.Text = LWH.FULL_UNIT_PER_PALLET
         txtPalletFrom2.Text = LWH.SCRAP_PALLET
         txtPalletPerUnit2.Text = LWH.SCRAP_UNIT_PER_PALLET
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         
         
      'ตรวจสอบว่า lot นี้ได้มีการปรับยอดไปแล้วหรือไม่
      If LTD.BALANCE_FLAG = "Y" Then
        ' MsgBox "Lot  นี้มีการปรับยอดแล้ว ไม่สามารถเปลี่ยนแปลงแก้ไขข้อมูลได้ "
         m_HasModify2 = True
         Call EnableForm(Me, True)
         Exit Sub
      End If
      End If
   Else
      If typeInput = 1 Or typeInput = 3 Or typeInput = 5 Then
         Set IWD = TempCollection5.Item(1)
         If IWD.C_LotItemsWH Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
               
         If (IWD.C_LotItemsWH.Count = 0) Or (IWD.C_LotItemsWH Is Nothing) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
          ID = 1
         Set LWH = IWD.C_LotItemsWH.Item(ID)
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, LWH.PART_ITEM_ID)
         If CountItem(LWH.C_LotDoc) <= 0 Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If LWH.C_LotDoc.Item(ID) Is Nothing Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
       
         
         Set LTD = LWH.C_LotDoc.Item(ID)
         Set TempCLotDoc = LTD
         
         If LWH.BL_START_DATE > 0 Then
            uctlStartDate.ShowDate = LWH.BL_START_DATE
         Else
            uctlStartDate.ShowDate = LWH.START_DATE
         End If
         uctlProductTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlProductTypeLookup.MyCombo, LWH.PRODUCT_TYPE_ID)
         If typeInput = 1 Then
            uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 109)
            uctlPlaceLookupRest.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupRest.MyCombo, 117)
            uctlPlaceLookupLose.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupLose.MyCombo, 104)
         ElseIf typeInput = 3 Then
            uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 78)
            uctlPlaceLookupRest.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupRest.MyCombo, 117)
            uctlPlaceLookupLose.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupLose.MyCombo, 104)
         ElseIf typeInput = 5 Then
            uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 109)
            uctlPlaceLookupRest.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupRest.MyCombo, 117)
            uctlPlaceLookupLose.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupLose.MyCombo, 104)
         End If
         
         cboLotNo.ListIndex = IDToListIndex(cboLotNo, LTD.LOT_ID)
         cboLockNo.ListIndex = IDToListIndex(cboLockNo, LWH.LOCK_NO)
         cboHead.ListIndex = IDToListIndex(cboHead, LWH.HEAD_PACK_NO)
         
         If cboHead.ListIndex = 1 Then 'head pack 1
            Call LoadLocation(cboBinNo, Nothing, 2, , -2, , , "BIN")
         ElseIf cboHead.ListIndex = 2 Then 'head pack 2
            Call LoadLocation(cboBinNo, Nothing, 2, , -3, , , "BIN")
         End If
         
         cboBinNo.ListIndex = IDToListIndex(cboBinNo, LWH.BIN_NO)
        
         txtWeightPerPack.Text = LWH.WEIGHT_PER_PACK
         OldWeightPerPack = LWH.WEIGHT_PER_PACK
         
         
         Dim R As Long
         Dim IsOne As Boolean
         Dim SumPD As Double
         IsOne = True
         For Each PD In LTD.C_PalletDoc
         SumPD = SumPD + PD.CAPACITY_AMOUNT
            If IsOne Then
               IsOne = False
               txtPalletFrom.Text = txtPalletFrom.Text 'PD.PALLET_DOC_NO
               txtPalletTo.Text = txtPalletFrom.Text
               txtPalletPerUnit.Text = PD.CAPACITY_AMOUNT
               
               PD.PALLET_DOC_NO = Format(txtPalletFrom.Text, "00")
            Else
                 txtPalletTo.Text = Val(txtPalletTo.Text) + 1
                 txtPalletPerUnit.Text = SumPD
                 
                 PD.PALLET_DOC_NO = Format(txtPalletTo.Text, "00")
            End If
             
         Next PD

         txtLoseAmount.Text = "0"
         txtRestAmount.Text = "0"
         txtPackAmount.Text = SumPD
         txtGoodAmount.Text = SumPD
         
         txtPalletFrom.Enabled = False
         txtPalletTo.Enabled = False
         txtPalletPerUnit.Enabled = False
         txtPalletFrom2.Enabled = False
         txtPalletPerUnit2.Enabled = False
      Else
         uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 109)
         uctlPlaceLookupRest.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupRest.MyCombo, 117)
         uctlPlaceLookupLose.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookupLose.MyCombo, 104)
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

   If Not VerifyCombo(lblHead, cboHead, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblProductType, uctlProductTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblLotNo, cboLotNo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblBinNo, cboBinNo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblLockNo, cboLockNo, False) Then
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
   
   If m_HasModify2 Then
      MsgBox "Lot  นี้มีการปรับยอดแล้ว ไม่สามารถเปลี่ยนแปลงแก้ไขข้อมูลได้ "
      Exit Function
   End If
   
 If ShowMode = SHOW_ADD Then
   If Not VerifyTextControl(lblPalletFrom, txtPalletFrom, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPalletTo, txtPalletTo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPalletPerUnit, txtPalletPerUnit, False) Then
      Exit Function
   End If
    If Len(txtPalletFrom2.Text) > 0 Then
      ElseIf Len(txtPalletFrom2.Text) = 0 Then
         txtPalletFrom2.Text = ""
         txtPalletPerUnit2.Text = ""
      End If
   
      If Val(txtPalletTo.Text) < Val(txtPalletFrom.Text) Then
         Call MsgBox("กรุณากรอกข้อมูลให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
            If txtPalletTo.Enabled Then
               txtPalletTo.SetFocus
            End If
         Exit Function
      End If
   
      If Val(txtPalletFrom2.Text) <= Val(txtPalletTo.Text) And Val(txtPalletFrom2.Text) > 0 Then
         If Val(txtPalletPerUnit2.Text) > 0 Then
            txtPalletFrom2.Text = Val(txtPalletTo.Text) + 1
         Else
            txtPalletFrom2.Text = ""
            txtPalletPerUnit2.Text = ""
         End If
      End If
   
      If Val(txtPalletFrom2.Text) >= Val(txtPalletTo.Text) Then
         txtPalletFrom2.Text = Val(txtPalletTo.Text) + 1
      End If
   End If

   
   If Val(txtAmount.Text) = 0 Then
        Call MsgBox(lblAmount.Caption & "ต้องไม่เท่ากับ 0 ", vbOKOnly, PROJECT_NAME)
        Exit Function
   End If

  If (txtTimePackBegin.HR) = "24" Then
        txtTimePackBegin.HR = "00"
        txtTimePackBegin.MI = "00"
  End If

  If (txtTimePackEnd.HR) = "24" Then
        txtTimePackEnd.HR = "23"
        txtTimePackEnd.MI = "59"
  End If
  
  If typeInput = 1 Or typeInput = 3 Or typeInput = 5 Then  'ถ้าเป็นการตัดแตกให้ เข้าSave Auto
         m_HasModify = True
  End If
        
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ma As CJobInput
   If ShowMode = SHOW_ADD Then
      Set Ma = New CJobInput
   Else
      Set Ma = TempCollection.Item(ID)
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

   If ShowMode = SHOW_ADD Then
      Set IWD = New CInventoryWHDoc
      Set LWH = New CLotItemWH
   Else
      Set IWD = TempCollection2.Item(ID)
      Set LWH = IWD.C_LotItemsWH.Item(ID)
   End If
   
   LWH.PART_NO = uctlProductLookup.MyTextBox.Text
   LWH.LOT_NO = cboLotNo.Text
   LWH.BL_START_DATE = uctlStartDate.ShowDate
   LWH.PART_ITEM_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   LWH.PRODUCT_TYPE_ID = uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex))
   LWH.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
   LWH.LOCK_NO = cboLockNo.ItemData(Minus2Zero(cboLockNo.ListIndex))
   LWH.HEAD_PACK_NO = cboHead.ItemData(Minus2Zero(cboHead.ListIndex))
   LWH.GOOD_AMOUNT = Val(txtGoodAmount.Text)
   LWH.LOSE_AMOUNT = Val(txtLoseAmount.Text)
   LWH.WEIGHT_PER_PACK = Val(txtWeightPerPack.Text)
   LWH.PACK_AMOUNT = Val(txtPackAmount.Text)
   LWH.REST_AMOUNT = Val(txtRestAmount.Text)
   LWH.PACK_DATE = uctlPackDate.ShowDate
   LWH.TIME_PACK_BEGIN = txtTimePackBegin.HR & ":" & txtTimePackBegin.MI
   LWH.TIME_PACK_END = txtTimePackEnd.HR & ":" & txtTimePackEnd.MI
   LWH.TX_AMOUNT = Val(txtAmount.Text)
   LWH.NOTE = txtNote.Text
   LWH.LOCATION_ID = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   LWH.LOCATION_NO = uctlPlaceLookup.MyTextBox.Text
   LWH.LOCATION_NAME = uctlPlaceLookup.MyCombo.Text
   LWH.LOCATION_REST_ID = uctlPlaceLookupRest.MyCombo.ItemData(Minus2Zero(uctlPlaceLookupRest.MyCombo.ListIndex))
   LWH.LOCATION_LOSE_ID = uctlPlaceLookupLose.MyCombo.ItemData(Minus2Zero(uctlPlaceLookupLose.MyCombo.ListIndex))
   LWH.TX_TYPE = "I" 'รับเข้า
   
   LWH.FULL_PALLET_FROM = Val(txtPalletFrom.Text)
   LWH.FULL_PALLET_TO = Val(txtPalletTo.Text)
   LWH.FULL_UNIT_PER_PALLET = Val(txtPalletPerUnit.Text)
   LWH.SCRAP_PALLET = Val(txtPalletFrom2.Text)
   LWH.SCRAP_UNIT_PER_PALLET = Val(txtPalletPerUnit2.Text)
   
   LWH.InputRebagToBagType = typeInput
   
   Dim LTD As CLotDoc
   Dim PD As CPalletDoc
   Set LTD = New CLotDoc
   If ShowMode = SHOW_ADD And typeInput = -1 Then
      'Gen pallet
      For I = Val(txtPalletFrom.Text) To Val(txtPalletTo.Text)
         Set PD = New CPalletDoc
         PD.Flag = "A"
         PD.PALLET_DOC_NO = Format(I, "00")
         PD.CAPACITY_AMOUNT = Val(txtPalletPerUnit.Text)
         PD.TX_TYPE = "I"
         PD.AddEditMode = ShowMode
         Call LTD.C_PalletDoc.add(PD)
         Set PD = Nothing
      Next I
      If Val(txtPalletFrom2.Text) > 0 And Val(txtPalletPerUnit2.Text) > 0 Then
         Set PD = New CPalletDoc
         PD.Flag = "A"
         PD.PALLET_DOC_NO = Format(txtPalletFrom2.Text, "00")
         PD.CAPACITY_AMOUNT = Val(txtPalletPerUnit2.Text)
         PD.TX_TYPE = "I"
         PD.AddEditMode = ShowMode
         Call LTD.C_PalletDoc.add(PD)
         Set PD = Nothing
      End If
      
      LTD.Flag = "A"
      LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
      LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
      LTD.AddEditMode = ShowMode
      Call LWH.C_LotDoc.add(LTD)
      LWH.Flag = "A"
      Call IWD.C_LotItemsWH.add(LWH)
      Call TempCollection2.add(IWD)
   ElseIf ShowMode = SHOW_ADD And (typeInput = 1 Or typeInput = 3 Or typeInput = 5) Then
      Dim Tiwd As CInventoryWHDoc
      Dim Tltd As CLotItemWH
      Set Tiwd = TempCollection5.Item(1)
      Set Tltd = Tiwd.C_LotItemsWH.Item(1)
      If Val(txtWeightPerPack.Text) <> Tltd.WEIGHT_PER_PACK Then   'ถ้าน้ำหนัก Input และ Putput ไม่เท่ากัน
              'Gen pallet
            For I = Val(txtPalletFrom.Text) To Val(txtPalletTo.Text)
               Set PD = New CPalletDoc
               PD.Flag = "A"
               PD.PALLET_DOC_NO = Format(I, "00")
               PD.CAPACITY_AMOUNT = Val(txtPalletPerUnit.Text)
               PD.TX_TYPE = "I"
               PD.AddEditMode = ShowMode
               Call LTD.C_PalletDoc.add(PD)
               Set PD = Nothing
            Next I
            If Val(txtPalletFrom2.Text) > 0 And Val(txtPalletPerUnit2.Text) > 0 Then
               Set PD = New CPalletDoc
               PD.Flag = "A"
               PD.PALLET_DOC_NO = Format(txtPalletFrom2.Text, "00")
               PD.CAPACITY_AMOUNT = Val(txtPalletPerUnit2.Text)
               PD.TX_TYPE = "I"
               PD.AddEditMode = ShowMode
               Call LTD.C_PalletDoc.add(PD)
               Set PD = Nothing
            End If
      Else 'ถ้าน้ำหนัก Input และ Putput เท่ากัน
          'Gen pallet
         For Each PD In TempCLotDoc.C_PalletDoc
            PD.Flag = "A"
            PD.AddEditMode = ShowMode
            PD.TX_TYPE = "I"
            Call LTD.C_PalletDoc.add(PD)
         Next PD
      End If
  
      LTD.Flag = "A"
      LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
      LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
      LTD.AddEditMode = ShowMode
      Call LWH.C_LotDoc.add(LTD)
      LWH.Flag = "A"
      Call IWD.C_LotItemsWH.add(LWH)
      Call TempCollection2.add(IWD)
      
   Else
      If LWH.Flag <> "A" Then
         LWH.Flag = "E"
      End If
      
       Set LTD = LWH.C_LotDoc.Item(ID)
         If LTD.Flag <> "A" Then
            LTD.Flag = "E"
         End If
         LTD.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
         LTD.AddEditMode = ShowMode
         
     If NewLotNo <> OldLotNo Then
        LTD.LOT_ID = cboLotNo.ItemData(Minus2Zero(cboLotNo.ListIndex))
       Dim m_collLD As Collection
       Dim temp_LD As CLotDoc
       Set m_collLD = New Collection
       Call LoadLDByLotDocIdRef(Nothing, m_collLD, LTD.LOT_DOC_ID)
       For Each temp_LD In m_collLD
         temp_LD.LOT_ID = LTD.LOT_ID
         temp_LD.AddEditMode = SHOW_EDIT
         temp_LD.Flag = "E"
         Call LWH.C_LotDoc.add(temp_LD)
       Next temp_LD
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
Private Sub cmdPalletNo_Click()
Dim Col_PD As Collection
      If ShowMode = SHOW_EDIT Then
         If LWH Is Nothing Then
            Exit Sub
         End If
      
         If Not (LWH.C_LotDoc.Item(ID).C_PalletDoc Is Nothing) Then
            Set frmLocation.TempLotItemWh = LWH
            Set frmLocation.TempCollection = LWH.C_LotDoc.Item(ID).C_PalletDoc
            Set frmLocation.m_CollLotExUse = m_CollLotExUse
            Set frmLocation.m_CollPalletInLot = m_CollPalletInLot
            frmLocation.LotNo = cboLotNo.Text
            frmLocation.DOCUMENT_TYPE = DOCUMENT_TYPE
            frmLocation.BALANCE_FLAG = LWH.C_LotDoc.Item(ID).BALANCE_FLAG
            Load frmLocation
            frmLocation.Show 1
         End If
      End If

      OKClick = frmLocation.OKClick
      If OKClick Then
         m_HasModify = True
         Call CheckMaxMinNamePallet(LWH.C_LotDoc.Item(ID).C_PalletDoc)
         txtGoodAmount.Text = SumPalletAmount(LWH.C_LotDoc.Item(ID).C_PalletDoc)
      End If


      Unload frmLocation
      Set frmLocation = Nothing
End Sub

Function SumPalletAmount(Cl As Collection) As Long
   Dim PD As CPalletDoc
   For Each PD In Cl
      If Not PD.Flag = "D" Then
         SumPalletAmount = SumPalletAmount + PD.CAPACITY_AMOUNT
      End If
   Next PD
End Function
Function CheckMaxMinNamePallet(Cl As Collection)
   Dim PD As CPalletDoc
   Dim MIN As Double
   Dim MAX As Double
   Dim CMin As Double
   Dim CMax As Double
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
   If CMax <> CMin Then
      txtPalletFrom.Text = MIN
      txtPalletTo.Text = MAX - 1
      txtPalletFrom2.Text = MAX
      txtPalletPerUnit.Text = CMin
      txtPalletPerUnit2.Text = CMax
   ElseIf CMax = CMin Then
      txtPalletFrom.Text = MIN
      txtPalletTo.Text = MAX
      txtPalletFrom2.Text = ""
      txtPalletPerUnit.Text = CMin
      txtPalletPerUnit2.Text = ""
   End If
End Function


Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      uctlPartTypeLookup.Enabled = False
    
      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2)
      Set uctlPlaceLookup.MyCollection = m_Locations
      
      Call LoadLocation(uctlPlaceLookupRest.MyCombo, m_LocationsRest, 2)
      Set uctlPlaceLookupRest.MyCollection = m_LocationsRest
      
      Call LoadLocation(uctlPlaceLookupLose.MyCombo, m_LocationsLose, 2)
      Set uctlPlaceLookupLose.MyCollection = m_LocationsLose
      
     Call LoadMaster(uctlProductTypeLookup.MyCombo, m_Units, PRODUCT_TYPE)
     Set uctlProductTypeLookup.MyCollection = m_Units
     
     Call LoadLocation(cboHead, Nothing, 2, , , , , "HEAD")
     Call LoadLotFromLot(cboLotNo, Nothing, , -1, -1, , , 1, , 2)
     Call LoadLocation(cboLockNo, Nothing, 2, , , , , "LOCK")
     
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         
         If m_CollLotExUse.Count > 0 Then '   ถ้า lotdocid นี้ มีการเบิกจ่ายไปแล้ว จะไม่ให้สามารถแก้ไขเลขที่ lot ได้แล้ว
             cboLotNo.Enabled = False
             cmdAuto2.Enabled = False
         End If
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlPackDate.ShowDate = Now
         txtTimePackBegin.HR = HOUR(Now)
         txtTimePackBegin.MI = Minute(Now)
         txtTimePackEnd.HR = HOUR(Now)
         txtTimePackEnd.MI = Minute(Now)
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, PartType)
         uctlStartDate.ShowDate = StartJob
         Call QueryData(False)
      End If
      
      If Not m_HasModify3 Then
         m_HasModify = False
      End If
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
   Set m_LocationsRest = New Collection
   Set m_LocationsLose = New Collection
   Set m_Units = New Collection
   Set m_CollLotItemWh = New Collection
   Set TempCollection3 = New Collection
   Set m_JobInOut = New Collection
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
   Set m_LocationsRest = Nothing
   Set m_LocationsLose = Nothing
   Set m_Units = Nothing
   Set m_CollLotItemWh = Nothing
   Set TempCollection3 = Nothing
   Set Lt = Nothing
   Set m_JobInOut = Nothing
End Sub

Private Sub txtGoodAmount_Change()
On Error Resume Next
   m_HasModify = True
   txtAmount.Text = Val(txtGoodAmount.Text) * Val(txtWeightPerPack.Text) 'น้ำหนักรวม=ดี * น้ำหนัก
   txtPackAmount.Text = Val(txtGoodAmount.Text) + Val(txtLoseAmount.Text) 'ยอดแพ็ค=ดี+เสีย
End Sub

Private Sub txtGoodAmount_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtLoseAmount_Change()
'On Error Resume Next
   m_HasModify = True
End Sub

Private Sub txtLotNo_Change()
   m_HasModify = True
End Sub

Private Sub txtLoseAmount_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
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

Private Sub txtPalletFrom_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtPalletFrom2_Change()
   m_HasModify = True
End Sub

Private Sub txtPalletFrom2_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtPalletPerUnit_Change()
   m_HasModify = True
'   If ShowMode = SHOW_ADD Then
'       txtGoodAmount.Text = (Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)
       txtGoodAmount.Text = ((Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)) + Val(txtPalletPerUnit2.Text)
'   End If
End Sub

Private Sub txtPalletPerUnit_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtPalletPerUnit2_Change()
   m_HasModify = True
   txtGoodAmount.Text = ((Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)) + Val(txtPalletPerUnit2.Text)
End Sub

Private Sub txtPalletPerUnit2_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtPalletTo_Change()
If Val(txtPalletTo.Text) < Val(txtPalletFrom.Text) Then
   Exit Sub
End If
   m_HasModify = True
   txtPalletFrom2.Text = Val(txtPalletTo.Text) + 1
   If (Val(txtPalletFrom.Text) = Val(txtPalletTo.Text)) And ShowMode = SHOW_ADD Then
      txtPalletPerUnit.Enabled = True
   Else
      txtPalletPerUnit.Enabled = False
   End If
   txtGoodAmount.Text = ((Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)) + Val(txtPalletPerUnit2.Text)
End Sub

Private Sub txtPalletTo_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtRestAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtRestAmount_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
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
   txtAmount.Text = Val(txtGoodAmount.Text) * Val(txtWeightPerPack.Text)
   txtPalletPerUnit.Text = getFormat(uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)), Val(txtWeightPerPack.Text))
   txtGoodAmount.Text = ((Val(txtPalletTo.Text) - Val(txtPalletFrom.Text) + 1) * Val(txtPalletPerUnit.Text)) + Val(txtPalletPerUnit2.Text)
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

Private Sub txtWeightPerPack_KeyPress(KeyAscii As Integer)
 KeyAscii = CheckIntAscii(KeyAscii)
End Sub

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

Private Sub uctlPlaceLookupLose_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookupRest_Change()
   m_HasModify = True
End Sub

Private Sub uctlProductLookup_Change()
Dim Pi As CPartItem
   PartItemID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   
   If PartItemID > 0 Then
      Set Pi = GetPartItem(m_PartItems, Trim(str(PartItemID)))
      txtWeightPerPack.Text = Pi.WEIGHT_PER_PACK
   End If

    Call LoadLotIdByPartItem(cboLotNo, m_CollLotItemWh, , , , , PartItemID, 5, 1, 1, "I", TempCollection3, 1, Lt)
   m_HasModify = True

   NewPartItemId = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   
   If Not (IWD Is Nothing) Then
       Call RefreshPallet
   End If
   
'   txtPalletPerUnit.Text = MyDiff(txtAmount.Text, Pi.WEIGHT_PER_PACK)
  
End Sub

Private Sub uctlProductTypeLookup_Change()
   m_HasModify = True
   txtPalletPerUnit.Text = getFormat(uctlProductTypeLookup.MyCombo.ItemData(Minus2Zero(uctlProductTypeLookup.MyCombo.ListIndex)), Val(txtWeightPerPack.Text))
End Sub

