VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmLockDate 
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   12210
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11853
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   975
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1720
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSFrame2"
         Begin prjFarmManagement.uctlDate uctlFromInventoryDate 
            Height          =   375
            Left            =   1920
            TabIndex        =   17
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
         End
         Begin prjFarmManagement.uctlDate uctlToInventoryDate 
            Height          =   375
            Left            =   7440
            TabIndex        =   18
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
         End
         Begin VB.Label lblFromInventoryDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label lblToInventoryDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5760
            TabIndex        =   7
            Top             =   480
            Width           =   1620
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12405
         _ExtentX        =   21881
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   975
         Left            =   360
         TabIndex        =   9
         Top             =   3000
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1720
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSFrame2"
         Begin prjFarmManagement.uctlDate uctlFromInvoiceDate 
            Height          =   375
            Left            =   1920
            TabIndex        =   19
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
         End
         Begin prjFarmManagement.uctlDate uctlToInvoiceDate 
            Height          =   375
            Left            =   7440
            TabIndex        =   20
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
         End
         Begin VB.Label lblToInvoiceDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5760
            TabIndex        =   11
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label lblFromInvoiceDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1620
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   975
         Left            =   360
         TabIndex        =   12
         Top             =   4200
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1720
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSFrame2"
         Begin prjFarmManagement.uctlDate uctlFromReceiptDate 
            Height          =   375
            Left            =   1920
            TabIndex        =   21
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
         End
         Begin prjFarmManagement.uctlDate uctlToReceiptDate 
            Height          =   375
            Left            =   7440
            TabIndex        =   22
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
         End
         Begin VB.Label lblToReceiptDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5760
            TabIndex        =   14
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label lblFromReceiptDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   1620
         End
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   375
         Left            =   7800
         TabIndex        =   16
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6120
         TabIndex        =   5
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   1620
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6090
         TabIndex        =   1
         Top             =   5700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4440
         TabIndex        =   0
         Top             =   5700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmLockDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_LockDate As CLockDate

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_LockDate.LOCK_DATE_ID = -1
      m_LockDate.LOCK_TYPE = 1 'Global Lock
      If Not glbDaily.QueryLockDate(m_LockDate, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If m_Rs.RecordCount > 0 Then
      Call m_LockDate.PopulateFromRS(1, m_Rs)
      
      uctlFromDate.ShowDate = m_LockDate.FROM_DATE
      uctlToDate.ShowDate = m_LockDate.TO_DATE
      
      uctlFromInventoryDate.ShowDate = m_LockDate.FROM_INVENTORY_DATE
      uctlToInventoryDate.ShowDate = m_LockDate.TO_INVENTORY_DATE
      uctlFromInvoiceDate.ShowDate = m_LockDate.FROM_INVOICE_DATE
      uctlToInvoiceDate.ShowDate = m_LockDate.TO_INVOICE_DATE
      uctlFromReceiptDate.ShowDate = m_LockDate.FROM_RECEIPT_DATE
      uctlToReceiptDate.ShowDate = m_LockDate.TO_RECEIPT_DATE
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
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
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_LockDate = New CLockDate
   
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   Call InitNormalLabel(lblFromDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(lblToDate, MapText("�֧�ѹ���"))
   
   Call InitNormalLabel(lblFromInventoryDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(lblToInventoryDate, MapText("�֧�ѹ���"))
   Call InitNormalLabel(lblFromInvoiceDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(lblToInvoiceDate, MapText("�֧�ѹ���"))
   Call InitNormalLabel(lblFromReceiptDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(lblToReceiptDate, MapText("�֧�ѹ���"))
   
   SSFrame2.Caption = "��ͤ�͡��� STOCK ��� ��ü�Ե (�����ŷ���������㹪�ǧ���ҷ���˹��������ö����¹�ŧ�͡���������)"
   SSFrame3.Caption = "��ͤ�͡��� ���觫��� INVOICE ����稢��ʴ ��Ѻ�׹�Թ��� (�����ŷ���������㹪�ǧ���ҷ���˹��������ö����¹�ŧ�͡���������)"
   SSFrame4.Caption = "��ͤ�͡��� �ѭ������ (�����ŷ���������㹪�ǧ���ҷ���˹��������ö����¹�ŧ�͡���������)"
   
   uctlFromDate.Enable = False
   uctlToDate.Enable = False
   SSFrame3.Enabled = False
   SSFrame4.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
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
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim I As Long
Dim ID As Long
Dim Cd As CConfigDoc
   
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_LockDate.ShowMode = SHOW_EDIT
   m_LockDate.FROM_DATE = uctlFromDate.ShowDate
   m_LockDate.TO_DATE = uctlToDate.ShowDate
   
   m_LockDate.FROM_INVENTORY_DATE = uctlFromInventoryDate.ShowDate
   m_LockDate.TO_INVENTORY_DATE = uctlToInventoryDate.ShowDate
   m_LockDate.FROM_INVOICE_DATE = uctlFromInvoiceDate.ShowDate
   m_LockDate.TO_INVOICE_DATE = uctlToInvoiceDate.ShowDate
   m_LockDate.FROM_RECEIPT_DATE = uctlFromReceiptDate.ShowDate
   m_LockDate.TO_RECEIPT_DATE = uctlToReceiptDate.ShowDate
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditLockDate(m_LockDate, IsOK, True, glbErrorLog) Then
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
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cmdOK_Click()
   If cmdOK.Enabled = False Then
      Exit Sub
   End If
   Call SaveData
   Call getLockDate
   m_HasModify = False
   Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_LockDate = Nothing
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromInventoryDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromInvoiceDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromReceiptDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToInventoryDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToInvoiceDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToReceiptDate_HasChange()
   m_HasModify = True
End Sub
