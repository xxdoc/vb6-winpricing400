VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCurrencyEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "q"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmAddEditCurrencyEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   7858
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   8
         Top             =   0
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlCurrencyDate 
         Height          =   405
         Left            =   2130
         TabIndex        =   0
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox TXT3 
         Height          =   435
         Left            =   2130
         TabIndex        =   3
         Top             =   2360
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox TXT1 
         Height          =   435
         Left            =   2130
         TabIndex        =   1
         Top             =   1440
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox TXT2 
         Height          =   435
         Left            =   2130
         TabIndex        =   2
         Top             =   1900
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox TXT4 
         Height          =   435
         Left            =   2130
         TabIndex        =   4
         Top             =   2800
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin VB.Label lblCountry4 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry4"
         Height          =   315
         Left            =   420
         TabIndex        =   13
         Top             =   2890
         Width           =   1635
      End
      Begin VB.Label lblCurrencyDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrencyDate"
         Height          =   435
         Left            =   960
         TabIndex        =   12
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label lblCountry3 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry3"
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Top             =   2500
         Width           =   1635
      End
      Begin VB.Label lblCountry1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry1"
         Height          =   315
         Left            =   420
         TabIndex        =   10
         Top             =   1530
         Width           =   1635
      End
      Begin VB.Label lblCountry2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry2"
         Height          =   435
         Left            =   420
         TabIndex        =   9
         Top             =   2040
         Width           =   1635
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3990
         TabIndex        =   6
         Top             =   3600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2340
         TabIndex        =   5
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCurrencyEx.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCurrencyEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Currency As CCurrencyEx

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Currency.CURRENCY_EX_ID = ID
      m_Currency.QueryFlag = 1
      If Not glbDaily.QueryCurrencyEx(m_Currency, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_Currency.PopulateFromRS(1, m_Rs)
      
      TXT1.Text = m_Currency.US
      TXT2.Text = m_Currency.EURO
      TXT3.Text = m_Currency.YEN
      TXT4.Text = m_Currency.SS
      uctlCurrencyDate.ShowDate = m_Currency.EXCHANGE_DATE
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblCountry1, TXT1, True) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblCountry2, TXT2, True) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblCountry3, TXT3, True) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblCountry4, TXT4, True) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblCurrencyDate, uctlCurrencyDate, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Currency.CURRENCY_EX_ID = ID
   m_Currency.AddEditMode = ShowMode
   m_Currency.US = Val(TXT1.Text)
   m_Currency.EURO = Val(TXT2.Text)
   m_Currency.YEN = Val(TXT3.Text)
   m_Currency.SS = Val(TXT4.Text)

   m_Currency.EXCHANGE_DATE = uctlCurrencyDate.ShowDate
   
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCurrencyEx(m_Currency, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlCurrencyDate.ShowDate = Now
         ID = 0
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCurrencyDate, MapText("วันที่"))
   Call InitNormalLabel(lblCountry1, MapText("US$"))
   Call InitNormalLabel(lblCountry2, MapText("EURO"))
   Call InitNormalLabel(lblCountry3, MapText("YEN"))
   Call InitNormalLabel(lblCountry4, MapText("S$"))
   
   Call TXT4.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call TXT3.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call TXT2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call TXT1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Currency = New CCurrencyEx
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub TXT1_Change()
m_HasModify = True
End Sub

Private Sub TXT2_Change()
m_HasModify = True
End Sub

Private Sub TXT3_Change()
m_HasModify = True
End Sub

Private Sub TXT4_Change()
m_HasModify = True
End Sub

Private Sub uctlCurrencyDate_HasChange()
m_HasModify = True
End Sub
