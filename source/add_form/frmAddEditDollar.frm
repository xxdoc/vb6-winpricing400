VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditDollar 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmAddEditDollar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   5953
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCountry1 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Width           =   2670
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   8025
         _ExtentX        =   14155
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
      Begin prjFarmManagement.uctlTextBox txtCoefficient 
         Height          =   435
         Left            =   2130
         TabIndex        =   2
         Top             =   1920
         Width           =   1755
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin VB.Label lblCurrencyDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrencyDate"
         Height          =   435
         Left            =   960
         TabIndex        =   9
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label lblCoefficient 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCoefficient"
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label lblCountry1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCountry1"
         Height          =   315
         Left            =   210
         TabIndex        =   7
         Top             =   1530
         Width           =   1875
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3990
         TabIndex        =   4
         Top             =   2520
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2340
         TabIndex        =   3
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDollar.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditDollar"
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
Public Dollar1 As Double
Public dollar2 As Double
Public COUNTRY_CURRENCY1 As Long
Public COUNTRY_CURRENCY2 As Long
Public dollarID As Long
Public COEF As Double
Public Date1 As Date


Private Sub cboCountry1_Change()
m_HasModify = True
End Sub

Private Sub cboCountry1_Click()
m_HasModify = True
End Sub


Private Sub cmdOK_Click()
 If Not QueryData(True) Then
        glbErrorLog.LocalErrorMsg = "ไม่พบข้อมูลในฐานข้อมูล"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   OKClick = True
   Unload Me
End Sub

Private Function QueryData(Flag As Boolean) As Boolean
Dim IsOK As Boolean
Dim itemcount As Long

 QueryData = False
   If Not VerifyTextControl(lblCoefficient, txtCoefficient, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCountry1, cboCountry1, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblCurrencyDate, uctlCurrencyDate, False) Then
      Exit Function
   End If
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Currency.EXCHANGE_DATE = uctlCurrencyDate.ShowDate
      m_Currency.QueryFlag = 1
      If Not glbDaily.QueryCurrencyEx(m_Currency, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If itemcount > 0 Then
      Call m_Currency.PopulateFromRS(1, m_Rs)
       COUNTRY_CURRENCY1 = cboCountry1.ItemData(Minus2Zero(cboCountry1.ListIndex))
      If COUNTRY_CURRENCY1 = 1 Then
         Dollar1 = Val(txtCoefficient.Text) * m_Currency.US
         COEF = m_Currency.US
         ElseIf COUNTRY_CURRENCY1 = 2 Then
         Dollar1 = Val(txtCoefficient.Text) * m_Currency.EURO
         COEF = m_Currency.EURO
         ElseIf COUNTRY_CURRENCY1 = 3 Then
         Dollar1 = Val(txtCoefficient.Text) * m_Currency.YEN
         COEF = m_Currency.YEN
         ElseIf COUNTRY_CURRENCY1 = 4 Then
         Dollar1 = Val(txtCoefficient.Text) * m_Currency.SS
         COEF = m_Currency.SS
         End If
         
       dollar2 = Val(txtCoefficient.Text)
       dollarID = m_Currency.CURRENCY_EX_ID
       
       
       Date1 = m_Currency.EXCHANGE_DATE
   Else
   Call EnableForm(Me, True)
   Exit Function
      End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Function
   End If
 QueryData = True
   Call EnableForm(Me, True)
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      uctlCurrencyDate.ShowDate = Now
       Call LoadMoneyFamilyEx(cboCountry1)
        
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
        
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
   Call InitNormalLabel(lblCountry1, MapText("สกุลเงินต้นทาง"))
   Call InitNormalLabel(lblCoefficient, MapText("จำนวนเงิน"))
   Call InitCombo(cboCountry1)

   Call txtCoefficient.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)

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

Private Sub txtCoefficient_Change()
m_HasModify = True
End Sub


Private Sub uctlCurrencyDate_HasChange()
m_HasModify = True
End Sub
