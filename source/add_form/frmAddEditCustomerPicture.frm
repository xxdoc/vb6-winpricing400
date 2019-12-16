VERSION 5.00
Object = "{4BD5A3A1-7FFE-11D4-A13A-004005FA6275}#1.0#0"; "ImagXpr6.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditCustomerPicture 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAddEditCustomerPicture.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8565
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15108
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin IMAGXPR6LibCtl.ImagXpress ImagXpress1 
         Height          =   6975
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   12303
         ErrStr          =   "F4NRO2IK2AP-ER3063PXEP"
         ErrCode         =   1054018678
         ErrInfo         =   -653612348
         Persistence     =   -1  'True
         _cx             =   1
         _cy             =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ScrollBars      =   3
         ScrollBarLargeChangeH=   10
         ScrollBarSmallChangeH=   1
         OLEDropMode     =   0
         ScrollBarLargeChangeV=   10
         ScrollBarSmallChangeV=   1
         DisplayProgressive=   -1  'True
         SaveTIFByteOrder=   0
         LoadRotated     =   0
         FTPUserName     =   ""
         FTPPassword     =   ""
         ProxyServer     =   ""
      End
      Begin prjFarmManagement.uctlTextBox txtPath 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   840
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   75
         TabIndex        =   2
         Top             =   6600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   11280
         TabIndex        =   1
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerPicture.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblPath 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   900
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   75
         TabIndex        =   4
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   75
         TabIndex        =   3
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerPicture.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCustomerPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public PictureType As PICTURE_TYPE

Public TempCollection As Collection
Public ParentForm As Form
Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
Dim id As Long
Dim MyName As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.JPG)|*..jpg;*.JPG;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   ShowMode = SHOW_ADD
      
   txtPath.Text = dlgAdd.FileName
   m_HasModify = True

End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(id, TempCollection)
      If id = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      id = NewID
   ElseIf ShowMode = SHOW_ADD Then
   
   End If
   Call QueryData(True)
   Call ParentForm.RefreshGrid
   
   Call cmdFileName.SetFocus
   m_HasModify = False
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
          Dim Di As CCustomerPicture
         
          Set Di = TempCollection.Item(id)
         
         
         
          txtPath.Text = Di.GetFieldValue("CUSTOMER_PICTURE_PATH")
      
         ImagXpress1.ZoomToFit ZOOMFIT_BEST
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblPath, txtPath, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Di As CCustomerPicture
   If ShowMode = SHOW_ADD Then
      Set Di = New CCustomerPicture

      Di.Flag = "A"
      Call TempCollection.add(Di)
   Else
      Set Di = TempCollection.Item(id)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If
   
   Call Di.SetFieldValue("CUSTOMER_PICTURE_PATH", txtPath.Text)
   Call Di.SetFieldValue("CUSTOMER_PICTURE_TYPE", PictureType)
   
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
         id = 0
      End If
      
      m_HasModify = False
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
      Call cmdNext_Click
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
   
   Call InitNormalLabel(lblPath, MapText("ที่อยู่รูปภาพ"))
   
   Call txtPath.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPath.Enabled = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdFileName, MapText(".B."))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
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
   
   ImagXpress1.Antialias = AA_ScaleToGray
   ImagXpress1.AutoSize = ISIZE_CropImage
   ImagXpress1.PictureEnabled = False
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub ImagXpress1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ขนาดดีที่สุด", "-", "ตามสูง", "-", "ตามกว้าง", "-", "25%", "-", "50%", "-", "75%", "-", "100%", "-", "125", "-", "150", "-", "175", "-", "200")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
      
      If lMenuChosen = 1 Then
         ' Note: This will be the equivalent of ZOOMFIT_HEIGHT ot ZOOMFIT_WIDTH
         ' depending on which one fits the entire image within the control.
         ImagXpress1.ZoomToFit ZOOMFIT_BEST
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & ImagXpress1.IPZoomF
      ElseIf lMenuChosen = 3 Then
         ImagXpress1.ZoomToFit ZOOMFIT_HEIGHT
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & ImagXpress1.IPZoomF
      ElseIf lMenuChosen = 5 Then
         ImagXpress1.ZoomToFit ZOOMFIT_WIDTH
         'Label1.Caption = "Zoom Factor:" & Chr$(10) & ImagXpress1.IPZoomF
      ElseIf lMenuChosen = 7 Then
         ImagXpress1.Zoom 0.25
      ElseIf lMenuChosen = 9 Then
         ImagXpress1.Zoom 0.5
      ElseIf lMenuChosen = 11 Then
         ImagXpress1.Zoom 0.75
      ElseIf lMenuChosen = 13 Then
         ImagXpress1.Zoom 1
      ElseIf lMenuChosen = 15 Then
         ImagXpress1.Zoom 1.25
      ElseIf lMenuChosen = 17 Then
         ImagXpress1.Zoom 1.5
      ElseIf lMenuChosen = 19 Then
         ImagXpress1.Zoom 1.75
      ElseIf lMenuChosen = 21 Then
         ImagXpress1.Zoom 2
      End If
      
      Call txtPath_Change
   End If
End Sub

Private Sub txtPath_Change()
On Error GoTo Errorhanderor
   m_HasModify = True
   
   
   If ShowMode = SHOW_ADD Then
      ImagXpress1.CancelLoad = True
      ImagXpress1.DeleteSaveBuffer
      
      ImagXpress1.FileName = txtPath.Text
      
   Else
      ImagXpress1.FileName = glbParameterObj.MapDrivePicture & txtPath.Text
   End If
   
   Exit Sub
Errorhanderor:
   glbErrorLog.LocalErrorMsg = "Error หารูปไม่พบ"
   glbErrorLog.ShowUserError
End Sub
