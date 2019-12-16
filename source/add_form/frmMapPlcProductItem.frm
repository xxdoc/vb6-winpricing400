VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmMapPlcProductItem 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMapPlcProductItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1665
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   2937
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   3705
         TabIndex        =   0
         Top             =   270
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6000
         TabIndex        =   5
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4200
         TabIndex        =   1
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMapPlcProductItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2115
         TabIndex        =   4
         Top             =   210
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmMapPlcProductItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE

Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean

Public mPartItemColl As Collection
Public PartItem As CPartItem
Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblPart, MapText("สินค้า/วัตถุดิบ"))
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Set mPartItemColl = Nothing
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
   
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   
   Dim Pi As CPartItem
   Set Pi = GetObject("CPartItem", mPartItemColl, Trim(str(uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)))), False)
   If Not Pi Is Nothing Then
      PartItem.PART_ITEM_ID = Pi.PART_ITEM_ID
      PartItem.PART_NO = Pi.PART_NO
      PartItem.DEFAULT_LOCATION = Pi.DEFAULT_LOCATION
      
      SaveData = True
   End If
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Dim I As Long
      Dim Pi As CPartItem
      
      uctlPartLookup.MyCombo.Clear
      I = 0
      uctlPartLookup.MyCombo.AddItem ("")

      For Each Pi In mPartItemColl
         I = I + 1
         uctlPartLookup.MyCombo.AddItem (Pi.PART_DESC & "  (" & Pi.PART_NO & ")")
         uctlPartLookup.MyCombo.ItemData(I) = Pi.PART_ITEM_ID
      Next Pi
      Set uctlPartLookup.MyCollection = mPartItemColl
      
      m_HasModify = False
      
      Call EnableForm(Me, True)
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
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub
