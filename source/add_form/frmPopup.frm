VERSION 5.00
Begin VB.Form frmPopup 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   510
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
Label1.Caption = UserGroup
End Sub
Private Sub Form_Resize()
    Dim popup_L As Single
    Dim popup_T As Single
    Dim popup_wid As Single
    Dim popup_hgt As Single
    popup_L = Screen.Width / 2
    popup_T = 0
    popup_wid = 4000
    popup_hgt = 465
    Label1.Move 0, 0, popup_wid, popup_hgt
    frmPopup.Move popup_L, popup_T, popup_wid, popup_hgt
End Sub

