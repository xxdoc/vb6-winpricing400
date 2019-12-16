VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditLocation 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmAddEditLocation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   4842
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtPalletNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtCapacityAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1470
         Width           =   1575
         _ExtentX        =   13361
         _ExtentY        =   767
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3480
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1080
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblCapacityAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblPalletNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Public TempCollection As Collection
Public m_CollPalletInLot As Collection
Public TempLotItemWh As CLotItemWH
Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public Area As Long
Public TempValue As String
Public tempPallet As String
Public FlagNotEditOver As Boolean
Public DocumentTypeInput As Long
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim X As Double
Dim PD  As CPalletDoc
   If Flag Then
     If Area = 2 Then
         txtPalletNo.Text = 99
     Else
       Set PD = TempCollection.Item(ID)
       txtPalletNo.Text = PD.PALLET_DOC_NO
       txtCapacityAmount.Text = PD.CAPACITY_AMOUNT
       PD.TEMP_PALLET_CAP_LAST = PD.CAPACITY_AMOUNT
      End If
   End If
   Call EnableForm(Me, True)
End Sub
Function CheckPalletNoUniqueInCol(Cl As Collection, Key As String) As Boolean
   Dim PD As CPalletDoc
   CheckPalletNoUniqueInCol = False
   For Each PD In Cl
      If Not PD.Flag = "D" Then
         If Key = PD.PALLET_DOC_NO Then
            CheckPalletNoUniqueInCol = True
            Exit Function
         End If
      End If
   Next PD
End Function
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim isError As Boolean
Dim PALLET_NO As String
Dim LIW As CLotItemWH
Dim PD As CPalletDoc

If FlagNotEditOver Then
   If Val(txtCapacityAmount.Text) > Val(TempValue) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถแก้ไขยอดใหม่่ให้มากกว่ายอดเดิมได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
End If
'TempLotItemWh
'TempLotItemWh.HEAD_PACK_NO

PALLET_NO = Format(Trim(txtPalletNo.Text), "00")
' If ShowMode = SHOW_ADD Then
'          If CheckPalletNoUniqueInCol(TempCollection, PALLET_NO) Then
'           If tempPallet <> Format(Trim(txtPalletNo.Text), "00") Then 'แก้ไขตัวเอง
'               glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & PALLET_NO & " " & MapText("ที่เพิ่งเพิ่มเข้าใหม่โดยยังไม่ได้บันทึก อยู่ใน Lot นี้แล้ว")
'               glbErrorLog.ShowUserError
'               Exit Function
'            End If
'
'         End If
'   End If
   
   If Area <> 3 Then
      Set PD = GetObject("CPalletDoc", m_CollPalletInLot, Trim(PALLET_NO) & "-" & Trim(str(TempLotItemWh.HEAD_PACK_NO)), False)
      If Not PD Is Nothing Then
        isError = True
         If ShowMode = SHOW_EDIT Then
              If tempPallet = Format(Trim(txtPalletNo.Text), "00") Then 'แก้ไขตัวเอง
                 isError = False
              End If
          End If
         If isError Then
            If PD.Flag = "A" Then
              If PD.PALLET_DOC_NO_OLD <> PALLET_NO Then 'ถ้าไม่ใช่ กลับมาแก้ ให้เป็น pallet เดิม ที่เข้าไปอยู่ใน collection แล้ว
                  glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & PALLET_NO & " " & MapText("ที่เพิ่งเพิ่มเข้าใหม่โดยยังไม่ได้บันทึก อยู่ใน Lot นี้แล้ว")
                  glbErrorLog.ShowUserError
                  Exit Function
               ElseIf PD.PALLET_DOC_NO_OLD <> tempPallet Then 'ถ้าไม่ใช่ กลับมาแก้ ให้เป็น pallet เดิม ที่เข้าไปอยู่ใน collection แล้ว
                  glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & PALLET_NO & " " & MapText("ที่เพิ่งเพิ่มเข้าใหม่โดยยังไม่ได้บันทึก อยู่ใน Lot นี้แล้ว")
                  glbErrorLog.ShowUserError
                  Exit Function
              End If
            Else
               glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลพาเลท เลทที่ ") & " " & PALLET_NO & " " & MapText("อยู่ใน ล๊อต " & PD.LOT_NO & " แล้ว")
               glbErrorLog.ShowUserError
               Exit Function
            End If
         End If
       Else
         
         If ShowMode = SHOW_EDIT Then
          glbErrorLog.LocalErrorMsg = MapText("ต้องลบพาเลท " & tempPallet & " แล้วเพิ่มพาเลท " & Format(Trim(txtPalletNo.Text), "00") & " เข้าไปใหม่")
           glbErrorLog.ShowUserError
           Exit Function
         End If
'        End If
      End If
   End If
   
   If Not VerifyTextControl(lblPalletNo, txtPalletNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCapacityAmount, txtCapacityAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   
   If ShowMode = SHOW_ADD Then
      Set PD = New CPalletDoc
   Else
      Set PD = TempCollection.Item(ID)
   End If
   
   Dim t_PD As CPalletDoc
   Dim LotDocId As Long
   If Area = 2 Then 'Dummy pallet
     For Each t_PD In TempCollection
      t_PD.AddEditMode = SHOW_EDIT
      t_PD.Flag = "E"
      LotDocId = t_PD.LOT_DOC_ID
     Next t_PD
      Set PD = New CPalletDoc
      ShowMode = SHOW_ADD
      PD.AddEditMode = SHOW_ADD
      PD.PALLET_DOC_NO = "ยอดยกมา"
      PD.CAPACITY_AMOUNT = Val(txtCapacityAmount.Text)
      PD.LOT_DOC_ID = LotDocId
      PD.TX_TYPE = "I"
   Else
       If PD.Flag <> "A" Then
         PD.AddEditMode = ShowMode
         PD.TX_TYPE = "I"
         PD.PALLET_DOC_NO = PALLET_NO
         PD.CAPACITY_AMOUNT = Val(txtCapacityAmount.Text)
      Else
         PD.PALLET_DOC_NO = PALLET_NO
         PD.CAPACITY_AMOUNT = Val(txtCapacityAmount.Text)
      End If
   End If
   
   If ShowMode = SHOW_ADD Then
      PD.Flag = "A"
      Call TempCollection.add(PD)
      
      PD.PALLET_DOC_NO_OLD = PD.PALLET_DOC_NO
      Call m_CollPalletInLot.add(PD, Trim(PALLET_NO) & "-" & Trim(str(TempLotItemWh.HEAD_PACK_NO)))
   Else
     If PD.Flag <> "A" Then
         PD.Flag = "E"
     End If
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
         tempPallet = txtPalletNo.Text
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         txtPalletNo.Enabled = True
      End If
      TempValue = txtCapacityAmount.Text
      
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
   
   Call InitNormalLabel(lblPalletNo, MapText("ชื่อพาเลท"))
   Call InitNormalLabel(lblCapacityAmount, MapText("จำนวนบรรจุ"))
   If DocumentTypeInput = 14 Or DocumentTypeInput = 15 Or DocumentTypeInput = 17 Then
      Call InitNormalLabel(Label1, MapText("ถุง"))
   ElseIf DocumentTypeInput = 13 Or DocumentTypeInput = 16 Or DocumentTypeInput = 18 Then
      Call InitNormalLabel(Label1, MapText("ก.ก."))
   End If
         
   Call txtPalletNo.SetTextLenType(TEXT_STRING, glbSetting.PALLET_NO)
   If Area = 2 Then
     txtPalletNo.Enabled = False
   ElseIf Area = 3 Then
    Call txtPalletNo.SetTextLenType(TEXT_STRING, glbSetting.PORT_TYPE)
    txtPalletNo.Enabled = False
   End If
   Call txtCapacityAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
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
   Set m_Rs = New ADODB.Recordset
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   '''Debug.Print ColIndex & " " & NewColWidth
End Sub
Private Sub txtCapacityAmount_Change()
 m_HasModify = True
End Sub

Private Sub txtCapacityAmount_KeyPress(KeyAscii As Integer)
KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtPalletNo_Change()
 m_HasModify = True
End Sub

Private Sub txtPalletNo_KeyPress(KeyAscii As Integer)
   KeyAscii = CheckIntAscii(KeyAscii)
End Sub
