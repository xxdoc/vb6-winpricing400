VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditLotNo 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditLotNo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6075
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   10716
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlStartDate 
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
      End
      Begin VB.ComboBox cboBinNo 
         Height          =   510
         Left            =   1800
         TabIndex        =   1
         Top             =   3000
         Width           =   2085
      End
      Begin VB.ComboBox cboHead 
         Height          =   510
         Left            =   8760
         TabIndex        =   6
         Top             =   810
         Visible         =   0   'False
         Width           =   800
      End
      Begin prjFarmManagement.uctlTextLookup uctlProductTypeLookup 
         Height          =   435
         Left            =   10800
         TabIndex        =   7
         Top             =   810
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTime uctlTime1 
         Height          =   375
         Left            =   60500
         TabIndex        =   10
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextBox txtLotNoNew 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   2520
         Width           =   3855
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaCode 
         Height          =   435
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   3855
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaName 
         Height          =   435
         Left            =   1800
         TabIndex        =   14
         Top             =   1560
         Width           =   3855
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNo 
         Height          =   435
         Left            =   2760
         TabIndex        =   17
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   2760
         TabIndex        =   19
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchFrom 
         Height          =   435
         Left            =   1800
         TabIndex        =   23
         Top             =   3600
         Width           =   2055
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchTo 
         Height          =   435
         Left            =   1800
         TabIndex        =   24
         Top             =   4080
         Width           =   2055
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBatchDetail 
         Height          =   435
         Left            =   1800
         TabIndex        =   28
         Top             =   4560
         Width           =   3855
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin VB.Label lblBatchDetail 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchDetail"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   4560
         Width           =   1515
      End
      Begin VB.Label lblBatchTo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchEnd"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   4080
         Width           =   1515
      End
      Begin VB.Label lblBatchFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchStart"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Label lblStartDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblStartDate"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBatchNo"
         Height          =   315
         Left            =   720
         TabIndex        =   20
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lblNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNo"
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   120
         Width           =   1995
      End
      Begin VB.Label lblBinNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBinNo"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label lblFormulaName 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaName"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblFormulaCode 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaCode"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1515
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   360
         TabIndex        =   2
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLotNo.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblLotNo2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLotNo2"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label lblProductType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProductType"
         Height          =   315
         Left            =   9720
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2040
         TabIndex        =   3
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLotNo.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3720
         TabIndex        =   4
         Top             =   5280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditLotNo"
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
Private m_Rs As ADODB.Recordset
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public COMMIT_FLAG As String
Public TempCollection As Collection
Public PartItemID As Long
Public StartDate As Date
Private CBatch As CBacthing
Public ParentForm As Form
Public SplitFlag As Boolean
Private Sub cboBinNo_Change()
   m_HasModify = True
End Sub

Private Sub cboBinNo_Click()
 m_HasModify = True
End Sub

Private Sub cboLotNo_Change()
   m_HasModify = True
End Sub

Private Sub cboLotNo_Click()
m_HasModify = True
End Sub

Private Sub cboLotNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub cboBinNo_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cmdExit_Click()
'   If Not ConfirmExit(m_HasModify) Then
'      Exit Sub
'      End If
   
   OKClick = False
   Unload Me
End Sub

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
   
   'lblFormulaCode
   Call InitNormalLabel(lblNo, MapText("ลำดับที่"))
   Call InitNormalLabel(lblBatchNo, MapText("หมายเลขการผลิต"))
   Call InitNormalLabel(lblFormulaCode, MapText("รหัสสูตร"))
   Call InitNormalLabel(lblFormulaName, MapText("ชื่อสูตร"))
   Call InitNormalLabel(lblStartDate, MapText("วันที่ผลิต"))
   Call InitNormalLabel(lblLotNo2, MapText("เลขล็อต"))
   Call InitNormalLabel(lblBinNo, MapText("เบอร์ถัง"))
   Call InitNormalLabel(lblBatchFrom, MapText("จากแบต"))
   Call InitNormalLabel(lblBatchTo, MapText("ถึงแบต"))
   Call InitNormalLabel(lblBatchDetail, MapText("รายละเอียดแบต"))

  Call txtNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
  Call txtBatchNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFormulaCode.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFormulaName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
   Call txtBatchFrom.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
   Call txtBatchTo.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
   Call txtBatchDetail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   txtNo.Enabled = False
   txtBatchNo.Enabled = False
  txtFormulaCode.Enabled = False
  txtFormulaName.Enabled = False
  txtBatchDetail.Enabled = False
  
  Call InitCombo(cboBinNo)
   
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long
   If Not (TempCollection Is Nothing) Then
   Set CBatch = TempCollection.Item(ID)
      If Not (CBatch Is Nothing) Then
           If Not (CBatch.FormulaCode = "10541" Or CBatch.FormulaCode = "10101") Then  'ถ้าไม่ใช่รำล้างไลน์ หรือ ข้าวโพดล้างไลน์
            If Not SaveData Then
               Exit Sub
            End If
           End If
      End If
   End If

   txtLotNoNew.Text = ""
   cboBinNo.ListIndex = -1


   NewID = GetNextID(ID, TempCollection)
   If ID = NewID Then
      glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
      glbErrorLog.ShowUserError
      
      Call ParentForm.RefreshGrid
      Exit Sub
   End If
   
   ID = NewID
   
If Not (TempCollection Is Nothing) Then
   Set CBatch = TempCollection.Item(ID)
   If Not (CBatch Is Nothing) Then
      If CBatch.FormulaCode = "10541" Or CBatch.FormulaCode = "10101" Then  'ถ้าเป็นรำล้างไลน์ หรือ ข้าวโพดล้างไลน์
         uctlStartDate.Enable = False
         txtLotNoNew.Enabled = False
         cboBinNo.Enabled = False
         txtBatchFrom.Enabled = False
         txtBatchTo.Enabled = False
      Else
         uctlStartDate.Enable = True
         txtLotNoNew.Enabled = True
         cboBinNo.Enabled = True
         txtBatchFrom.Enabled = True
         txtBatchTo.Enabled = True
      End If
      txtNo.Text = ID
      txtBatchNo.Text = CBatch.ProductionNumber
      txtFormulaCode.Text = CBatch.FormulaCode
      txtFormulaName.Text = CBatch.FormulaName
      uctlStartDate.ShowDate = DateSerial(Mid(CBatch.BatchStartDate, 7, 4), Mid(CBatch.BatchStartDate, 4, 2), Mid(CBatch.BatchStartDate, 1, 2))
      If Len(CBatch.LotNo) > 0 Then
         Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
      Else
         Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
      End If
      txtLotNoNew.Text = CBatch.LotNo
      If CBatch.BIN_NO > 0 Then
         cboBinNo.ListIndex = IDToListIndex(cboBinNo, CBatch.BIN_NO)
      End If
      txtBatchFrom.Text = CBatch.FromBatch
      txtBatchTo.Text = CBatch.ToBatch
      txtBatchDetail.Text = CBatch.BatchDetail
   End If
End If

   Call ParentForm.RefreshGrid
   
   If txtLotNoNew.Enabled Then
      Call txtLotNoNew.SetFocus
   End If
   m_HasModify = False
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

   If Not VerifyDate(lblStartDate, uctlStartDate, False) Then
     Exit Function
   End If
      
   If Not VerifyTextControl(lblLotNo2, txtLotNoNew, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblLotNo2, txtLotNoNew, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblBinNo, cboBinNo, False) Then
      Exit Function
   End If
   
 Set CBatch = TempCollection.Item(ID)
 If Not (CBatch Is Nothing) Then
   CBatch.LotNo = txtLotNoNew.Text
   CBatch.BIN_NO = cboBinNo.ItemData(Minus2Zero(cboBinNo.ListIndex))
   CBatch.BIN_NAME = cboBinNo.Text
   CBatch.BatchStartDate = Format(uctlStartDate.ShowDate, "DD") & "/" & Format(uctlStartDate.ShowDate, "MM") & "/" & Year(uctlStartDate.ShowDate) & " " & Format(CBatch.BatchStartDate, "HH:mm:ss")
   CBatch.tempFromBatch = CBatch.FromBatch
   CBatch.tempToBatch = CBatch.ToBatch
   CBatch.FromBatch = Val(txtBatchFrom.Text)
   CBatch.ToBatch = Val(txtBatchTo.Text)
   Call GenBatchDetail(CBatch)
   
   If SplitFlag And Len(CBatch.BatchDetail) <> Len(CBatch.tempBatchDetail) Then 'ถ้าเป็นการ split ให้เข้าทำด้วย
      Dim tempCBatch As CBacthing
      Set tempCBatch = New CBacthing
      
      tempCBatch.FormulaCode = CBatch.FormulaCode
      tempCBatch.FormulaDate = CBatch.FormulaDate
      tempCBatch.FormulaName = CBatch.FormulaName
      tempCBatch.TotalBatch = CBatch.TotalBatch
      tempCBatch.TempProductionNumber = CBatch.TempProductionNumber & "-Sp"
      tempCBatch.ProductionNumber = CBatch.ProductionNumber
      tempCBatch.tempFromBatch = CBatch.tempFromBatch
      tempCBatch.tempToBatch = CBatch.tempToBatch
      tempCBatch.BatchDetail = CBatch.BatchDetail
      tempCBatch.BatchNumber = CBatch.BatchNumber
      tempCBatch.BatchStartDate = CBatch.BatchStartDate
      tempCBatch.BatchEndDate = CBatch.BatchEndDate
      tempCBatch.BIN_NAME = CBatch.BIN_NAME
      tempCBatch.BIN_NO = CBatch.BIN_NO
      tempCBatch.DestinationBin = CBatch.DestinationBin
      tempCBatch.SKIP_PART_ITEM_NO = CBatch.SKIP_PART_ITEM_NO
      tempCBatch.ProductionDate = CBatch.ProductionDate
      tempCBatch.LotId = CBatch.LotId
      tempCBatch.LotNo = CBatch.LotNo
      tempCBatch.JOB_ID = CBatch.JOB_ID
      

      Call GenBatchDetail2(tempCBatch)
      
      Call TempCollection.add(tempCBatch, Trim(tempCBatch.TempProductionNumber))
      Set tempCBatch = Nothing
   End If
 End If
 SaveData = True
End Function
Function GenBatchDetail(CBatch As CBacthing) 'สร้าง DetailBatch ใหม่ หากมีการเปลี่ยนแปลง แบตเริ่มต้นสิ้นสุด
Dim I As Long
Dim J As Long
Dim strArr() As String
Dim tempBatchDetail As String
Dim BatchFirst As Long
Dim BatchLast As Long
   tempBatchDetail = ""
    strArr = Split(CBatch.BatchDetail, ",")
   If UBound(strArr) > -1 Then
   For I = CBatch.FromBatch To CBatch.ToBatch
      For J = 0 To UBound(strArr)
         If I = strArr(J) Then
            If Not Len(tempBatchDetail) > 0 Then
                tempBatchDetail = strArr(J)
                BatchFirst = strArr(J)
            Else
               tempBatchDetail = tempBatchDetail & "," & strArr(J)
            End If
            BatchLast = strArr(J)
         End If
        Next J
      Next I
      
      CBatch.tempBatchDetail = CBatch.BatchDetail
      
      CBatch.FromBatch = BatchFirst
      CBatch.ToBatch = BatchLast
      CBatch.BatchDetail = tempBatchDetail
   End If
End Function
Function GenBatchDetail2(CBatch As CBacthing) 'สร้าง DetailBatch ใหม่ หากมีการเปลี่ยนแปลง แบตเริ่มต้นสิ้นสุด
Dim I As Long
Dim J As Long
Dim strArr() As String
Dim tempBatchDetail As String
Dim BatchFirst As Long
Dim BatchLast As Long
Dim Find As Boolean
   tempBatchDetail = ""
    strArr = Split(CBatch.BatchDetail, ",")
   If UBound(strArr) > -1 Then
   For I = CBatch.tempFromBatch To CBatch.tempToBatch
   Find = False
      For J = 0 To UBound(strArr)
         If I = strArr(J) Then
            Find = True
           Exit For
         End If
        Next J
        If Not Find Then
             If Not Len(tempBatchDetail) > 0 Then
                tempBatchDetail = I
                BatchFirst = I
            Else
               tempBatchDetail = tempBatchDetail & "," & I
            End If
            BatchLast = I
        End If
      Next I
      
      CBatch.FromBatch = BatchFirst
      CBatch.ToBatch = BatchLast
      CBatch.BatchDetail = tempBatchDetail
      CBatch.SplitFlag = "Sp"
   End If
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      Call LoadLocation(cboBinNo, Nothing, 2, , , , 2, "BIN")
     If Not (TempCollection Is Nothing) And TempCollection.Count > 0 Then
         Set CBatch = TempCollection.Item(ID)
         If Not (CBatch Is Nothing) Then
            txtNo.Text = ID
            txtBatchNo.Text = CBatch.ProductionNumber
            txtFormulaCode.Text = CBatch.FormulaCode
            txtFormulaName.Text = CBatch.FormulaName
            uctlStartDate.ShowDate = DateSerial(Mid(CBatch.BatchStartDate, 7, 4), Mid(CBatch.BatchStartDate, 4, 2), Mid(CBatch.BatchStartDate, 1, 2))
            txtBatchFrom.Text = CBatch.FromBatch
            txtBatchTo.Text = CBatch.ToBatch
            txtBatchDetail.Text = CBatch.BatchDetail
               
           If CBatch.FormulaCode <> "10541" And CBatch.FormulaCode <> "10101" Then   'ถ้าไม่ใช่รำล้างไลน์ หรือ ข้าวโพดล้างไลน์
               If Len(CBatch.LotNo) > 0 Then
                  Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               Else
                  Call txtLotNoNew.SetTextLenType(TEXT_STRING, glbSetting.LOT_NO)
               End If
               txtLotNoNew.Text = CBatch.LotNo
               If CBatch.BIN_NO > 0 Then
                  cboBinNo.ListIndex = IDToListIndex(cboBinNo, CBatch.BIN_NO)
               End If
            Else
               
               uctlStartDate.Enable = False
               txtLotNoNew.Enabled = False
               cboBinNo.Enabled = False
               txtBatchFrom.Enabled = False
               txtBatchTo.Enabled = False
            End If
         End If
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
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub txtLotNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlPlaceLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtBatchFrom_Change()
   m_HasModify = True
End Sub

Private Sub txtBatchFrom_KeyPress(KeyAscii As Integer)
   KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub txtBatchTo_Change()
   m_HasModify = True
End Sub

Private Sub txtBatchTo_KeyPress(KeyAscii As Integer)
   KeyAscii = CheckIntAscii(KeyAscii)
End Sub

Private Sub uctlStartDate_HasChange()
   m_HasModify = True
End Sub

Private Sub txtLotNoNew_Change()
     m_HasModify = True
End Sub

Private Sub txtLotNoNew_KeyPress(KeyAscii As Integer)
   KeyAscii = CheckIntAscii(KeyAscii)
End Sub
