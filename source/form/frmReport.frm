VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   8520
   ClientLeft      =   1740
   ClientTop       =   555
   ClientWidth     =   11910
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Top             =   8010
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VSPrinter7LibCtl.VSPrinter VSPrinter1 
      Height          =   7185
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11895
      _cx             =   20981
      _cy             =   12674
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   37.8787878787879
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   7
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   705
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1244
      _Version        =   131073
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureBackgroundStyle=   2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   615
      Left            =   7740
      TabIndex        =   1
      Top             =   7890
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   131073
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   9840
      TabIndex        =   2
      Top             =   7890
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   131073
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmReport"

Private HasActivate As Boolean
Public HeaderText As String
Public ReportID As String
Public ReportObject As CReportInterface
Public OKClick As Boolean
Private m_ErrorFlag As Boolean
Public ClassName As String

Private Sub cmdPrint_Click()
On Error GoTo ErrorHandler
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

'VSPrinter1.PrintDoc

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("พิมพ์ไปเครื่องพิมพ์", "-", "บันทึกไปที่ไฟล์")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      m_ErrorFlag = True
      VSPrinter1.PrintDoc (True)
      If m_ErrorFlag Then 'Error
         glbErrorLog.LocalErrorMsg = "โปรแกรมได้ทำการพิมพ์รายงานเสร็จสิ้นแล้ว"
         glbErrorLog.ShowUserError
         If (glbParameterObj.ReportKey = "CReportNormalDO001") And glbParameterObj.DocType = 1 Then
            Dim BD As CBillingDoc
            Set BD = New CBillingDoc
            BD.BILLING_DOC_ID = glbParameterObj.ID
            BD.PRINT_COUNT = glbParameterObj.PrintCount + 1
            BD.UpdatePrintCount
            Set BD = Nothing
            glbParameterObj.DocType = 0
            glbParameterObj.ReportKey = ""
         End If
      Else
         glbErrorLog.LocalErrorMsg = "โปรแกรมได้ทำการพิมพ์รายงานเสร็จสิ้นแล้ว"
         glbErrorLog.ShowUserError
         Exit Sub
      End If

   ElseIf lMenuChosen = 3 Then
      CommonDialog1.Filter = "Save Files (*.html, *.htm)|*.html;*.htm;"
      CommonDialog1.DialogTitle = "Select access file to import"
      CommonDialog1.ShowSave
      If CommonDialog1.FileName = "" Then
         Exit Sub
      End If
      
      Call FileCopy(glbParameterObj.ReportFile, CommonDialog1.FileName)
      'Call test
   End If
   
   OKClick = True
   Unload Me
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Sub test()
    Dim oDom As Object: Set oDom = CreateObject("htmlFile")
    Dim X As Long, Y As Long
    Dim oRow As Object, oCell As Object
    Dim data
    
    Y = 1: X = 1
    
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", "Deutsche Bundesbank - Macro-economic time series detail view values", False
        .Send
        oDom.Body.innerHtml = .responseText
    End With
    
    With oDom.getelementsbytagname("table")(0)
        ReDim data(1 To .Rows.Length, 1 To .Rows(1).Cells.Length)
        For Each oRow In .Rows
            For Each oCell In oRow.Cells
                data(X, Y) = oCell.innerText
                Y = Y + 1
            Next oCell
            Y = 1
            X = X + 1
        Next oRow
    End With
    
    Sheets(1).Cells(1, 1).Resize(UBound(data), UBound(data, 2)).Value = data
End Sub
Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   VSPrinter1.SaveDoc ("C:\xxx.rtf")
End Sub

Private Sub Form_Activate()
   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Me.Refresh
      Set ReportObject.VsPrint = VSPrinter1
      If Not ReportObject.Preview Then
         glbErrorLog.LocalErrorMsg = ReportObject.ErrorMsg
         glbErrorLog.ShowUserError
      End If
      Call EnableForm(Me, True)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = ClassName
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler

   Me.BackColor = GLB_FORM_COLOR
   VSPrinter1.NavBarColor = GLB_FORM_COLOR
   VSPrinter1.PaperSize = pprA4
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   HasActivate = False
   m_ErrorFlag = False

   Me.Caption = MapText("พิมพ์รายงาน")
   pnlHeader.Caption = MapText("พิมพ์รายงาน")
'    VSPrinter1.NavBar = vpnbTopPrint
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"

   Call InitMainButton(cmdPrint, "พิมพ์ (F10)")
   Call InitMainButton(cmdExit, "ออก (ESC)")

   Call EnableForm(Me, True)
   Exit Sub

ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ReportObject = Nothing
   Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
   With VSPrinter1
      .Move .Left, .Top, ScaleWidth - .Left * 2, ScaleHeight - .Top - .Left - 650
      cmdPrint.Top = .Top + ScaleHeight - .Top - .Left - 650
      cmdPrint.Left = .Left + ScaleWidth - .Left * 2 - cmdPrint.Width - cmdExit.Width - 20
      cmdExit.Top = .Top + ScaleHeight - .Top - .Left - 650
      cmdExit.Left = .Left + ScaleWidth - .Left * 2 - cmdExit.Width
      pnlHeader.Width = ScaleWidth
      .ZoomMode = zmPageWidth
   End With
End Sub
