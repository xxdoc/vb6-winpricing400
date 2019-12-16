VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSummaryReport 
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   Icon            =   "frmSummaryReport.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8595
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   15161
      _Version        =   131073
      Begin Threed.SSFrame SSFrame2 
         Height          =   7380
         Left            =   5160
         TabIndex        =   7
         Top             =   885
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   13018
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   2880
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   12
            Top             =   3600
            Visible         =   0   'False
            Width           =   1635
         End
         Begin prjFarmManagement.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   2520
            TabIndex        =   9
            Top             =   1470
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin VB.ComboBox cboGeneric 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            ItemData        =   "frmSummaryReport.frx":27A2
            Left            =   2520
            List            =   "frmSummaryReport.frx":27A4
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1110
            Visible         =   0   'False
            Width           =   3855
         End
         Begin prjFarmManagement.uctlDate uctlGenericDate 
            Height          =   435
            Index           =   0
            Left            =   2520
            TabIndex        =   8
            Top             =   690
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin Threed.SSCommand cmdSelCusGrade 
            Height          =   525
            Left            =   2760
            TabIndex        =   18
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdSelCus 
            Height          =   525
            Left            =   2760
            TabIndex        =   17
            Top             =   6120
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCheck chkCommit 
            Height          =   405
            Index           =   0
            Left            =   2520
            TabIndex        =   16
            Top             =   1920
            Visible         =   0   'False
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   714
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   810
            Visible         =   0   'False
            Width           =   2205
         End
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   0
         TabIndex        =   4
         Top             =   8280
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8460
            TabIndex        =   15
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmSummaryReport.frx":27A6
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10110
            TabIndex        =   14
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "JasmineUPC"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   2145
         End
         Begin Threed.SSCommand cmdConfig 
            Height          =   525
            Left            =   6810
            TabIndex        =   13
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   615
            Left            =   2160
            TabIndex        =   0
            Top             =   60
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   615
            Left            =   2610
            TabIndex        =   1
            Top             =   60
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   1508
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":2AC0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":339C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":36B8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":3F92
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2850
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":486C
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   7395
         Left            =   0
         TabIndex        =   6
         Top             =   870
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   13044
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   15.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String
 
Public HeaderText As String
Public MasterMode As Long

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_Checks As Collection
Private m_CyclePerMonth As Long
Private m_selCusGrade As Collection
Private m_selCus As Collection

Private m_PartGroups As Collection
Private Mr As CMasterRef
Private m_FromDate As Date
Private m_ToDate As Date
Private m_ToRcp As Date
Private m_PrintDate As Date
Private m_DocDate As Date
Private m_FromWeek As Date
Private m_ToWeek As Date
Private m_DueDate As Date
Private TempKey  As String
Private Sub GenerateTree1(MenuItems As Collection, N As Node, NodeID As String, PID As String, Level As Long)
Dim O As CMenuItem
Dim Node As Node
Dim NewNodeID As String
Dim L As Long

   For Each O In MenuItems
      If O.PARENT_KEY = PID Then
         If Level = 0 Then
            Set Node = trvMaster.Nodes.add(, tvwFirst, O.KEYWORD, O.MENU_TEXT, O.ICON_INDEX1, O.ICON_INDEX2)
            Node.Tag = O.KEYWORD
            Call GenerateTree1(MenuItems, Node, O.KEYWORD, O.KEYWORD, Level + 1)
            Node.Expanded = True
         Else
            NewNodeID = O.KEYWORD 'NodeID & "-" & O.KEYWORD
            Set Node = trvMaster.Nodes.add(N, tvwChild, NewNodeID, O.MENU_TEXT, O.ICON_INDEX1, O.ICON_INDEX2)
            Node.Tag = O.KEYWORD
            Call GenerateTree1(MenuItems, Node, NewNodeID, O.KEYWORD, Level + 1)
            Node.Expanded = False
         End If
      End If
   Next O
End Sub

Private Sub InitTreeView()
Dim Node As Node
Dim MI As CMenuItem

   Dim Programowner As String
   Programowner = glbParameterObj.Programowner

   trvMaster.Font.NAME = GLB_FONT
   trvMaster.Font.Size = 14
   
   If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("��§ҹ�����š���������ҹ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("��§ҹ�����ż����ҹ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("��§ҹ�����ͤ�Թ����к�"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-1", MapText("��§ҹ�������Թ���/��ԡ��"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-2", MapText("��§ҹ������ᾤࡨ�Թ���/��ԡ��"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-3", MapText("��§ҹ�Ҥһ�С��˹���ç "), 4, 4)
      Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 2-3", tvwChild, ROOT_TREE & " 2-3-1", MapText("��§ҹ�Ҥһ�С���Թ���˹���ç"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 2-3", tvwChild, ROOT_TREE & " 2-3-2", MapText("��§ҹ�Ҥһ�С�Ȥ�Ң���"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 2-3", tvwChild, ROOT_TREE & " 2-3-3", MapText("��§ҹ��������Ҥ��Թ���"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 2-3", tvwChild, ROOT_TREE & " 2-3-4", MapText("��§ҹ��������ҤҤ�Ң���"), 1, 2)
         Node.Expanded = False
      
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("��§ҹ�������١���"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1-1", MapText("��§ҹ�������١��� ���§����ѧ��Ѵ"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1-2", MapText("��§ҹ�������١���Ẻ�����´"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-2", MapText("��§ҹ�����ūѾ���������"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-3", MapText("��§ҹ�����ž�ѡ�ҹ"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("��§ҹ�����Թ�������ѵ�شԺ (ST001)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-3", MapText("��§ҹ Stock Card �Թ���/�ѵ�شԺ (ST002)"), 1, 2)
'      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-3-1", MapText("��§ҹ Stock Card �Թ���/�ѵ�شԺ (ST002)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-4", MapText("��§ҹ Stock ���������¤�ѧ (ST003.1)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-4-1", MapText("��§ҹ Stock �����������ѵ�شԺ (ST003.2)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5", MapText("��§ҹ Stock ���������� (ST004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-1-1", MapText("��§ҹ Stock ���������� �¡����ѹ��� (ST004-1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-1", MapText("��§ҹ��ػ�ӹǹ�������͹��� Stock (ST005)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-2", MapText("��§ҹ��ػ��Ť�ҡ������͹��� Stock (ST006)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-3", MapText("��§ҹ��û�Ѻ�ʹ Stock (ST007)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-4", MapText("��§ҹ��û�Ѻ�ʹ��¤�ѧ (ST008)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-6", MapText("��§ҹ��ػ�������͹����ѵ�شԺ (ST010)"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-6-2", MapText("��§ҹ��ػ�������͹����ѵ�شԺ Ẻ��� (ST010-2)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-7", MapText("��§ҹ��ػ�������͹����ѵ�شԺ �¡����������١���(ST011)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-8", MapText("��§ҹ�鹷ع�����Ť�Ң�µ���Թ���(ST012)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5-9", MapText("��§ҹ�鹷ع��Ť����������Ե �������١���(ST013)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-A", MapText("��§ҹ����Ѻ����ѵ�شԺ"), 4, 4)
      Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-7", MapText("��§ҹ��ë����ѵ�شԺ (RM001)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-7-1", MapText("��§ҹ��ë����ѵ�شԺ (MGP) (RM001-1)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-8", MapText("��§ҹ��ػ��ë����ѵ�شԺ 1 (RM002.1)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-8-1", MapText("��§ҹ��ػ��ë����ѵ�شԺ 2 (RM002.2)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-9", MapText("��§ҹ��ػ�ʹ���͵������� (RM003)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-10", MapText("��§ҹ���ö�������ͧ��� (RM004)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-A-1", MapText("��§ҹ�����ѵ�شԺ�������� (RM005)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-A-2", MapText("��§ҹ�����ѵ�شԺ����ѵ�شԺ (RM006)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-A-3", MapText("��§ҹ��Ť�ҫ����ѵ�شԺ�������� (RM007)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-A-4", MapText("��§ҹ�����ѵ�شԺ����������ѵ�شԺ (RM008)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-A-5", MapText("��§ҹ�����ѵ�شԺ����������ѵ�شԺ �¡����͹ (RM009)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-A", tvwChild, ROOT_TREE & " 4-A-6", MapText("��§ҹ�����ѵ�شԺ�������� ����ѵ�شԺ (RM010)"), 1, 2)
         Node.Expanded = False
         
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-B", MapText("��§ҹ����Ѻ���/�ԡ��ʴ��ػ�ó�"), 4, 4)
      Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-11", MapText("��§ҹ�������¡���ԡ���Ἱ� (IV001)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12", MapText("��§ҹ�������¡���ԡ����������� (IV002)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-1", MapText("��§ҹ����Ѻ�����ʴ�/�Ѻ��ҷ���� (IV003)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-2", MapText("��§ҹ��ػ�������¡���ԡ���Ἱ� (IV004)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-3", MapText("��§ҹ�������¡���ԡ���Ἱ�����ѹ (IV005)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-4", MapText("��§ҹ��ë��ͻ�Ш���͹  (IV006)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-5", MapText("��§ҹ�������¡���ԡ���Ἱ�����ѹ (IV007)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-6", MapText("��§ҹ�������¡���ԡ����������ѹ (IV008)"), 1, 2)
         Node.Expanded = False
         
        Set Node = trvMaster.Nodes.add(ROOT_TREE & " 4-B", tvwChild, ROOT_TREE & " 4-12-7", MapText("��§ҹ㺤���ç���   �ʴ��������¡���ԡ�¡����ç�ҹ(IV009)"), 1, 2)
         Node.Expanded = False
         
         
   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Call GenerateTree1(glbGuiConfigs.ReportMenuItems, Node, "", ROOT_TREE, 1)
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-0-1", MapText("��§ҹ LOGISTIC (PL006-0-1)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 8 Then
   Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
'     Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-6", MapText("��§ҹ��ٵá�ü�Ե (PD001)"), 1, 2)
'      Node.Expanded = False
'
'     Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-7", MapText("��§ҹ���觼�Ե (PD002)"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-8", MapText("��§ҹ㺻����Թ�Ҥ� (PD003)"), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-18", MapText("��§ҹ�ٵá�ü�Ե (PD001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-9", MapText("��§ҹ������ (PD004)"), 1, 2)
      Node.Expanded = False

       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-10", MapText("��§ҹ��ػ��ü�Ե (PD005)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-11", MapText("��§ҹ�ʹ������Ѻ��Ե (PD006)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-12", MapText("��§ҹ�鹷ع��ü�Ե (PD007)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-13", MapText("��§ҹ��ػ�ʹ���ѵ�شԺ (��ԧ) (PD008.1)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-13-1", MapText("��§ҹ��ػ�ʹ���ѵ�شԺ (�ٵ�) (PD008.2)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-14", MapText("��§ҹ��ػ�ʹ��Ե����ѵ�شԺ (�ٵ�) (PD009.1)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-14-0", MapText("��§ҹ��ػ�ʹ��Ե����ѵ�شԺ (��ԧ) (PD009.2)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-14-1", MapText("��§ҹ��ػ�ʹ��Ե�����ԧ (PD009.3)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-14-2", MapText("��§ҹ��ػ�ʹ��Ե�����ԧ��Ť�ҵ���ѵ�شԺ (PD009.4)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-15", MapText("��§ҹ�ʹ���ѵ�شԺ�����Ե�ѳ�� (PD010)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-16", MapText("��§ҹ�ʹ��Ե����ѵ�شԺ (PD011)"), 1, 2)
      Node.Expanded = False
   
       Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-17", MapText("��§ҹ��ػ�ӹǹ������ѵ�شԺ����Թ��� (PD012)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-19", MapText("��§ҹ����������� ����ѵ�شԺ (PD013)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-20", MapText("��§ҹ��ػ�ʹ������ѵ�����ͼ�Ե�¡����ѵ�شԺ ��� ��͹��Ե(PD014)"), 1, 2)
      Node.Expanded = False
      
      '8-18 is already used
   ElseIf MasterMode = 10 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 10-0-1", MapText("��§ҹ �ʹ��µ����ѡ�ҹ��� �١��� �Թ��� (COM001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 10-0-2", MapText("��§ҹ �ʹ��µ����ѡ�ҹ��� + ����˹��Ŵ˹�� + %��� (COM002)"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 10-0-3", MapText("��§ҹ ��Ҥ���Ԫ��蹢ͧ ����Ź�� ����ʹ�Ѻ����˹��  (COM003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 10-0-4", MapText("��§ҹ�����š�â�¢ͧ����Ź�� ���§����١��� ��Ф�� INCENTIVE (COM004)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 9 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 9-1", MapText("��§ҹ�Ѻ����Թ��� "), 4, 4)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-1", tvwChild, ROOT_TREE & " 9-1-1", MapText("��§ҹ��ú�è��Թ��� (IW001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 9-2", MapText("��§ҹ�����͡�Թ���"), 4, 4)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-2", tvwChild, ROOT_TREE & " 9-2-1", MapText("��§ҹ��è��������������ٻ (EW001)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-2", tvwChild, ROOT_TREE & " 9-2-1", MapText("��§ҹ��è��������������ٻ (EW001)"), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-2", tvwChild, ROOT_TREE & " 9-2-2", MapText("��§ҹ��è��������������ٻ �¡��������� (EW002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 9-3", MapText("��§ҹ��ѧ�Թ���"), 4, 4)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-3", tvwChild, ROOT_TREE & " 9-3-1", MapText("��§ҹ KPI (WH001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-3", tvwChild, ROOT_TREE & " 9-3-2", MapText("��§ҹ�ӹǹ��������� (WH002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-3", tvwChild, ROOT_TREE & " 9-3-3", MapText("��§ҹ STOCK CARD ��ѧ�Թ���������ٻ (WH003)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " 9-3", tvwChild, ROOT_TREE & " 9-3-4", MapText("��§ҹ STOCK CARD ��ѧ�Թ���������ٻ �����͵ (WH004)"), 1, 2)
      Node.Expanded = False
      
    ElseIf MasterMode = 11 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-1", MapText("��§ҹ�Ҥһ�С���Թ���˹���ç"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-2", MapText("��§ҹ�Ҥһ�С�Ȥ�Ң���"), 1, 2)
      Node.Expanded = False
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
   End If

   Label1.Caption = ItemCount
   
   Call EnableForm(Me, True)
End Sub

Private Sub FillReportInput(R As CReportInterface)
Dim C As CReportControl

   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_DOC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_DOC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               ElseIf C.Param2 = "TO_PAY_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "PRINT_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               End If
            End If
            If C.Param2 = "FROM_DOC_DATE" Or C.Param2 = "FROM_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_DOC_DATE" Or C.Param2 = "TO_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_PAY_DATE" Then
               m_ToRcp = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "PRINT_DATE" Then
               m_PrintDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "DOC_DATE" Then
               m_DocDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "FROM_WEEK" Or C.Param2 = "FROM_WEEK_DATE" Then
               m_FromWeek = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_WEEK_DATE" Or C.Param2 = "TO_SUP_DATE" Then
               m_ToWeek = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
   
        If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param2)
         End If
      End If
    
   Next C
End Sub

Private Function VerifyReportInput() As Boolean
Dim C As CReportControl

   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "D") Then
         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   Next C
   VerifyReportInput = True
End Function

Private Sub cboGeneric_Click(Index As Integer)
Dim Node As Node
Dim TempID As Long

   Set Node = trvMaster.SelectedItem
   
   If (Node.Key = ROOT_TREE & " 4-1") Or _
      (Node.Key = ROOT_TREE & " 8-13") Or _
      (Node.Key = ROOT_TREE & " 8-20") Or _
      (Node.Key = ROOT_TREE & " 8-13-1") Or _
      (Node.Key = ROOT_TREE & " 8-14") Or _
      (Node.Key = ROOT_TREE & " 8-14-0") Or _
      (Node.Key = ROOT_TREE & " 8-14-1") Or _
      (Node.Key = ROOT_TREE & " 8-14-2") Or _
      (Node.Key = ROOT_TREE & " 8-15") Or _
      (Node.Key = ROOT_TREE & " 8-16") Or _
      (Node.Key = ROOT_TREE & " 8-17") Then
      If Index = 2 Then
         TempID = cboGeneric(Index).ItemData(Minus2Zero(cboGeneric(Index).ListIndex))
         If TempID > 0 Then
            Call LoadPartType(cboGeneric(Index + 1), , TempID)
         End If
      End If
   ElseIf (Node.Key = ROOT_TREE & " 4-3") Or _
      (Node.Key = ROOT_TREE & " 4-3-1") Or _
      (Node.Key = ROOT_TREE & " 4-5-6") Or _
      (Node.Key = ROOT_TREE & " 4-4") Or _
      (Node.Key = ROOT_TREE & " 4-4-1") Or _
      (Node.Key = ROOT_TREE & " 4-7") Or _
      (Node.Key = ROOT_TREE & " 4-12-1") Or _
      (Node.Key = ROOT_TREE & " 4-12-2") Or _
      (Node.Key = ROOT_TREE & " 4-12-3") Or _
      (Node.Key = ROOT_TREE & " 4-12-5") Or _
      (Node.Key = ROOT_TREE & " 4-8") Or _
      (Node.Key = ROOT_TREE & " 4-8-1") Or _
      (Node.Key = ROOT_TREE & " 4-9") Or _
      (Node.Key = ROOT_TREE & " 4-10") Or _
      (Node.Key = ROOT_TREE & " 4-A-1") Or _
      (Node.Key = ROOT_TREE & " 4-A-2") Or _
      (Node.Key = ROOT_TREE & " 4-A-3") Or (Node.Key = ROOT_TREE & " 4-A-4") Or (Node.Key = ROOT_TREE & " 4-A-5") Or (Node.Key = ROOT_TREE & " 4-A-6") Or _
      (Node.Key = ROOT_TREE & " 4-5") Or (Node.Key = ROOT_TREE & " 4-5-1") Or _
      (Node.Key = ROOT_TREE & " 4-5-1") Or _
      (Node.Key = ROOT_TREE & " 4-5-2") Or _
      (Node.Key = ROOT_TREE & " 4-5-3") Or _
      (Node.Key = ROOT_TREE & " 4-5-4") Or _
      (Node.Key = ROOT_TREE & " 4-5-5") Or _
      (Node.Key = ROOT_TREE & " 4-11") Or _
      (Node.Key = ROOT_TREE & " 4-12") Then
      If Index = 1 Then
         TempID = cboGeneric(Index).ItemData(Minus2Zero(cboGeneric(Index).ListIndex))
         If TempID > 0 Then
            Call LoadPartType(cboGeneric(Index + 1), , TempID)
         End If
      End If
   ElseIf (Node.Key = ROOT_TREE & " A-2-5") Then
      If Index = 1 Then
         TempID = cboGeneric(Index).ItemData(Minus2Zero(cboGeneric(Index).ListIndex))
         If TempID > 0 Then
            Call LoadBankBranch(cboGeneric(Index + 1), , TempID)
         End If
      End If
   End If
End Sub

Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long
Dim ReportMode As Long

   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
      
   ReportKey = trvMaster.SelectedItem.Key
   
   ReportMode = 1
   
   If ReportKey = "Root J-1-2" Then 'Or ReportKey = "Root 9-1-1"
      ReportMode = 2
   End If
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   'Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If
   
   frmReportConfig.ReportMode = ReportMode
   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMaster.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim NAME As String
Dim ClassName As String
   
   Key = trvMaster.SelectedItem.Key
   NAME = trvMaster.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Not (trvMaster.SelectedItem Is Nothing) Then
      Call Report.AddParam(trvMaster.SelectedItem.Text, "REPORT_TEXT")
   End If
   
   If Key = ROOT_TREE & " 1-1" Then
      Set Report = New CReportAdmin001
      ClassName = "CReportAdmin001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 1-2" Then
      Set Report = New CReportAdmin002
      ClassName = "CReportAdmin002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 1-3" Then
      Set Report = New CReportAdmin003
      ClassName = "CReportAdmin003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 2-1" Then
      Set Report = New CReportPackage001
      ClassName = "CReportPackage001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 2-2" Then
      Set Report = New CReportPackage002
      ClassName = "CReportPackage002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 2-3-1" Then
      Set Report = New CReportExWorksPice001
      ClassName = "CReportExWorksPice001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 2-3-2" Then
      Set Report = New CReportExWorksPice002
      ClassName = "CReportExWorksPice002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 2-3-3" Then
      Set Report = New CReportExWorksPice003
      ClassName = "CReportExWorksPice003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 2-3-4" Then
      Set Report = New CReportExWorksPice004
      ClassName = "CReportExWorksPice004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-1" Then
      Set Report = New CReportMain001
      ClassName = "CReportMain001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-1-1" Then
      Set Report = New CReportMain001_1
      ClassName = "CReportMain001_1"
      SelectFlag = True
 ElseIf Key = ROOT_TREE & " 3-1-2" Then
      Set Report = New CReportMain001_2
      ClassName = "CReportMain001_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-2" Then
      Set Report = New CReportMain002
      ClassName = "CReportMain002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-3" Then
      Set Report = New CReportMain003
      ClassName = "CReportMain003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-1" Then
      Set Report = New CReportInventory001
      ClassName = "CReportInventory001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-2" Then
      Set Report = New CReportInventory002
      ClassName = "CReportInventory002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-3" Then
      Set Report = New CReportInventory003
      ClassName = "CReportInventory003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-3-1" Then
      Set Report = New CReportInventory023
      ClassName = "CReportInventory023"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-4" Then
      Set Report = New CReportInventory004_5
      ClassName = "CReportInventory004_5"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-4-1" Then
      Set Report = New CReportInventory004_4
      ClassName = "CReportInventory004_4"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5" Then
      Set Report = New CReportInventory004_3
      ClassName = "CReportInventory004_3"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-1-1" Then
      Set Report = New CReportInventory004_6
      ClassName = "CReportInventory004_6"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-1" Then
      Set Report = New CReportInventory015
      ClassName = "CReportInventory015"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-2" Then
      Set Report = New CReportInventory016
      ClassName = "CReportInventory016"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-3" Then
      Set Report = New CReportInventory017
      ClassName = "CReportInventory017"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-4" Then
      Set Report = New CReportInventory018
      ClassName = "CReportInventory018"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-6" Then
      Set Report = New CReportInventory026
      ClassName = "CReportInventory026"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-6-2" Then
      Set Report = New CReportInventory026_2
      ClassName = "CReportInventory026_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-7" Then
      Set Report = New CReportInventory026_1
      ClassName = "CReportInventory026_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-8" Then
      Set Report = New CReportInventory028
      ClassName = "CReportInventory028"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5-9" Then
      Set Report = New CReportInventory030
      ClassName = "CReportInventory030"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-6" Then
      Set Report = New CReportInventory006
      ClassName = "CReportInventory006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-7" Then
      Set Report = New CReportInventory007
      ClassName = "CReportInventory007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-7-1" Then
      Set Report = New CReportInventory007_3
      ClassName = "CReportInventory007_3"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12-1" Then
      Set Report = New CReportInventory007_1
      ClassName = "CReportInventory007_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12-2" Then
      Set Report = New CReportInventory024_1
      ClassName = "CReportInventory024_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12-3" Then
      Set Report = New CReportInventory025
      ClassName = "CReportInventory025"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12-4" Then
      Set Report = New CReportInventory007_2
      ClassName = "CReportInventory007_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12-5" Then
      Set Report = New CReportInventory029
      ClassName = "CReportInventory029"
      SelectFlag = True
      
   ElseIf Key = ROOT_TREE & " 4-12-7" Then
      Set Report = New CReportInventory033
      ClassName = "CReportInventory033"
      SelectFlag = True
      '4-12-7
   ElseIf Key = ROOT_TREE & " 4-12-6" Then
      Set Report = New CReportInventory031
      ClassName = "CReportInventory031"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-8" Then
      Set Report = New CReportInventory008
      ClassName = "CReportInventory008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-8-1" Then
      Set Report = New CReportInventory008_1
      ClassName = "CReportInventory008_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-9" Then
      Set Report = New CReportInventory009
      ClassName = "CReportInventory009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-10" Then
      Set Report = New CReportInventory010
      ClassName = "CReportInventory010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-11" Then
      Set Report = New CReportInventory021
      ClassName = "CReportInventory021"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12" Then
      Set Report = New CReportInventory022
      ClassName = "CReportInventory022"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-A-1" Then
      Set Report = New CReportInventory012
      ClassName = "CReportInventory012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-A-2" Then
      Set Report = New CReportInventory013
      ClassName = "CReportInventory013"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-A-3" Then
      Set Report = New CReportInventory014
      ClassName = "CReportInventory014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-A-4" Then
      Set Report = New CReportInventory013_1
      ClassName = "CReportInventory013_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-A-5" Then
      Set Report = New CReportInventory032
      ClassName = "CReportInventory032"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-A-6" Then
      Set Report = New CReportInventory012_1
      ClassName = "CReportInventory012_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-1" Then
      Set Report = New CReportSell001
      ClassName = "CReportSell001"
      Call Report.AddParam(12, "DOCUMENT_TYPE")
      SelectFlag = True
    ElseIf Key = ROOT_TREE & " 5-1-0" Then
      Set Report = New CReportSell001_0
      ClassName = "CReportSell001_0"
      Call Report.AddParam(12, "DOCUMENT_TYPE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-1-0-1" Then
      Set Report = New CReportSell001_0_1
      ClassName = "CReportSell001_0_1"
      Call Report.AddParam(12, "DOCUMENT_TYPE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-1-1" Then
      Set Report = New CReportSell001_1
      ClassName = "CReportSell001_1"
      Call Report.AddParam(12, "DOCUMENT_TYPE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-1-2" Then
      Set Report = New CReportSell001_2
      ClassName = "CReportSell001_2"
      Call Report.AddParam(12, "DOCUMENT_TYPE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-2" Then
      Set Report = New CReportSell006
      ClassName = "CReportSell006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-2-1" Then
      Set Report = New CReportSell006_1
      ClassName = "CReportSell006_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-2-1-1" Then
      Set Report = New CReportSell006_3
      ClassName = "CReportSell006_3"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-2-1-2" Then
      Set Report = New CReportSell006_4
      ClassName = "CReportSell006_4"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-2-2" Then
      Set Report = New CReportSell006_2
      ClassName = "CReportSell006_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-3" Then
      Set Report = New CReportSell004
      ClassName = "CReportSell004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-4" Then
      Set Report = New CReportSell005
      ClassName = "CReportSell005"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-5" Then
      Set Report = New CReportSell002
      ClassName = "CReportSell002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-5-1" Then
      Set Report = New CReportSell002_1
      ClassName = "CReportSell002_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-5-1-2" Then
      Set Report = New CReportSell002_4
      ClassName = "CReportSell002_4"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-5-1-3" Then
      Set Report = New CReportSell002_5
      ClassName = "CReportSell002_5"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-5-2" Then
      Set Report = New CReportSell002_2
      ClassName = "CReportSell002_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-6" Then
      Set Report = New CReportSell003
      ClassName = "CReportSell003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-1-1" Then
      Set Report = New CReportSell007
      ClassName = "CReportSell007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-1-2" Then
      Set Report = New CReportSell008
      ClassName = "CReportSell008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-1-3" Then
      Set Report = New CReportSell009
      ClassName = "CReportSell009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-1-4" Then
      Set Report = New CReportSell010
      ClassName = "CReportSell010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-1-5" Then
      Set Report = New CReportSell006_5
      ClassName = "CReportSell006_5"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-7-1" Then
      Set Report = New CReportSellCT001
      ClassName = "CReportSellCT001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-7-2" Then
      Set Report = New CReportSellCT002
      ClassName = "CReportSellCT002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-7-3" Then
      Set Report = New CReportSellCT003
      ClassName = "CReportSellCT003"
      SelectFlag = True

   ElseIf Key = ROOT_TREE & " 5-6-1" Then
      Set Report = New CReportLedger007
      ClassName = "CReportLedger007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-7" Then
   ElseIf Key = ROOT_TREE & " 5-8" Then
   ElseIf Key = ROOT_TREE & " 5-9" Then
   ElseIf Key = ROOT_TREE & " 5-10" Then
   ElseIf Key = ROOT_TREE & " 5-11" Then
   ElseIf Key = ROOT_TREE & " 5-12" Then
   ElseIf Key = ROOT_TREE & " 5-13" Then
   ElseIf Key = ROOT_TREE & " 5-14" Then
   ElseIf Key = ROOT_TREE & " 5-15" Then
   ElseIf Key = ROOT_TREE & " F-1-1" Then
      Set Report = New CReportProfit001
      ClassName = "CReportProfit001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-1-2" Then
      Set Report = New CReportProfit001_2
      ClassName = "CReportProfit001_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-2" Then
      Set Report = New CReportProfit002
      ClassName = "CReportProfit002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-3" Then
      Set Report = New CReportProfit003
      ClassName = "CReportProfit003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-4" Then
      Set Report = New CReportProfit004
      ClassName = "CReportProfit004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-5" Then
      Set Report = New CReportProfit005
      ClassName = "CReportProfit005"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-6" Then
      Set Report = New CReportProfit006
      ClassName = "CReportProfit006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " F-1-7" Then
      Set Report = New CReportProfit007
      ClassName = "CReportProfit007"
      SelectFlag = True
 ElseIf Key = ROOT_TREE & " F-1-8" Then
      Set Report = New CReportProfit008
      ClassName = "CReportProfit008"
      SelectFlag = True
 ElseIf Key = ROOT_TREE & " F-1-9" Then
      Set Report = New CReportProfit009
      ClassName = "CReportProfit009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-1" Then
      Set Report = New CReportPerson001
      ClassName = "CReportPerson001"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " 6-2" Then
      Set Report = New CReportPerson002
      ClassName = "CReportPerson002"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " 6-3" Then
      Set Report = New CReportPerson003
      ClassName = "CReportPerson003"
      SelectFlag = True
   
   
   ElseIf Key = ROOT_TREE & " 6-4" Then
      Set Report = New CReportPerson004
      ClassName = "CReportPerson004"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " 6-5" Then
      Set Report = New CReportPerson005
      ClassName = "CReportPerson005"
      SelectFlag = True
      
   ElseIf Key = ROOT_TREE & " 8-1" Then
      Set Report = New CReportProduct001
      ClassName = "CReportProduct001"
      SelectFlag = True
      
   ElseIf Key = ROOT_TREE & " 8-2" Then
      Set Report = New CReportProduct002
      ClassName = "CReportProduct002"
      SelectFlag = True
 ElseIf Key = ROOT_TREE & " 8-3" Then
      Set Report = New CReportProduct003
      ClassName = "CReportProduct003"
      SelectFlag = True
 ElseIf Key = ROOT_TREE & " 8-4" Then
      Set Report = New CReportProduct004
      ClassName = "CReportProduct004"
      SelectFlag = True
      
 ElseIf Key = ROOT_TREE & " 8-5" Then
      Set Report = New CReportProduct005
      ClassName = "CReportProduct005"
      SelectFlag = True
ElseIf Key = ROOT_TREE & " 8-6" Then
      Set Report = New CReportProduct006
      ClassName = "CReportProduct006"
      SelectFlag = True
ElseIf Key = ROOT_TREE & " 8-7" Then
      Set Report = New CReportProduct007
      ClassName = "CReportProduct007"
      SelectFlag = True
ElseIf Key = ROOT_TREE & " 8-8" Then
      Set Report = New CReportProduct008
      ClassName = "CReportProduct008"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 8-9" Then
      Set Report = New CReportProduct009
      ClassName = "CReportProduct009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-10" Then
      Set Report = New CReportProduct010
      ClassName = "CReportProduct010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-11" Then
      Set Report = New CReportProduct011
      ClassName = "CReportProduct011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-12" Then
      Set Report = New CReportProduct012
      ClassName = "CReportProduct012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-13" Then
      Set Report = New CReportProduct013
      ClassName = "CReportProduct013"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-13-1" Then
      Set Report = New CReportProduct013
      ClassName = "CReportProduct013"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-14" Then
      Set Report = New CReportProduct014
      ClassName = "CReportProduct014"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-14-1" Then
      Set Report = New CReportProduct014_1
      ClassName = "CReportProduct014_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-14-2" Then
      Set Report = New CReportProduct014_2
      ClassName = "CReportProduct014_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-14-0" Then
      Set Report = New CReportProduct014
      ClassName = "CReportProduct014"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-15" Then
      Set Report = New CReportProduct015
      ClassName = "CReportProduct015"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-16" Then
      Set Report = New CReportProduct016
      ClassName = "CReportProduct016"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-17" Then
      Set Report = New CReportProduct017
      ClassName = "CReportProduct017"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-18" Then
      Set Report = New CReportProduct018
      ClassName = "CReportProduct018"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-19" Then
      Set Report = New CReportProduct019
      ClassName = "CReportProduct019"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-20" Then
      Set Report = New CReportProduct020
      ClassName = "CReportProduct020"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-1-1" Then
      Set Report = New CReportIncomeWH
      
      Picture1.Picture = LoadPicture(glbParameterObj.IncomeGoods)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         
      ClassName = "CReportIncomeWH"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-2-1" Then
      Set Report = New CReportPayOffWH
      
      Picture1.Picture = LoadPicture(glbParameterObj.PayOffGoods)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      
      ClassName = "CReportPayOffWH"
      SelectFlag = True
'   ElseIf Key = ROOT_TREE & " 9-2-1-2" Then
'      Set Report = New CReportPayOffWH002
'
''      Picture1.Picture = LoadPicture(glbParameterObj.PayOffGoods)
''      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
'
'      ClassName = "CReportPayOffWH002"
'      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-2-2" Then
      Set Report = New CReportEW002
      ClassName = "CReportEW002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-3-1" Then
      Set Report = New CReportKPIWh
      ClassName = "CReportKPIWh"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-3-2" Then
      Set Report = New CReportAgePartWh
      ClassName = "CReportAgePartWh"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-3-3" Then
      Set Report = New CReportInventoryWh001
      ClassName = "CReportInventoryWh001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 9-3-4" Then
      Set Report = New CReportInventoryWh002
      ClassName = "CReportInventoryWh002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-1" Then
      Set Report = New CReportAR001
      ClassName = "CReportAR001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-2" Then
      Set Report = New CReportAR002
      ClassName = "CReportAR002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-3" Then
      Set Report = New CReportAR003
      ClassName = "CReportAR003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-3.1" Then
      Set Report = New CReportAR003_1
      ClassName = "CReportAR003_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-4" Then
      Set Report = New CReportAR004
      ClassName = "CReportAR004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-4-2" Then
      Set Report = New CReportAR004_6
      ClassName = "CReportAR004_6"
      SelectFlag = True
      
     ElseIf Key = ROOT_TREE & " A-2-4-4" Then
      Set Report = New CReportAR004_9
      ClassName = "CReportAR004_9"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-4-22" Then
      Set Report = New CReportAR004_11
      ClassName = "CReportAR004_11"
      SelectFlag = True
      
      Call Report.AddParam(m_selCusGrade, "CollCusGrade")
      Call Report.AddParam(m_selCus, "CollCus")
   ElseIf Key = ROOT_TREE & " A-2-4-1" Then
      Set Report = New CReportAR004_2
      ClassName = "CReportAR004_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-4-3" Then
      Set Report = New CReportAR004_7_1
      ClassName = "CReportAR004_7_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-4-5" Then
      Set Report = New CReportAR004_8
      ClassName = "CReportAR004_8"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-4-6" Then
      Set Report = New CReportAR004_10
      ClassName = "CReportAR004_10"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-5" Then
      Set Report = New CReportAR005
      ClassName = "CReportAR005"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-6" Then
      Set Report = New CReportAR006
      ClassName = "CReportAR006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-7" Then
      Set Report = New CReportAR007
      ClassName = "CReportAR007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-7-1" Then
      Set Report = New CReportAR015
      ClassName = "CReportAR015"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-8" Then
      Set Report = New CReportAR009
      ClassName = "CReportAR009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-9" Then
      Set Report = New CReportAR004_4
      ClassName = "CReportAR004_4"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-11" Then
      Set Report = New CReportCash001
      ClassName = "CReportCash001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-12" Then
      Set Report = New CReportCash002
      ClassName = "CReportCash002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-13" Then
      Set Report = New CReportAR011
      ClassName = "CReportAR011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-14" Then
      Set Report = New CReportAR012
      ClassName = "CReportAR012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-15" Then
      Set Report = New CReportAR013
      ClassName = "CReportAR013"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-16" Then
      Set Report = New CReportAR014
      ClassName = "CReportAR014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-17" Then
      Set Report = New CReportAR009_2
      ClassName = "CReportAR009_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-18" Then
      Set Report = New CReportAR018
      ClassName = "CReportAR018"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " A-2-21" Then
      Set Report = New CReportAR021
      ClassName = "CReportAR021"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " J-1-1" Then
      Set Report = New CReportJv001
      ClassName = "CReportJv001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " J-1-2" Then
      Set Report = New CReportJv002
      ClassName = "CReportJv002"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " BUY-1-1" Then
      Set Report = New CReportMain002
      ClassName = "CReportMain002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-1" Then
      Set Report = New CReportAP004
      ClassName = "CReportAP004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-2" Then
      Set Report = New CReportAP006
      ClassName = "CReportAP006"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " AP-1-3" Then
      Set Report = New CReportAP002
      ClassName = "CReportAP002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-4" Then
      Set Report = New CReportAP003
      ClassName = "CReportAP003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-5" Then
      Set Report = New CReportAP011
      ClassName = "CReportAP011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-5-1" Then
      Set Report = New CReportAP014
      ClassName = "CReportAP014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-6" Then
      Set Report = New CReportAP012
      ClassName = "CReportAP012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-6-1" Then
      Set Report = New CReportAP015
      ClassName = "CReportAP015"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-7" Then
      Set Report = New CReportAP007
      ClassName = "CReportAP007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-8" Then
      Set Report = New CReportAP008
      ClassName = "CReportAP008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-9" Then
      Set Report = New CReportAP009
      ClassName = "CReportAP009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-10" Then
      Set Report = New CReportAP010
      ClassName = "CReportAP010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-11" Then
      Set Report = New CReportAP013
      ClassName = "CReportAP013"
      SelectFlag = True
      
    ElseIf Key = ROOT_TREE & " AP-1-11-1" Then
      Set Report = New CReportAP017
      ClassName = "CReportAP017"
      SelectFlag = True
     ElseIf Key = ROOT_TREE & " AP-1-11-2" Then
      Set Report = New CReportAP018
      ClassName = "CReportAP018"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-11-3" Then
      Set Report = New CReportAP019
      ClassName = "CReportAP019"
      Call Report.AddParam(1, "REPORT_GROUP")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-11-4" Then
      Set Report = New CReportAP019
      ClassName = "CReportAP019"
      Call Report.AddParam(2, "REPORT_GROUP")
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " AP-1-11-5" Then
      Set Report = New CReportAP020
      ClassName = "CReportAP020"
      Call Report.AddParam(2, "REPORT_GROUP")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-11-6" Then
      Set Report = New CReportAP021
      Picture1.Picture = LoadPicture(glbParameterObj.PaymentVoucher)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      
      Picture1.Picture = LoadPicture(glbParameterObj.AccountTransfer)
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND2")
      
      ClassName = "CReportAP021"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-11-7" Then
      Set Report = New CReportAP022
      ClassName = "CReportAP022"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AP-1-11-8" Then
      Set Report = New CReportAP023
      ClassName = "CReportAP023"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AB-1-1" Then
      Set Report = New CReportAP016
      ClassName = "CReportAP016"
      SelectFlag = True
    ElseIf Key = ROOT_TREE & " AB-1-3" Then
      Set Report = New CReportAP016_2
      ClassName = "CReportAP016_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " AB-1-4" Then
      Set Report = New CReportAP016_4
      ClassName = "CReportAP016_4"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-1-1" Then
      Set Report = New CReportPlanPart001
      ClassName = "CReportPlanPart001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-1-2" Then
      Set Report = New CReportPlanPart002
      ClassName = "CReportPlanPart002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-0-1" Then
      Set Report = New CReportPlanning001
      ClassName = "CReportPlanning001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 10-0-1" Then
      Set Report = New CReportCommission001
      ClassName = "CReportCommission001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 10-0-2" Then
      Set Report = New CReportCommission002
      ClassName = "CReportCommission002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 10-0-3" Then
      Set Report = New CReportCommission003
      ClassName = "CReportCommission003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 10-0-4" Then
      Set Report = New CReportFrelance001
      ClassName = "CReportFrelance001"
      SelectFlag = True
   End If

   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If

      Call FillReportInput(Report)
      Call Report.AddParam(NAME, "REPORT_NAME")
      Call Report.AddParam(Key, "REPORT_KEY")

      Set frmReport.ReportObject = Report
      frmReport.ClassName = ClassName
      frmReport.HeaderText = MapText("�������§ҹ")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
   End If
End Sub

Private Sub cmdSelCus_Click()
 Dim OKClick As Boolean
      frmSelectItem.itemType = 2
      Set frmSelectItem.TempCollection = m_selCus
      Load frmSelectItem
      frmSelectItem.Show 1
      OKClick = frmSelectItem.OKClick
      Unload frmSelectItem
      Set frmSelectItem = Nothing
      
      If OKClick Then
      End If
End Sub

Private Sub cmdSelCusGrade_Click()
   Dim OKClick As Boolean
      frmSelectItem.itemType = 1
      Set frmSelectItem.TempCollection = m_selCusGrade
      Load frmSelectItem
      frmSelectItem.Show 1
      OKClick = frmSelectItem.OKClick
      Unload frmSelectItem
      Set frmSelectItem = Nothing
      
      If OKClick Then
      End If
End Sub

Private Sub Form_Activate()
Dim ItemCount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
            
      Call QueryData(True)
      m_HasActivate = True
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
'      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      If cmdOK.Enabled Then
         Call cmdOK_Click
      End If
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Resize()
   pnlHeader.Width = ScaleWidth
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   If ScaleWidth <= 0 Then
      trvMaster.Width = 0
   Else
      trvMaster.Width = ScaleWidth - SSFrame2.Width
   End If
   SSFrame2.Left = trvMaster.Width
   If ScaleHeight <= 0 Then
      trvMaster.HEIGHT = 0
   Else
      trvMaster.HEIGHT = ScaleHeight - pnlHeader.HEIGHT - pnlFooter.HEIGHT
   End If
   SSFrame2.HEIGHT = trvMaster.HEIGHT
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.HEIGHT
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20
   cmdConfig.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20 - cmdConfig.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_Checks = Nothing
   Set m_PartGroups = Nothing
   Set Mr = Nothing
   Set m_selCusGrade = Nothing
   Set m_selCus = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   SSFrame2.BackColor = GLB_FORM_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitMainButton(cmdOK, MapText("����� (F10)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("����� (F10)"))
   Call InitMainButton(cmdConfig, MapText("��Ѻ���"))
   
   Call InitMainButton(cmdSelCusGrade, MapText("���͡�дѺ�١���"))
   Call InitMainButton(cmdSelCus, MapText("���͡�١���"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelCus.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelCusGrade.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCheckBox(chkCommit(0), "�ҹ��������")
   
   Call InitTreeView
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   Call InitFormLayout
   
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
   
   Set m_ReportControls = New Collection
   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
   Set m_Checks = New Collection
   Set m_PartGroups = New Collection
   Set Mr = New CMasterRef
   Set m_selCusGrade = New Collection
   Set m_selCus = New Collection
End Sub

Private Sub UnloadAllControl()
Dim I As Long
Dim J As Long

   I = m_Labels.Count
   While I > 0
      Call Unload(m_Labels(I))
      Call m_Labels.Remove(I)
      I = I - 1
   Wend
   
   I = m_Texts.Count
   While I > 0
      Call Unload(m_Texts(I))
      Call m_Texts.Remove(I)
      I = I - 1
   Wend

   I = m_Dates.Count
   While I > 0
      Call Unload(m_Dates(I))
      Call m_Dates.Remove(I)
      I = I - 1
   Wend

   I = m_Combos.Count
   While I > 0
      Call Unload(m_Combos(I))
      Call m_Combos.Remove(I)
      I = I - 1
   Wend
   
   I = m_TextLookups.Count
   While I > 0
      Call Unload(m_TextLookups(I))
      Call m_TextLookups.Remove(I)
      I = I - 1
   Wend
   
   I = m_Checks.Count
   While I > 0
      Call Unload(m_Checks(I))
      Call m_Checks.Remove(I)
      I = I - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub

Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long
Dim C As CReportControl

   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "LU") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            m_Combos(C.ControlIndex).Left = PrevLeft
            m_Combos(C.ControlIndex).Top = PrevTop
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).HEIGHT
            PrevLeft = m_Combos(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).HEIGHT
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            m_Texts(C.ControlIndex).Left = PrevLeft
            m_Texts(C.ControlIndex).Left = PrevLeft
            m_Texts(C.ControlIndex).Top = PrevTop
            m_Texts(C.ControlIndex).Width = C.Width
            Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
            m_Texts(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).HEIGHT
            PrevLeft = m_Texts(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).HEIGHT
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
            ElseIf C.ControlType = "CH" Then
            m_Checks(C.ControlIndex).Left = PrevLeft
            m_Checks(C.ControlIndex).Top = PrevTop + 50
            m_Checks(C.ControlIndex).Width = C.Width
            m_Checks(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Checks(C.ControlIndex).Top + m_Checks(C.ControlIndex).HEIGHT
            PrevLeft = m_Checks(C.ControlIndex).Left
            PrevWidth = C.Width
            
            ElseIf C.ControlType = "CH" Then
            m_Checks(C.ControlIndex).Left = PrevLeft
            m_Checks(C.ControlIndex).Top = PrevTop
            m_Checks(C.ControlIndex).Width = C.Width
            m_Checks(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Checks(C.ControlIndex).Top + m_Checks(C.ControlIndex).HEIGHT
            PrevLeft = m_Checks(C.ControlIndex).Left
            PrevWidth = C.Width
         End If
      Else 'Label
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            m_Labels(C.ControlIndex).Visible = True
      End If
   Next C
End Sub

Private Sub LoadComboData()
Dim C As CReportControl
   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-1" Then
            If C.ComboLoadID = 1 Then
               Call InitUserGroupOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitUserOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitLoginOrderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 2-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadFeatureType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadUnit(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitFeatureOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 2-2" Then
            If C.ComboLoadID = 1 Then
               Call InitSocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadUnit(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadUnit(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-3-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-4-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-6" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-6-2" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-8" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-9" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-1" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-6" Or _
          trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-4" Then
          
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitImportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport4_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitExportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadLayout(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport4_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitExportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadLayout(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport4_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
                                 
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitExportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadLayout(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport4_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-8-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_9Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-10" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport4_10Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_A_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_A_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_A_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_9Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitExportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadLayout(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport4_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitExportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadLayout(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport4_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_1OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
        
        If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If

        If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If
        
        If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If
        
        If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If
        
        
        If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1-0" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_1OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1-0-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_1OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_1OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_2OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-2-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_2OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-2-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_2OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-2-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_2OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-2-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_2OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_3OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_4OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport5_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-5-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , PRTITEM_SET)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport5_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-5-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , PRTITEM_SET)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call initDataMode(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport5_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_6OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_6OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_6OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportA_1_3OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-6-1" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport5_5_1OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-7" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-8" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-9" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-10" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-11" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-12" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-13" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-14" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-15" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         

         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadWorkStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSex(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitThaiYear(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitThaiYear(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-4" Then
            If C.ComboLoadID = 1 Then
               Call InitEmpReceivableOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
          End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitThaiYear(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitThaiYear(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
                    If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadFormulaType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartItem(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitFormulaOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
                 If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadFormula(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitFormulaItemOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
                  Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
                 If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitJobStatus(m_Combos(C.ControlIndex))
               ElseIf C.ComboLoadID = 5 Then
               Call InitJobOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
                    
                       If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadJob(m_Combos(C.ControlIndex))
            End If
         End If
              
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadSerialNo(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitJobOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadFormulaType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartItem(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitFormulaOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitJobStatus(m_Combos(C.ControlIndex))
               ElseIf C.ComboLoadID = 5 Then
               Call InitJobOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
              
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadEmployee(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitJobStatus(m_Combos(C.ControlIndex))
               ElseIf C.ComboLoadID = 5 Then
               Call InitJobOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-9" Then
           If C.ComboLoadID = 1 Then
               Call LoadUnit(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
      End If
          
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-10" Then
           If C.ComboLoadID = 1 Then
               Call LoadUnit(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
      End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-11" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-12" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_12Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-13" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-20" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-13-1" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-14" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-14-0" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-14-1" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-14-2" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-15" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
                  
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-16" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-17" Then
           If C.ComboLoadID = 1 Then
               Call LoadProcess(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , PRTITEM_SET)
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport8_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-18" Then
           If C.ComboLoadID = 1 Then
               Call LoadFormulaType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_18Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-19" Then
           If C.ComboLoadID = 1 Then
               Call LoadParameterProcess(m_Combos(C.ControlIndex))
           ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_19Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-1" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-2" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-3.1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If

         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitARIntervalType(m_Combos(C.ControlIndex))
            End If
         End If
         
          If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadBank(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadBankBranch(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-6" Then
            If C.ComboLoadID = 1 Then
               Call InitPaymentType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportA_2_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-7-1" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-7" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-8" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitARIntervalType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , BANK_ACCOUNT)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportCashTx(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-12" Then
            If C.ComboLoadID = 1 Then
               Call InitReportCashTx(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-13" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-14" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-15" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-17" Then
            If C.ComboLoadID = 1 Then
               Call InitReportA_2_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " J-1-1" Or _
             trvMaster.SelectedItem.Key = ROOT_TREE & " J-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_1OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " BUY-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitAPChequeStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , PAY_TO)
            ElseIf C.ComboLoadID = 2 Then
               Call InitCheckLayout(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportAP_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportAP_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
            
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportAP_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitAPChequeStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      End If
   Next C
   Call EnableForm(Me, True)
End Sub
Private Sub LoadComboData2()
Dim C As CReportControl

   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
           If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitAPChequeStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-5" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-5-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportAP_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-8" Or trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If
        
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-6" Or _
         trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-6-1" Then
         
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportAP_1_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitAPChequeStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
          If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitPoType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitSupplierDocNoOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitBillingDocCloseApproved(m_Combos(C.ControlIndex), 1)
            End If
         End If
         
           If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitPoType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitSupplierDocNoOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitBillingDocApproved(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitBillingDocCloseApproved(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitPoType(m_Combos(C.ControlIndex))
            End If
         End If
         
        If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitPoType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitSupplierDocNoOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport1_11_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType2(m_Combos(C.ControlIndex), 2)
            End If
         End If
         

         
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport4_A_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AB-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " AB-1-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " AB-1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitPrType(m_Combos(C.ControlIndex))
            End If
         End If
      
      If trvMaster.SelectedItem.Key = ROOT_TREE & " AB-1-4" Then
            If C.ComboLoadID = 1 Then
               Call InitPrType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            End If
         End If
         
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitExportDocTypeSet(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadLayout(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport4_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
         
           If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 5 Then
'               Call InitARIntervalType(m_Combos(C.ControlIndex))
            End If
         End If
         
             If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 5 Then
'               Call InitARIntervalType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 5 Then
'               Call InitReportA_1_4Orderby(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 6 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
        If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 5 Then
'               Call InitARIntervalType(m_Combos(C.ControlIndex))
            End If
         End If
         
       If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-4-22" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            End If
         End If
                  
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-A-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_A_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-16" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_8Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-2-21" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReportA_2_21Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 10-0-2" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            End If
         End If
         
        If trvMaster.SelectedItem.Key = ROOT_TREE & " 10-0-3" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 10-0-4" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerGrade(m_Combos(C.ControlIndex))
            End If
         End If
      
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 9-1-1" Then
           If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2, , , , , "HEAD")
           ElseIf C.ComboLoadID = 2 Then
               Call InitLoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport9_1_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 9-2-1" Then
           If C.ComboLoadID = 1 Then
               Call InitLoadPartPayType(m_Combos(C.ControlIndex))
           ElseIf C.ComboLoadID = 2 Then
               Call InitLoadProcessType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport9_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
'         If trvMaster.SelectedItem.Key = ROOT_TREE & " 9-2-1-2" Then
'           If C.ComboLoadID = 1 Then
'               Call InitLoadPartPayType(m_Combos(C.ControlIndex))
'           ElseIf C.ComboLoadID = 2 Then
'               Call InitLoadProcessType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 3 Then
'               Call InitReport9_2_1Orderby(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 4 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
'            End If
'         End If
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 9-2-2" Then
            If C.ComboLoadID = 1 Then
               Call InitLoadPartPayType(m_Combos(C.ControlIndex))
           ElseIf C.ComboLoadID = 2 Then
               Call InitLoadProcessType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport9_2_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
'           If C.ComboLoadID = 1 Then
'               Call InitLoadPartPayType2(m_Combos(C.ControlIndex))
'           ElseIf C.ComboLoadID = 2 Then
'               Call InitLoadProcessType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 3 Then
'               Call InitReport9_2_1Orderby(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 4 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
'            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 9-3-1" Then
           If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport9_2_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-5-1-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 5-5-1-3" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitReport5_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 8 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 2-3-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartItem(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport2_3_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 2-3-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartItem(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport2_3_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 2-3-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartItem(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport2_3_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 2-3-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartItem(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport2_3_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
        If trvMaster.SelectedItem.Key = ROOT_TREE & " A-7-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_3OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-6" Then
            If C.ComboLoadID = 1 Then
               Call InitReport5_7OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReportAP_1_11_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " AP-1-11-8" Then
            If C.ComboLoadID = 1 Then
               Call InitReport5_7OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-7-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_3OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-7-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_3OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 9-3-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 9-3-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex), m_PartGroups)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2)
            ElseIf C.ComboLoadID = 4 Then
              Call InitLoadPartPayType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " A-1-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadBank(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadBankBranch(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
                Call LoadMaster(m_Combos(C.ControlIndex), , BANK_ACCOUNT)
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport5_2OrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
        If trvMaster.SelectedItem.Key = ROOT_TREE & " F-1-1-2" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 6 Then
'               Call InitReportF1_1Orderby(m_Combos(C.ControlIndex))
'            ElseIf C.ComboLoadID = 7 Then
'               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
        End If
         
     
   Next C
   Call EnableForm(Me, True)
End Sub

Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "", Optional KeySearch As String, Optional ShowNow As Boolean = True, Optional StrText As String = "")
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim ChIdx As Long
Dim C As CReportControl

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   ChIdx = m_Checks.Count + 1
  
   
   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
      txtGeneric(TxtIdx).SetKeySearch (KeySearch)
      txtGeneric(TxtIdx).Text = StrText
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx
      
      If Param1 = "FROM_DOC_DATE" Or Param1 = "FROM_DATE" Then
         If m_FromDate > 0 And ShowNow Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         ElseIf Not ShowNow Then
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      ElseIf Param1 = "TO_DOC_DATE" Or Param1 = "TO_DATE" Then
          If m_FromDate > 0 And ShowNow Then
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         ElseIf Not ShowNow Then
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         End If
      ElseIf Param1 = "TO_PAY_DATE" Then
         If m_ToRcp > 0 And ShowNow Then
            uctlGenericDate(DateIdx).ShowDate = m_ToRcp
         ElseIf Not ShowNow Then
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToRcp)
            uctlGenericDate(DateIdx).ShowDate = m_ToRcp
         End If
      ElseIf Param1 = "PRINT_DATE" Then
         If m_PrintDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_PrintDate
         Else
            uctlGenericDate(DateIdx).ShowDate = Now
         End If
      ElseIf Param1 = "BETWEEN_DATE" Then
         If m_DocDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_DocDate
         Else
            uctlGenericDate(DateIdx).ShowDate = Now
         End If
      ElseIf Param1 = "FROM_WEEK" Or Param1 = "FROM_WEEK_DATE" Then
           If m_FromWeek > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromWeek
           Else
            uctlGenericDate(DateIdx).ShowDate = getBeginDay("MON")
           End If
       ElseIf Param1 = "TO_WEEK_DATE" Or Param1 = "TO_SUP_DATE" Then
           If m_ToWeek > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToWeek
           Else
            uctlGenericDate(DateIdx).ShowDate = getBeginDay("MON") + 6 ' �� 1 �ѻ����
           End If
      ElseIf Param1 = "FROM_DUE_DATE" Or Param1 = "TO_DUE_DATE" Then
         uctlGenericDate(DateIdx).ShowDate = Now
      End If
   ElseIf ControlType = "LU" Then
'      Load uctlTextLookup(LkupIdx)
'      Call m_TextLookups.add(uctlTextLookup(LkupIdx))
'      C.ControlIndex = LkupIdx
    ElseIf ControlType = "CH" Then
      Load chkCommit(ChIdx)
      Call m_Checks.add(chkCommit(ChIdx))
      Call InitCheckBox(chkCommit(ChIdx), TextMsg)
      C.ControlIndex = ChIdx
   End If
   
   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.add(C)
   Set C = Nothing
End Sub

Private Sub InitReport1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���͡����"))

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͼ����"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "GROUP_ID", "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���͡����"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͼ����"))
   
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).WIDTH, True, "", 1, "GROUP_ID", "GROUP_NAME")
'   Call LoadControl("L", lblGeneric(0).WIDTH, True, GetTextMessage("TEXT-KEY71"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub


Private Sub InitReport3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_GRADE", "CUSTOMER_GRADE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "NO", , "SHOW_NO")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�����١���", , "SHOW_CODE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "������", , "SHOW_TYPE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�дѺ", , "SHOW_LEVEL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�������", , "SHOW_ADDRESS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ôԵ�ѹ", , "SHOW_CREDIT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ǧ�Թ", , "SHOW_CREDIT_LIMIT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ѡ�ҹ���", , "SHOW_SALE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ѹ�֡ŧ File", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport3_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_GRADE", "CUSTOMER_GRADE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ѹ�֡ŧ File", , "PRINT_TO_FILE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ���������"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ͫѾ���������"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Ѿ �"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���;�ѡ�ҹ"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ�ž�ѡ�ҹ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport3_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ����Ź��"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���;�ѡ�ҹ����Ź��"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ�ž�ѡ�ҹ����Ź��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReport4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "UNIT_COUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("˹��¹Ѻ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "����Ѻ��Ǩ�ͺ", , "FOR_CHECK")
   Call LoadControl("CH", cboGeneric(0).Width, True, "੾�����������¡��ԡ", , "CANCEL_FLAG")
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "COLUMN2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�.���ҧ ������� 2"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "COLUMN3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�.���ҧ ������� 3"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "COLUMN4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�.���ҧ ������� 4"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "COLUMN5")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�.���ҧ ������� 5"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "COLUMN6")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�.���ҧ ������� 6"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "COLUMN7")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�.���ҧ ������� 7"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport9_3_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "LOT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("LOT ��ü�Ե"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 4, "DOCUMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ú�è�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReport4_5_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����ѵ�شԺ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����ѵ�شԺ"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_5_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����ѵ�شԺ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����ѵ�شԺ"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����������´(15)"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOCATION_ID1", "LOCATION_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡʶҹ���Ѵ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID2", "LOCATION_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧ʶҹ���Ѵ��"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_12_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "IMPORT_DOC_TYPE", "IMPORT_DOC_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������͡���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("� �ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "UNIT_COUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("˹��¹Ѻ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_6OLD()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SUPPLIER_TYPE", "SUPPLIER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_A_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SUPPLIER_TYPE", "SUPPLIER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_A_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SUPPLIER_TYPE", "SUPPLIER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport4_A_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "SUPPLIER_TYPE", "SUPPLIER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReport4_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ʴ�/�ػ�ó�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ʴ�/�ػ�ó�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "EXPORT_DOC_TYPE", "EXPORT_DOC_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������͡���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "DEPARTMENT_ID", "DEPARTMENT_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("Ἱ�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
'InitReport4_11_1
Private Sub InitReport4_11_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ʴ�/�ػ�ó�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ʴ�/�ػ�ó�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "EXPORT_DOC_TYPE", "EXPORT_DOC_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������͡���"))

  ' 3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "DEPARTMENT_ID", "DEPARTMENT_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("Ἱ�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport4_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport5_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "SHRINK", , "SHRINK", , "0.5")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("SHRINK"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "BAG_VALUE", , "BAG_VALUE", , "260")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҷا/�ѹ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "OH_GRAIN", , "OH_GRAIN", , "850")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("OH (���)/�ѹ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "OH_POWDER", , "OH_POWDER", , "650")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("OH (��)/�ѹ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "OH_NC_GRAIN", , "OH_NC_GRAIN", , "1000")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("OH ��⪤ (���)/�ѹ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "OH_NC_POWDER", , "OH_NC_POWDER", , "850")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("OH ��⪤ (��)/�ѹ"))
   
   If TempKey = "Root F-1-8" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������١���", , "SHOW_CUS_NAME")
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���º�Ţ��", , "SHOW_ORDER_BILL")
   ElseIf TempKey = "Root F-1-9" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL_PART")
   End If
'   If TempKey = "5-2-1" Then
'      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�Ŵ˹��/����˹��", , "CREDIT")
'   End If
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportA_7_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   If TempKey = "Root A-7-2" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������Թ���", , "SHOW_PART_NO_FLAG")
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡�÷���Ѻ����¹�Ҥ�", , "SHOW_EDIT_PRICE_FLAG")
   End If
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportA_7_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))


   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
    Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ������Թ���", , "SHOW_PART_NO_FLAG")
    Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡�÷���Ѻ����¹�Ҥ�", , "SHOW_EDIT_PRICE_FLAG")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAP_1_11_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))


   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´��Ң���", , "FLAG_DELIVERY_DETAIL")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   If TempKey = "5-2-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�Ŵ˹��/����˹��", , "CREDIT")
   End If
   If TempKey = "Root " & "5-2-1-1" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�ʶҹ���Ѵ��", , "FLAG_DELIVERY_CUS")
   End If
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport5_1_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
'
'   '2 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
   

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
'
'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_5_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
   
'1 =============================
 Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "POSITION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 8, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ Collumn", , "SUMMARY_COLLUMN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ Row", , "SUMMARY_ROW")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_1_0_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���Ң���", , "TRANSPORT_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
'   If TempKey = "Root " & "5-2-1-1" Then
     Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�Ŵ˹��/����˹��", , "CREDIT")
     Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�ʶҹ���Ѵ��", , "FLAG_DELIVERY_CUS")
'   End If
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportA_1_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

      Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BANKS", "BANKS_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҥ��"))
      
      Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "BANK_BRANCHS", "BANK_BRANCHS_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ң�"))
      
      Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ACCOUNT_ID", "ACCOUNT_ID_NAME")
      Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ���ѭ��"))
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���¡���Թ���", , "FLAG_SHOW_PART_NO")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�ʶҹ���Ѵ��", , "FLAG_DELIVERY_CUS")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

'   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_CREDIT")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ôԵ"))
'
'   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_CREDIT")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ôԵ"))


   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�Ŵ˹��/����˹��", , "CREDIT")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "POSITION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportA_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "POSITION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ�ʶҹ���Ѵ��", , "FLAG_DELIVERY_CUS")
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))
'
'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRTITEM_SET_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("૵�ͧ�Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "POSITION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_5_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PRTITEM_SET_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("૵�ͧ�Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))
   
   Call LoadControl("C", cboGeneric(0).Width, False, "", 5, "MODE", "MODE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ť��/���˹ѡ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportAP_1_11_6() '
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ��袹��"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ��袹��"))

Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TRUCK_NO", , "TRUCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����¹ö"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE_TRANSPORT", , "SUPPLIER_CODE_TRANSPORT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ �ö����"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ����͡���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "��Ẻ�����", , "USE_FORM_FLAG", "��Ẻ�����")
   Call LoadControl("CH", chkCommit(0).Width, True, "���١���", , "CUSTOMER_ONLY_FLAG", "੾��ö����")
   Call LoadControl("CH", chkCommit(0).Width, True, "��ҧ����觺ѭ��", , "INVOICE_FLAG", "��ҧ���")
   Call LoadControl("CH", chkCommit(0).Width, True, "��ػ��è��¤�Ң���", , "SUMMARY_DOC_FLAG", "��ػ��è��¤�Ң���")
   Call LoadControl("CH", chkCommit(0).Width, True, "��Ӥѭ����", , "PAYMENT_VOUCHER_FLAG", "��Ӥѭ����")
   Call LoadControl("CH", chkCommit(0).Width, True, "��Ӥѭ�͹�ѭ��", , "ACCOUNT_TRANSFER_FLAG", "��Ӥѭ�͹�ѭ��")
   Call LoadControl("CH", chkCommit(0).Width, True, "���������ѡ��Ң���", , "COLUMN_DEDUC_FLAG", "���������ѡ��Ң���")
   Call LoadControl("CH", chkCommit(0).Width, True, "��������������Ң���", , "ACCOUNT_ADD_FLAG", "��������������Ң���")
   Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ������١������ҧ���", , "SHOW_CUS_NAME_FLAG", "�ʴ������١������ҧ���")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAP_1_11_8() '
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ����͡���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����͡���"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TRUCK_NO", , "TRUCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����¹ö"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE_TRANSPORT", , "SUPPLIER_CODE_TRANSPORT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ �ö����"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ����͡���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "���١���", , "CUSTOMER_ONLY_FLAG", "੾��ö����")
   Call LoadControl("CH", chkCommit(0).Width, True, "��ҧ����觺ѭ��", , "INVOICE_FLAG", "��ҧ���")
   Call LoadControl("CH", chkCommit(0).Width, True, "���������ѡ��Ң���", , "COLUMN_DEDUC_FLAG", "���������ѡ��Ң���")
   Call LoadControl("CH", chkCommit(0).Width, True, "��������������Ң���", , "ACCOUNT_ADD_FLAG", "��������������Ң���")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_6() '
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TRUCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����¹ö"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_1_3() '
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "BANK_ACCOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ش�ѭ��"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))

   '8 =============================
   Call LoadControl("T", cboGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ �"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʫѾ� �"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean

   If LastKey = Node.Key Then
      Exit Sub
   End If
   
   Status = True
   QueryFlag = False
   
   Call UnloadAllControl
   cmdOK.Enabled = True
   cmdSelCus.Visible = False
   cmdSelCusGrade.Visible = False
   TempKey = ""
   
   If Node.Key = ROOT_TREE & " 1-1" Then
      If Not VerifyAccessRight("ADMIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport1_1
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      If Not VerifyAccessRight("ADMIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport1_2
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      If Not VerifyAccessRight("ADMIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport1_3
   ElseIf Node.Key = ROOT_TREE & " 2-1" Then
      If Not VerifyAccessRight("PACKAGE_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport2_1
   ElseIf Node.Key = ROOT_TREE & " 2-2" Then
      If Not VerifyAccessRight("PACKAGE_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport2_2
   ElseIf Node.Key = ROOT_TREE & " 2-3-1" Then
      If Not VerifyAccessRight("PACKAGE_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport2_3_1
   ElseIf Node.Key = ROOT_TREE & " 2-3-2" Then
      If Not VerifyAccessRight("PACKAGE_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReport2_3_2
   ElseIf Node.Key = ROOT_TREE & " 2-3-3" Then
      If Not VerifyAccessRight("PACKAGE_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport2_3_1
   ElseIf Node.Key = ROOT_TREE & " 2-3-4" Then
      If Not VerifyAccessRight("PACKAGE_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReport2_3_2
   ElseIf Node.Key = ROOT_TREE & " 3-1" Then
      If Not VerifyAccessRight("MAIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport3_1
   ElseIf Node.Key = ROOT_TREE & " 3-1-1" Then
      If Not VerifyAccessRight("MAIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport3_1
   ElseIf Node.Key = ROOT_TREE & " 3-1-2" Then
      If Not VerifyAccessRight("MAIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport3_1_2
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      If Not VerifyAccessRight("MAIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         cmdOK.Enabled = False
         Exit Sub
      End If
      Call InitReport3_2
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      If Not VerifyAccessRight("MAIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport3_3
   ElseIf Node.Key = ROOT_TREE & " 4-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         cmdOK.Enabled = False
         Exit Sub
      End If
      Call InitReport4_1
   ElseIf Node.Key = ROOT_TREE & " 4-2" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         cmdOK.Enabled = False
         Exit Sub
      End If
      Call InitReport4_2
   ElseIf Node.Key = ROOT_TREE & " 4-3" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         cmdOK.Enabled = False
         Exit Sub
      End If
      Call InitReport4_3
   ElseIf Node.Key = ROOT_TREE & " 4-3-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         cmdOK.Enabled = False
         Exit Sub
      End If
      Call InitReport4_3
   ElseIf Node.Key = ROOT_TREE & " 4-4" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_4
   ElseIf Node.Key = ROOT_TREE & " 4-4-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_4
   ElseIf Node.Key = ROOT_TREE & " 4-5" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_4
   ElseIf Node.Key = ROOT_TREE & " 4-5-1-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_1_1
   ElseIf Node.Key = ROOT_TREE & " 4-5-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_1
   ElseIf Node.Key = ROOT_TREE & " 4-5-2" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_1
   ElseIf Node.Key = ROOT_TREE & " 4-5-3" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_1
   ElseIf Node.Key = ROOT_TREE & " 4-5-4" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_1
   ElseIf Node.Key = ROOT_TREE & " 4-5-6" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_6
   ElseIf Node.Key = ROOT_TREE & " 4-5-6-2" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_6
   ElseIf Node.Key = ROOT_TREE & " 4-5-7" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_6
   ElseIf Node.Key = ROOT_TREE & " 4-5-8" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_6
   ElseIf Node.Key = ROOT_TREE & " 4-5-9" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_5_6
   ElseIf Node.Key = ROOT_TREE & " 4-6" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_6
   ElseIf Node.Key = ROOT_TREE & " 4-7" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_7
   ElseIf Node.Key = ROOT_TREE & " 4-7-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_7
   ElseIf Node.Key = ROOT_TREE & " 4-12-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_12_1
   ElseIf Node.Key = ROOT_TREE & " 4-12-2" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_11
   ElseIf Node.Key = ROOT_TREE & " 4-12-3" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_11
   ElseIf Node.Key = ROOT_TREE & " 4-12-4" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_12_1
   ElseIf Node.Key = ROOT_TREE & " 4-12-5" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_11
      
       ElseIf Node.Key = ROOT_TREE & " 4-12-7" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_11_1
   ElseIf Node.Key = ROOT_TREE & " 4-12-6" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_12_1
   ElseIf Node.Key = ROOT_TREE & " 4-8" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_8
   ElseIf Node.Key = ROOT_TREE & " 4-8-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_8
   ElseIf Node.Key = ROOT_TREE & " 4-9" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_9
   ElseIf Node.Key = ROOT_TREE & " 4-10" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_10
   ElseIf Node.Key = ROOT_TREE & " 4-A-1" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_A_1
   ElseIf Node.Key = ROOT_TREE & " 4-A-2" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_A_2
   ElseIf Node.Key = ROOT_TREE & " 4-A-3" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_9
   ElseIf Node.Key = ROOT_TREE & " 4-A-4" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_A_2
   ElseIf Node.Key = ROOT_TREE & " 4-A-5" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_A_5
   ElseIf Node.Key = ROOT_TREE & " 4-A-6" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_A_6
   ElseIf Node.Key = ROOT_TREE & " 4-11" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_11
   ElseIf Node.Key = ROOT_TREE & " 4-12" Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport4_11
   ElseIf Node.Key = ROOT_TREE & " 5-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReport5_1
    ElseIf Node.Key = ROOT_TREE & " 5-1-0" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-1-0-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1_0_1
   ElseIf Node.Key = ROOT_TREE & " 5-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-2-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = "5-2-1"
      Call InitReport5_1
      TempKey = ""
   ElseIf Node.Key = ROOT_TREE & " 5-2-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReport5_1_1
   ElseIf Node.Key = ROOT_TREE & " 5-2-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1_2
   ElseIf Node.Key = ROOT_TREE & " 5-2-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_5
   ElseIf Node.Key = ROOT_TREE & " 5-5-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_5_1
   ElseIf Node.Key = ROOT_TREE & " 5-5-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_5_1_2
   ElseIf Node.Key = ROOT_TREE & " 5-5-1-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_5_1_2
   ElseIf Node.Key = ROOT_TREE & " 5-5-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_5_2
   ElseIf Node.Key = ROOT_TREE & " 5-6" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_6
   ElseIf Node.Key = ROOT_TREE & " A-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_6
   ElseIf Node.Key = ROOT_TREE & " A-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_6
   ElseIf Node.Key = ROOT_TREE & " A-1-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_1_3
   ElseIf Node.Key = ROOT_TREE & " A-1-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_1_4
   ElseIf Node.Key = ROOT_TREE & " A-1-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReportA_1_5
   ElseIf Node.Key = ROOT_TREE & " 5-6-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_6_1
   ElseIf Node.Key = ROOT_TREE & " 5-7" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_7
   ElseIf Node.Key = ROOT_TREE & " 5-8" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_8
   ElseIf Node.Key = ROOT_TREE & " 5-9" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_9
   ElseIf Node.Key = ROOT_TREE & " 5-10" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_10
   ElseIf Node.Key = ROOT_TREE & " 5-11" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_11
   ElseIf Node.Key = ROOT_TREE & " 5-12" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_12
   ElseIf Node.Key = ROOT_TREE & " 5-13" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_13
   ElseIf Node.Key = ROOT_TREE & " 5-14" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_14
   ElseIf Node.Key = ROOT_TREE & " 5-15" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_15
   ElseIf Node.Key = ROOT_TREE & " F-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1_11
   ElseIf Node.Key = ROOT_TREE & " F-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-6" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-7" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " F-1-8" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReport5_1_3
   ElseIf Node.Key = ROOT_TREE & " F-1-9" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReport5_1_3
   ElseIf Node.Key = ROOT_TREE & " 6-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport6_1
   ElseIf Node.Key = ROOT_TREE & " 6-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport6_4
   ElseIf Node.Key = ROOT_TREE & " 6-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport6_5
  ElseIf Node.Key = ROOT_TREE & " A-7-1" Then
  
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReportA_7_1
   ElseIf Node.Key = ROOT_TREE & " A-7-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      TempKey = Node.Key
      Call InitReportA_7_1
   ElseIf Node.Key = ROOT_TREE & " A-7-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_7_3
   ElseIf Node.Key = ROOT_TREE & " 8-1" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-2" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_2
   ElseIf Node.Key = ROOT_TREE & " 8-3" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_3
     ElseIf Node.Key = ROOT_TREE & " 8-4" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_4
   ElseIf Node.Key = ROOT_TREE & " 8-5" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_5
   ElseIf Node.Key = ROOT_TREE & " 8-6" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-7" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_3
   ElseIf Node.Key = ROOT_TREE & " 8-8" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_8
   ElseIf Node.Key = ROOT_TREE & " 8-9" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_9
   ElseIf Node.Key = ROOT_TREE & " 8-10" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_10
   ElseIf Node.Key = ROOT_TREE & " 8-11" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-12" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_12
   ElseIf Node.Key = ROOT_TREE & " 8-13" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-13-1" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-14" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-14-0" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-14-1" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-14-2" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-15" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-16" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 8-17" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13_1
   ElseIf Node.Key = ROOT_TREE & " 8-18" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_18
   ElseIf Node.Key = ROOT_TREE & " 8-19" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_19
   ElseIf Node.Key = ROOT_TREE & " 8-20" Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport8_13
   ElseIf Node.Key = ROOT_TREE & " 9-1-1" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_1_1
   ElseIf Node.Key = ROOT_TREE & " 9-2-1" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_2_1
'   ElseIf Node.Key = ROOT_TREE & " 9-2-1-2" Then
'      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
'         cmdOK.Enabled = False                                                                                                                                                               '''''''''
'         Exit Sub
'      End If
'      Call InitReport9_2_1
  ElseIf Node.Key = ROOT_TREE & " 9-2-2" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_2_2
   ElseIf Node.Key = ROOT_TREE & " 9-3-1" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_3_1
   ElseIf Node.Key = ROOT_TREE & " 9-3-2" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_2_1
   ElseIf Node.Key = ROOT_TREE & " 9-3-3" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_3_3
   ElseIf Node.Key = ROOT_TREE & " 9-3-4" Then
      If Not VerifyAccessRight("INVENTORY-WH_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReport9_3_3
   ElseIf Node.Key = ROOT_TREE & " A-2-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_1
   ElseIf Node.Key = ROOT_TREE & " A-2-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_2
   ElseIf Node.Key = ROOT_TREE & " A-2-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_3
   ElseIf Node.Key = ROOT_TREE & " A-2-3.1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_3
   ElseIf Node.Key = ROOT_TREE & " A-2-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4
   ElseIf Node.Key = ROOT_TREE & " A-2-4-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4
    ElseIf Node.Key = ROOT_TREE & " A-2-4-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4_2
  ElseIf Node.Key = ROOT_TREE & " A-2-4-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4_4
   ElseIf Node.Key = ROOT_TREE & " A-2-4-22" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4_22
   ElseIf Node.Key = ROOT_TREE & " A-2-4-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4_3
      ElseIf Node.Key = ROOT_TREE & " A-2-4-6" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4_6
     ElseIf Node.Key = ROOT_TREE & " A-2-4-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4_5
   ElseIf Node.Key = ROOT_TREE & " A-2-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_5
   ElseIf Node.Key = ROOT_TREE & " A-2-6" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_6
   ElseIf Node.Key = ROOT_TREE & " A-2-7" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_7
   ElseIf Node.Key = ROOT_TREE & " A-2-7-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_7
   ElseIf Node.Key = ROOT_TREE & " A-2-8" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_8
   ElseIf Node.Key = ROOT_TREE & " A-2-17" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_8
   ElseIf Node.Key = ROOT_TREE & " A-2-18" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_18
   ElseIf Node.Key = ROOT_TREE & " A-2-9" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_4
   ElseIf Node.Key = ROOT_TREE & " A-2-11" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_11
   ElseIf Node.Key = ROOT_TREE & " A-2-12" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_12
   ElseIf Node.Key = ROOT_TREE & " A-2-13" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_13
   ElseIf Node.Key = ROOT_TREE & " A-2-14" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_14
   ElseIf Node.Key = ROOT_TREE & " A-2-15" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_15
   ElseIf Node.Key = ROOT_TREE & " A-2-16" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_16
   ElseIf Node.Key = ROOT_TREE & " A-2-21" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportA_2_21
   ElseIf Node.Key = ROOT_TREE & " J-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportJ_1_1
   ElseIf Node.Key = ROOT_TREE & " J-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportJ_1_2
   ElseIf Node.Key = ROOT_TREE & " BUY-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportBuy_1_1
   ElseIf Node.Key = ROOT_TREE & " AP-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_1
   ElseIf Node.Key = ROOT_TREE & " AP-1-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_2
   ElseIf Node.Key = ROOT_TREE & " AP-1-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_3
   ElseIf Node.Key = ROOT_TREE & " AP-1-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_3
   ElseIf Node.Key = ROOT_TREE & " AP-1-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_5
   ElseIf Node.Key = ROOT_TREE & " AP-1-5-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_5
   ElseIf Node.Key = ROOT_TREE & " AP-1-6" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_5
   ElseIf Node.Key = ROOT_TREE & " AP-1-6-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_5
   ElseIf Node.Key = ROOT_TREE & " AP-1-7" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_7
   ElseIf Node.Key = ROOT_TREE & " AP-1-8" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_1
   ElseIf Node.Key = ROOT_TREE & " AP-1-9" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_1
   ElseIf Node.Key = ROOT_TREE & " AP-1-10" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_1
   ElseIf Node.Key = ROOT_TREE & " AP-1-11" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_1
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-2" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_2
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_3
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_3
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-5" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_5
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-6" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_6
    ElseIf Node.Key = ROOT_TREE & " AP-1-11-7" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_7
   ElseIf Node.Key = ROOT_TREE & " AP-1-11-8" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAP_1_11_8
   ElseIf Node.Key = ROOT_TREE & " AB-1-1" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAB_1_1
'   ElseIf Node.Key = ROOT_TREE & " AB-1-2" Then
'      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
'         cmdOK.Enabled = False                                                                                                                                                               '''''''''
'         Exit Sub
'      End If
'      Call InitReportAB_1_1
   ElseIf Node.Key = ROOT_TREE & " AB-1-3" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAB_1_3
   ElseIf Node.Key = ROOT_TREE & " AB-1-4" Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
      Call InitReportAB_1_4
    ElseIf Node.Key = ROOT_TREE & " 6-1-1" Or Node.Key = ROOT_TREE & " 6-1-2" Then
      If Not VerifyAccessRight("PLANNING_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               ''''''''
         Exit Sub
      End If
      Call InitReport6_1_1
   ElseIf Node.Key = ROOT_TREE & " 6-0-1" Then
      If Not VerifyAccessRight("PLANNING_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               ''''''''
         Exit Sub
      End If
      Call InitReport6_0_1
   ElseIf Node.Key = ROOT_TREE & " 10-0-1" Then
      If Not VerifyAccessRight("COMMISSION_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               ''''''''
         Exit Sub
      End If
      Call InitReport10_1
   ElseIf Node.Key = ROOT_TREE & " 10-0-2" Then
      If Not VerifyAccessRight("COMMISSION_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               ''''''''
         Exit Sub
      End If
      Call InitReport10_2
   ElseIf Node.Key = ROOT_TREE & " 10-0-3" Then
      If Not VerifyAccessRight("COMMISSION_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               ''''''''
         Exit Sub
      End If
       Call InitReport10_3
   ElseIf Node.Key = ROOT_TREE & " 10-0-4" Then
      If Not VerifyAccessRight("COMMISSION_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               ''''''''
         Exit Sub
      End If
       Call InitReport10_4
    End If
End Sub

Private Sub InitReportA_2_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportA_2_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PAYMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ê����Թ"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "�ʹ¡�ҹѺ�ҡ����͹", , "BALANCE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
'   Call LoadControl("CH", chkCommit(0).Width, True, "�ʹ¡�ҹѺ�ҡ����͹", , "BALANCE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "BANK_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҥ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BANK_BRANCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҢҸ�Ҥ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʾ�ѡ�ҹ"))
   
    '2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧���ʾ�ѡ�ҹ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE", "CUSTOMER_GRADE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 5, "INTERVAL_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ٻẺ��ǧ����˹��"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թ���", , "OVERDUE_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "����ʹ� DUE", , "SUM_INDUE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportA_2_4_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 150
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 5, "INTERVAL_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ٻẺ��ǧ����˹��"))
   
          
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թ���", , "OVERDUE_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "����ʹ� DUE", , "SUM_INDUE_FLAG")
'   Call LoadControl("CH", chkCommit(0).Width, True, "����ʴ��ʹ� DUE", , "HIDE_DUE_FLAG")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportA_2_4_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 150
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 5, "INTERVAL_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ٻẺ��ǧ����˹��"))
   
          
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թ���", , "OVERDUE_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "����ʹ� DUE", , "SUM_INDUE_FLAG")
'   Call LoadControl("CH", chkCommit(0).Width, True, "����ʴ��ʹ� DUE", , "HIDE_DUE_FLAG")
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReportA_2_4_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 150
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 5, "INTERVAL_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ٻẺ��ǧ����˹��"))
   
          
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թ���", , "OVERDUE_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "����ʹ� DUE", , "SUM_INDUE_FLAG")
'   Call LoadControl("CH", chkCommit(0).Width, True, "����ʴ��ʹ� DUE", , "HIDE_DUE_FLAG")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportA_2_4_22()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
''   '1 =============================
''   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
''
''
''   '1 =============================
''   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
''
''   '1 =============================
''   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
''    '2 =============================
''   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "EMP_CODE", , "EMP_CODE")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
''   '3 =============================
''   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
''
''   '3 =============================
''   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE", "CUSTOMER_GRADE_NAME")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
      
    
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թǧ�Թ����ǧ�ѹ", , "OVERDUE_FLAG")
    
    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
    Call LoadControl("CH", chkCommit(0).Width, True, "��ػ", , "SUM_FLAG")
    
'    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���õԴ����� Column", , "SHOW_COLUMN_FLAG")
'
'    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���õԴ����� Row", , "SHOW_ROW_FLAG")
'    Call LoadControl("CH", chkCommit(0).Width, True, "੾�о�ѡ�ҹ����ѧ������͡", , "SHOW_EMP_RESIGN_FLAG")
    
   cmdSelCusGrade.Top = 4000
   cmdSelCusGrade.Left = chkCommit(0).Left '3220
   cmdSelCusGrade.Visible = True
   
   cmdSelCus.Top = cmdSelCusGrade.Top + cmdSelCusGrade.HEIGHT + 30
   cmdSelCus.Left = cmdSelCusGrade.Left
   cmdSelCus.Visible = True

   Call ShowControl
   Call LoadComboData2

End Sub
Private Sub InitReportA_2_4_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
    '2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE", "CUSTOMER_GRADE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
      
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թǧ�Թ����ǧ�ѹ", , "OVERDUE_FLAG")
    
    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
    
    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���õԴ����� Column", , "SHOW_COLUMN_FLAG")
    
    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���õԴ����� Row", , "SHOW_ROW_FLAG")
    Call LoadControl("CH", chkCommit(0).Width, True, "੾�о�ѡ�ҹ����ѧ������͡", , "SHOW_EMP_RESIGN_FLAG")
    'EMP_RESIGN_FLAG
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportA_2_4_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
    '2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "EMP_CODE", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
      
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾���Թǧ�Թ����ǧ�ѹ", , "OVERDUE_FLAG")
    
    Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
    
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
      '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ�����͹", , "SUM_MONTH")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ�����͹����������´", , "SUM_MONTH_DETAIL")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
      
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100

'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   uctlGenericDate(0).Enable = False

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʹ��������� 0", , "INCLUDE_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���;�ѡ�ҹ"))

   '1 =============================
   Call LoadControl("T", cboGeneric(0).Width, True, "", , "EMP_LAST_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "EMP_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ�"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "EMP_SEX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���;�ѡ�ҹ"))

   '1 =============================
   Call LoadControl("T", cboGeneric(0).Width, True, "", , "EMP_LAST_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LEND_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ����"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "LENDER")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʾ�ѡ�ҹ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���;�ѡ�ҹ"))

   '1 =============================
   Call LoadControl("T", cboGeneric(0).Width, True, "", , "EMP_LAST_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʡ��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���˹�"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "FROM_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��͹"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "FROM_YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 4, "TO_MONTH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��͹"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 5, "TO_YEAR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧��"))
  
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FEATURE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FEATURE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "FEATURE_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))


   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "UNIT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("˹����Ѵ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport2_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SOC_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţᾤࡨ"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SOC_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����ᾤࡨ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport2_3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ����С��"))
   
      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "BETWEEN_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ����ռ�"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "WORKS_PRICE_CODE", , "WORKS_PRICE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţྨᡨ"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 1, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/��ԡ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReport2_3_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ����С��"))
   
      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "BETWEEN_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ǧ�ѹ����ռ�"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "WORKS_PRICE_CODE", , "WORKS_PRICE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţྨᡨ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   If TempKey = "Root 2-3-2" Then
      Call LoadControl("CH", chkCommit(0).Width, True, "�����Ź���� IMPORT ����", , "FOR_IMPORT_FLAG")
   End If
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReport8_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FORMULA_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�ٵ�"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "FORMULA_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������´�ٵ�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "FORMULA_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ٵ�"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FORMULA_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������ҧ�ٵ�"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APPROVED_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ҧ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "FORMULA_ITEM")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ե�ѳ�������ҧ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "FORMULA_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�ٵ�"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub


Private Sub InitReport8_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "JOB_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�ҹ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "JOB_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������´�ҹ"))
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "BATCH_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţặ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "JOB_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�����觼�Ե"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "START_JOB")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���������ҹ"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FINISH_JOB")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������稧ҹ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APPROVED_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("͹��ѵ���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "RESPONSE_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ѻ�Դ�ͺ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PROCESS_ID", "PROCESS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "JOB_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹЧҹ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "JOB_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�ҹ"))
   
  
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "SERIAL_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ�Թ���"))
   
 '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
  
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport8_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "JOB_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ㺻����Թ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "JOB_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������´"))
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "BATCH_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţặ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "JOB_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������Թ"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "START_JOB")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���������ҹ"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FINISH_JOB")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������稧ҹ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APPROVED_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("͹��ѵ���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "RESPONSE_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ѻ�Դ�ͺ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PROCESS_ID", "PROCESS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "JOB_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹЧҹ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport8_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("� �ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "UNIT_COUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("˹��¹Ѻ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
'   Call LoadControl("T", txtGeneric(0).Width, True, "", , "JOB_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ㺻����Թ"))

   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width, True, "", , "JOB_DESC")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������´"))
'1 =============================
'   Call LoadControl("T", txtGeneric(0).Width, True, "", , "BATCH_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţặ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FINISH_JOB")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ������稧ҹ"))

   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APPROVED_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("͹��ѵ���"))
   
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "RESPONSE_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ѻ�Դ�ͺ��"))
   
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PROCESS_ID", "PROCESS_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "JOB_STATUS")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹЧҹ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   '4 =============================
   Call LoadControl("CH", chkCommit(0).Width, True, "", , "SUCESS_FLAGE")
  ' Call LoadControl("L", lblGeneric(0).Width, True, MapText("'�ҹ����"))
   
   
   Call ShowControl
   Call LoadComboData
  
End Sub

Private Sub InitReport8_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PROCESS_ID", "PROCESS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PROCESS_ID", "PROCESS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PROCESS_ID", "PROCESS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_GROUP", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_13_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PROCESS_ID", "PROCESS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "PRTITEM_SET_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ش�������Թ���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "FORMULA_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ٵ�"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportA_2_16()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   uctlGenericDate(0).Enable = False
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport8_19()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PARAMETER_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
  Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport9_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
  
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ����Ե"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HEAD_PACK_NO", "HEAD_PACK_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����ͧ��è�"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "DOCUMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ú�è�"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
''
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�������", , "NOT_SHOW")
      
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport9_2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
  
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�����������"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DOCUMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ú�è�"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOAD_FLAG")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "LOT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("LOT ��ü�Ե"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_CUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ�����١���(15)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ�����Թ���(25)"))
   
''
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�������", , "NOT_SHOW")
      
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport9_2_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
  
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�����������"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�����������"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "DOCUMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ú�è�"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOAD_FLAG")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "LOT_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("LOT ��ü�Ե"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_CUS_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ�����١���(15)"))
'
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ�����Թ���(25)"))
'
'''
'   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�������", , "NOT_SHOW")
      
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport9_3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
  
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
''   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_CUS_NAME")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ�����١���(15)"))
''
''   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_NAME")
''   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ�����Թ���(15)"))
''
''''
''   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ�������", , "NOT_SHOW")
      
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportJ_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("� �ѹ���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportJ_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "JV_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Ţ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������͡���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportBuy_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportAP_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ�������"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ�������"))
        
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "CHEQUE_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))

   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReportAP_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, False, "", , "CHEQUE_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Ţ�����"))
              
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_ID", "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Ѻ"))
           
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CHEQUE_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӹǹ�Թ"))
           
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "CHEQUE_LAYOUT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ٻẺ��"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FONT_SIZE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��Ҵ���ͺѭ��"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "����ͧŧ�ѹ���", , "DATE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportAP_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   
'   1 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
'
'   1 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "�Ѵ˹��������͡��", , "EFFECTIVE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportAP_1_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
         
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   
'   1 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
'
'   1 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE", "SUPPLIER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�Ѵ˹��������͡��", , "EFFECTIVE_FLAG")

   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReportAP_1_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
         
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   
'   1 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
'
'   1 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportAP_1_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ��� PO"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ��� PO"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAP_1_11_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѵ�شԺ"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE") ' ����ѹ���    "�ҡ�ѹ��� PO"  �� �ҡ�ѹ����Ѻ PO � CReportAP017
   Call LoadControl("L", lblGeneric(0).Width, True, MapText(" �ѹ������� "))
   
      
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PO_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ PO"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "PO_CLOSE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ûԴ PO"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ѵ�شԺ(25)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������Ǫ��ͼ����(25)"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ�˹���", , "SHOW_UNIT_NAME_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "ʶҹ��͡���", , "SHOW_STATUS_PO_FLAG")
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReportAP_1_11_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ��� PO"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ��� PO"))
   
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PO_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ PO"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "PO_APPROVED")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���͹��ѵ� PO"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "PO_CLOSE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��ûԴ PO"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAP_1_11_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѵ�شԺ"))
   
      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE") ' ����ѹ���    "�ҡ�ѹ��� PO"  �� �ҡ�ѹ����Ѻ PO � CReportAP017
   Call LoadControl("L", lblGeneric(0).Width, True, MapText(" �ѹ������� "))
      
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PO_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ PO"))
      
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������ѵ�شԺ(25)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������Ǫ��ͼ����(25)"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ�˹���", , "SHOW_UNIT_NAME_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "��ػ", , "SUMMARY_MODE")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAP_1_11_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
'   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����ѵ�شԺ"))
'
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport5_1_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������Թ���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��������â��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport4_A_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "��Ѻ����ѵ�شԺ", , "DOCUMENT_TYPE1")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��Ѻ�����ʴ��ػ�ó�", , "DOCUMENT_TYPE19")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��Ѻ��Ҩ����͡��ʴ��ػ�ó�", , "DOCUMENT_TYPE20")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��Ѻ��ҷ����", , "DOCUMENT_TYPE23")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAB_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
         
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAB_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
         
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����ѵ�شԺ"))
   
    '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "RO_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ RO"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�����"))
   
   'Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������´", , "SHOW_DETAIL")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportAB_1_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
         
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʼ����"))

   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_NO", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����Թ���/�ѵ�شԺ"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�Թ���/�ѵ�شԺ"))
   
    '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "RO_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ RO"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����������"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "AMPHUR")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����/ࢵ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "PROVINCE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѧ��Ѵ"))
   
   Call ShowControl
   Call LoadComboData2
End Sub

Private Sub InitReportA_2_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   uctlGenericDate(0).Enable = False

'1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, False, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport6_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����ѵ�شԺ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����ѵ�شԺ"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���¡��"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("����������´(15)"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_0_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_WEEK")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("Plan��:�ѹ���������ѻ����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_WEEK_DATE") 'FROM_DATE
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_WEEK_DATE") 'TO_DATE
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_SUP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ���"))
   
   
   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_DESC")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӡѴ��������´(15)"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "���͡��ǧ�ѹ����Թ 7 �ѹ", , "SHOW_DATE_OVER")
   Call LoadControl("CH", cboGeneric(0).Width, True, "���͡�ҡ������ѵ�شԺ", , "SHOW_RM_ALL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "InvA", , "SHOW_INV_ACTUAL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "Diff", , "SHOW_DIFF_INV")
   Call LoadControl("CH", cboGeneric(0).Width, True, "Plan", , "SHOW_PLAN_DAILY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "Actual", , "SHOW_ACT_DAILY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "SumActual", , "SHOW_SUM_ACT_DAILY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ҳ�����ǧ˹��", , "SHOW_USE_PLAN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʹ�Ѻ��ԧ", , "SHOW_RX_ACTUAL")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport6_0_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_WEEK")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("Plan��:�ѹ���������ѻ����"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_WEEK_DATE") 'FROM_DATE
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))
   
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_WEEK_DATE") 'TO_DATE
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_SUP_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����Ѻ���"))
   
   
   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ѹ�������"))
   
      '3 =============================PART_GROUP PART_TYPE LOCATION_ID
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("������ѵ�شԺ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������ѵ�شԺ"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ʶҹ���Ѵ��"))
   
      '1 =============================FROM_PART_NO   TO_PART_NO
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "FROM_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����ѵ�شԺ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "TO_PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����ѵ�شԺ"))
   
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "LIMIT_PART_DESC")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ӡѴ��������´(15)"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "InvA", , "SHOW_INV_ACTUAL")
   Call LoadControl("CH", cboGeneric(0).Width, True, "Diff", , "SHOW_DIFF_INV")
   Call LoadControl("CH", cboGeneric(0).Width, True, "Plan", , "SHOW_PLAN_DAILY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "Actual", , "SHOW_ACT_DAILY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "SumActual", , "SHOW_SUM_ACT_DAILY")
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ҳ�����ǧ˹��", , "SHOW_USE_PLAN")
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʹ�Ѻ��ԧ", , "SHOW_RX_ACTUAL")
   
   Call ShowControl
   Call LoadComboData
   Call LoadComboData2
End Sub
Public Sub InitOrderType2(C As ComboBox, Optional Index As Integer = 0)
   C.Clear
   C.AddItem ("")
   C.ItemData(0) = 0
   C.AddItem (MapText("������ҡ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("�ҡ仹���"))
   C.ItemData(2) = 2
   
   C.ListIndex = Index
End Sub
Private Sub InitReport10_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "����ʴ���ͧ��ǹŴ", , "HIDE_DISCOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "੾�о�ѡ�ҹ���㹵��ҧ Organize", , "ONLY_HAVE_ORGANIZE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "੾�л������Թ��ҷ��Դ���", , "ONLY_TYPE_COM")
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport10_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹�ͧ��ҡ�â��"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բͧ��ҡ�â��"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "LIMIT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ����ռ�"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE", "CUSTOMER_GRADE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "INTEREST_RATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("% �͡������Ҫ�ҵ�ͻ�"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport10_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹���"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բ��"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "FREELANCE_CODE", , "FREELANCE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʿ���Ź��"))
   
      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_FREELANCE_CODE", , "FREELANCE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʿ���Ź��"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_FREELANCE_CODE", , "FREELANCE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧���ʿ���Ź��"))
      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
  Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ", , "SUMMARY")
  Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ��� ��ѡ�ҹ��� ����١���", , "ONLY_CUS")
   Call LoadControl("CH", cboGeneric(0).Width, True, "��ػ��� ��ѡ�ҹ��� ����١��� �����", , "ONLY_CUS_DOC")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������Թ���", , "SHOW_PART_NO")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "�ʴ���������١���", , "SHOW_CUSTOMER_NAME")
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "NOTE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����˵�"))
   
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReport10_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("��͹���"))
'
'   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�բ��"))
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ���"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ���"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "FREELANCE_CODE", , "FREELANCE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���ʿ���Ź��"))
   
      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_FREELANCE_CODE", , "FREELANCE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ���ʿ���Ź��"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_FREELANCE_CODE", , "FREELANCE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧���ʿ���Ź��"))
      '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
      
   Call ShowControl
   Call LoadComboData2
End Sub
Private Sub InitReportA_2_21()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 150
   
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ��� DUE"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DUE_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ��� DUE"))
   
      '1 =============================
      
      '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE", , , False)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�ѹ����͡���"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE", , , False)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ����͡���"))
      
      
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE", , , False)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�ѹ������˹��"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�����١���"))
   
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�ҡ�����١���"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_CUSTOMER_CODE", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�֧�����١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�������١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_GRADE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("�дѺ�١���"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY", "ORDER_BY_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§���"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("���§�ҡ"))
   
   Call LoadControl("CH", chkCommit(0).Width, True, "���º��º TARGET-ACTUAL", , "COMPARE_TARGET_ACTUAL_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾�з�����ʹ˹�餧�����", , "BALANCE_SUMMIT_FLAG")
   Call LoadControl("CH", chkCommit(0).Width, True, "���͡੾�з�����ʹ ACTUAL", , "BALANCE_ACTUAL_FLAG")
'   Call LoadControl("CH", chkCommit(0).Width, True, "�ʴ������ǹ��ҧ ACTUAL", , "DIFF_ACTUAL_FLAG")
   
   Call ShowControl
   Call LoadComboData2
End Sub
