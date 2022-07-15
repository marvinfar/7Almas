VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmNetwork 
   BorderStyle     =   0  'None
   Caption         =   "‰„«Ì‘ »—Ê“ «ÿ·«⁄« "
   ClientHeight    =   11520
   ClientLeft      =   990
   ClientTop       =   -390
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "B Traffic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   14340
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   14280
      TabIndex        =   1
      Top             =   0
      Width           =   14340
      Begin VB.Timer Timer1 
         Interval        =   10000
         Left            =   1560
         Top             =   120
      End
      Begin VB.TextBox TxtTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8040
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin FarDate1.FarDate FarDate1 
         Height          =   495
         Left            =   11160
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Traffic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HaftAlmas.TypeButton CmdClose 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   50
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   4
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Traffic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmNetwork.frx":0000
         PICN            =   "FrmNetwork.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin HaftAlmas.TypeButton CmdMinimize 
         Height          =   495
         Left            =   960
         TabIndex        =   8
         ToolTipText     =   "òÊçò ò—œ‰ Å‰Ã—Â"
         Top             =   50
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   4
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Traffic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmNetwork.frx":3A8A
         PICN            =   "FrmNetwork.frx":3AA6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ —Ê“"
         Height          =   405
         Left            =   13320
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "“„«‰ »Â —Ê“ —”«‰Ì "
         Height          =   405
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "œﬁÌﬁÂ"
         Height          =   405
         Left            =   7395
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   525
      End
   End
   Begin FlexCell.Grid Grid1 
      Align           =   1  'Align Top
      Height          =   10920
      Left            =   0
      TabIndex        =   0
      Top             =   615
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   19262
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
End
Attribute VB_Name = "FrmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdOK_Click()
 Timer1.Enabled = True
End Sub

Private Sub Form_Load()
 Call SetGrid
 TxtTime = 1
 FarDate1.Text = FarDate1.today
End Sub

Private Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 17
      .Rows = 1
      '
      .DefaultFont.Name = "B Traffic"
      .DefaultFont.Bold = True
      .DefaultFont.Size = 11
      
      .DefaultRowHeight = 25
      .AllowUserResizing = False
      '.AllowUserSort = True
      '
      '.BackColor1 = RGB(245, 180, 80)
      '.BackColor2 = RGB(255, 125, 9)
      '.BackColorBkg = vbBlack
      '.BackColorFixed = RGB(255, 215, 179)
      '.BackColorScrollBar = &H80FF&    'RGB(255, 125, 199)
      '
      .Column(0).Width = 20
      .Column(1).Width = 60
      .Column(2).Width = 85
      .Column(3).Width = 85
      .Column(4).Width = 60
      .Column(5).Width = 45
      .Column(6).Width = 60
      .Column(7).Width = 120
      .Column(8).Width = 90
      .Column(9).Width = 100
      .Column(10).Width = 75
      .Column(11).Width = 85
      .Column(12).Width = 45
      .Column(13).Width = 65
      .Column(14).Width = 80
      .Column(15).Width = 70
      .Column(16).Width = 40
      '
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      'Make Titr
      .Cell(0, 1).Text = "Å—Ê‰œÂ"
      .Cell(0, 2).Text = " ·›‰ —«‰‰œÂ"
      .Cell(0, 3).Text = "ò‹—«Ì‹Â"
      .Cell(0, 4).Text = "”«Ì‹“"
      .Cell(0, 5).Text = " ⁄œ«œ"
      .Cell(0, 6).Text = "Ê“‰"
      .Cell(0, 7).Text = "‘„«—Â „«‘Ì‰"
      .Cell(0, 8).Text = "ÕÊ«·Â"
      .Cell(0, 9).Text = "¬œ—”"
      .Cell(0, 10).Text = " «—ÌŒ Õ„·"
      .Cell(0, 11).Text = "‘„«—Â »«—‰«„Â"
      .Cell(0, 12).Text = "Å«— "
      .Cell(0, 13).Text = "Å—Ê«‰Â"
      .Cell(0, 14).Text = "‘„«—Â «⁄ »«—"
      .Cell(0, 15).Text = "‰«„ »«—»—Ì"
      .Cell(0, 16).Text = "—œÌ›"  ' Radife Jadval
      '
      .ReadOnly = True
      .ReadOnlyFocusRect = Solid
      '
      .SelectionMode = cellSelectionFree
      .Appearance = Flat
      
 End With

End Sub

Private Sub Timer1_Timer()
 '
 Dim rs As New Recordset
 Dim strSql As String
 Dim r As Long, i As Integer
 Dim cnsNet As String
 
 cnsNet = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.1.101\f\7Almas-Shayan\db7Almas.mdb;Persist Security Info=False"
 
 strSql = "SELECT Name,Etebar,Parvane,Part, "
 strSql = strSql & "BarName,Tarikh,Address,Havale,ShomareMashin,Vazn, "
 strSql = strSql & "Tedad,Size0,Keraye,Mobile,Parvande "
 strSql = strSql & "FROM Main7 INNER JOIN Detail7 ON Main7.Code = Detail7.Code "
 strSql = strSql & "WHERE (((Detail7.Tarikh)='" & Mid(FarDate1.Text, 3) & "')) "
 strSql = strSql & "ORDER BY Count0 "
 rs.Open strSql, cnsNet
 '
 If Not rs.EOF Then
    Grid1.Rows = 1
    Do While Not rs.EOF
       Grid1.AddItem ""
       r = Grid1.Rows - 1
       For i = 0 To 14
           Grid1.Cell(r, 16 - (i + 1)).Text = IIf(IsNull(rs(i)), "", rs(i))
       Next
       Grid1.Cell(r, 16).Text = r
       rs.MoveNext
    Loop
    rs.Close
    Grid1.Cell(1, 16).SetFocus
    Grid1.SetFocus
 Else
    rs.Close
 End If
 Set rs = Nothing
 
End Sub

Private Sub CmdMinimize_Click()
 Me.WindowState = 1
End Sub
