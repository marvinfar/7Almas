VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmDeleteBank 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
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
   ScaleHeight     =   5310
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CombCodeBank 
      Height          =   525
      Left            =   2520
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.ComboBox CombAccountName 
      Height          =   525
      Left            =   3120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "»” ‰ Å‰Ã—Â"
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
      BCOL            =   16761087
      BCOLO           =   12583104
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDeleteBank.frx":0000
      PICN            =   "FrmDeleteBank.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdDeleteTransAction 
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Õ–› ”ÿ— ⁄„·Ì« "
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
      BCOL            =   16761087
      BCOLO           =   12583104
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDeleteBank.frx":3A8A
      PICN            =   "FrmDeleteBank.frx":3AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin HaftAlmas.TypeButton CmdDeleteBank 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Õ–› ò· Õ”«»"
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
      BCOL            =   16761087
      BCOLO           =   12583104
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDeleteBank.frx":76A6
      PICN            =   "FrmDeleteBank.frx":76C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Õ”«» Ì« »«‰ò"
      Height          =   405
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â Õ–› Õ”«» Ì« ⁄„·Ì«  Õ”«» »«‰òÌ"
      Height          =   405
      Left            =   300
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   7335
   End
   Begin VB.Image ImgBackground 
      Height          =   5340
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "FrmDeleteBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdDeleteBank_Click()
 Dim msg As Integer
 Dim L As Integer
 
 L = CombAccountName.ListIndex
 If L <> -1 Then
    msg = MsgBox(" Õ–› Õ”«» »Â Â„—«Â ò· ⁄„·Ì«  «‰Ã«„ ŒÊ«Âœ ‘œ! „ÿ„∆‰ Â” Ìœø ", vbCritical + vbYesNo + vbDefaultButton2, "")
    If msg = vbYes Then
       Dim strSql As String
       Dim rs As New Recordset
       'delete detail
       strSql = "DELETE FROM TransactionBank "
       strSql = strSql & "WHERE CodeBank=" & Val(CombCodeBank)
       rs.Open strSql, CNS
       'delete main
       strSql = "DELETE FROM DefBank "
       strSql = strSql & "WHERE CodeBank=" & Val(CombCodeBank)
       rs.Open strSql, CNS
       '
       Grid1.Rows = 1
       CombAccountName.RemoveItem L
       CombCodeBank.RemoveItem L
    End If
 End If
 Set rs = Nothing
End Sub

Private Sub CmdDeleteTransAction_Click()
 Dim L As Integer
 
 L = Grid1.ActiveCell.Row
 If L > 0 Then
    Dim strSql As String
    Dim rs As New Recordset
    Dim i As Integer
    '
    i = MsgBox("»—«Ì Õ–› ⁄„·Ì«  „ÿ„∆‰ „Ì »«‘Ìœø", vbQuestion + vbYesNo, "")
    If i = vbYes Then
       strSql = strSql & "AND Count0=" & Val(Grid1.Cell(L, 5).Text)
       rs.Open strSql, CNS
       Grid1.RemoveItem L
       'Make Count0 AutoNumber
       strSql = "SELECT Count0 FROM TransactionBank "
       strSql = strSql & "WHERE CodeBank=" & Val(CombCodeBank)
       rs.Open strSql, CNS, adOpenStatic, adLockOptimistic
       i = 1
       Do While Not rs.EOF
          rs(0) = i
          Grid1.Cell(i, 5).Text = i
          rs.Update
          rs.MoveNext
          i = i + 1
       Loop
       rs.Close
    End If
 Else
   MsgBox "»—«Ì Õ–› »«Ìœ Ìò ”ÿ— —« «‰ Œ«» ò‰Ìœ", vbExclamation, ""
 End If
 Set rs = Nothing
End Sub

Private Sub CombAccountName_Click()
 CombCodeBank.ListIndex = CombAccountName.ListIndex
 If CombAccountName.ListIndex = -1 Then
    Grid1.Rows = 1
 Else
    Dim strSql As String
    Dim rs As New Recordset
    Dim i As Integer
    '
    Grid1.Rows = 1
    strSql = "SELECT * FROM TransactionBank "
    strSql = strSql & "WHERE CodeBank=" & Val(CombCodeBank)
    rs.Open strSql, CNS
    Do While Not rs.EOF
       With Grid1
         .AddItem ""
         For i = 1 To 5
             .Cell(.Rows - 1, i).Text = rs(6 - i)
         Next
       End With
       rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
 End If
End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 Me.Top = Me.Top - 1000
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackFormsBanki.jpg")
 '
 Call LoadAccountName
 Call SetGrid
End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 If Grid1.MouseCol = 3 Then
    Grid1.ToolTipText = Grid1.Cell(Grid1.MouseRow, Grid1.MouseCol).Text
 End If
End Sub

Private Sub LblTitle_DblClick()
  CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub LoadAccountName()
 Dim strSql As String
 Dim rs As New Recordset
 '
 strSql = "SELECT CodeBank,AccountName FROM DefBank "
 rs.Open strSql, CNS
 
 CombAccountName.Clear
 CombCodeBank.Clear
 Do While Not rs.EOF
    CombCodeBank.AddItem rs(0)
    CombAccountName.AddItem rs(1)
    rs.MoveNext
 Loop
 rs.Close
 Set rs = Nothing
End Sub

Private Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 6
      .Rows = 1
      '
      .DefaultFont.Name = "B Traffic"
      .DefaultFont.Bold = True
      .DefaultFont.Size = 11
      
      .DefaultRowHeight = 25
      .AllowUserResizing = False
      '.AllowUserSort = True
      '
      .BackColor1 = RGB(111, 200, 81)
      .BackColor2 = RGB(108, 181, 100)
      '.BackColorBkg = vbBlack
      '.BackColorFixed = RGB(255, 215, 179)
      '.BackColorScrollBar = &H80FF&    'RGB(255, 125, 199)
      '
      .Column(0).Width = 15
      .Column(1).Width = 100 ' bedehkar
      .Column(2).Width = 100 ' bestankar
      .Column(3).Width = 150 ' Sharh
      .Column(4).Width = 80 ' Tarikh
      .Column(5).Width = 50  ' Radif
      '
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      'Make Titr
      .Cell(0, 1).Text = "»œÂò«—"
      .Cell(0, 2).Text = "»” «‰ò«—"
      .Cell(0, 3).Text = "‘—Õ ⁄„·Ì‹‹« "
      .Cell(0, 4).Text = " «—ÌŒ"
      .Cell(0, 5).Text = "—œÌ›"
      '
      .ReadOnly = True
      .ReadOnlyFocusRect = Solid
      '
      .SelectionMode = cellSelectionByRow
      .Appearance = Flat
      
 End With
End Sub

Private Sub TypeButton1_Click()
 
End Sub
