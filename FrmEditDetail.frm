VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmEditDetail 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   FillStyle       =   2  'Horizontal Line
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtKeshtiName 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.ComboBox CombMoshtari 
      Height          =   525
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Text            =   "CombMoshtari"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ComboBox CombMoshtariCode 
      Height          =   525
      Left            =   1680
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtWeight 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3855
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.TextBox TxtEtebar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox TxtPart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox TxtParvane 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   2880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   3375
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   8280
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmEditDetail.frx":0000
      PICN            =   "FrmEditDetail.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdEditDetail 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   7560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ÊÌ—«Ì‘ «ÿ·«⁄«  Ã«‰»Ì"
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
      MICON           =   "FrmEditDetail.frx":3A8A
      PICN            =   "FrmEditDetail.frx":3AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdFind 
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Ã” ÃÊ"
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
      MICON           =   "FrmEditDetail.frx":71EA
      PICN            =   "FrmEditDetail.frx":7206
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdEditMain 
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   3000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "«’·«Õ «ÿ·«⁄«  «’·Ì"
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
      MICON           =   "FrmEditDetail.frx":AB05
      PICN            =   "FrmEditDetail.frx":AB21
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdShowDetail 
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "‰„«Ì‘ «ÿ·«⁄«  —Ê“«‰Â"
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
      MICON           =   "FrmEditDetail.frx":E4A8
      PICN            =   "FrmEditDetail.frx":E4C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdFindDetail 
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Ã” ÃÊÌ «ÿ·«⁄«  —Ê“«‰Â œ— ÃœÊ· »—«”«” ‘„«—Â »«—‰«„Â"
      Top             =   8280
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
      MICON           =   "FrmEditDetail.frx":12086
      PICN            =   "FrmEditDetail.frx":120A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdDeleteDetail 
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Õ–› »«—‰«„Â «‰ Œ«»Ì"
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
      MICON           =   "FrmEditDetail.frx":159A1
      PICN            =   "FrmEditDetail.frx":159BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò‘ Ì"
      Height          =   405
      Left            =   3135
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’«Õ» ò«·«"
      Height          =   405
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2280
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê“‰ ‰«Œ«·’"
      Height          =   405
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label LblAlarm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmEditDetail.frx":195BD
      ForeColor       =   &H00FF0000&
      Height          =   1125
      Left            =   3480
      MouseIcon       =   "FrmEditDetail.frx":19657
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7440
      Width           =   7170
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â «⁄ »«—"
      Height          =   405
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‹‹«— "
      Height          =   405
      Left            =   9900
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   10560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»—«Ì Ã” ÃÊ ‘„«—Â Å—Ê«‰Â —« Ê«—œ ò‰Ìœ"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   6315
      MouseIcon       =   "FrmEditDetail.frx":19961
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   3360
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã‹—Â Ã” Ã‹Ê Ê ÊÌ‹—«Ì‘ «ÿ·«⁄«  «’·Ì Ê —Ê“«‰‹Â "
      Height          =   405
      Left            =   705
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   120
      Width           =   8880
   End
   Begin VB.Image ImgBackground 
      Height          =   9060
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "FrmEditDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodePaEt As Long
Dim EditMode As Boolean
Dim CellTedadChange As Boolean
Dim CellVaznChange As Boolean
Dim mbIgnoreListClick   As Boolean

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()

End Sub

Private Sub CmdDeleteDetail_Click()
 Dim r As Long
 Dim Count0 As Integer
 Dim Tarikh As String
 Dim msg As Integer, i As Integer
 
 r = Grid1.ActiveCell.Row
 If r > 0 Then
    Count0 = Val(Grid1.Cell(r, 14).Text)
    Tarikh = Grid1.Cell(r, 1).Text
    '
    Grid1.Range(r, 1, r, 15).BackColor = vbRed
    msg = MsgBox("¬Ì« „ÿ„∆‰ »Â Õ–› Â” Ìœø", vbQuestion + vbYesNo, "")
    Grid1.Range(r, 1, r, 15).BackColor = vbWhite
    If msg = vbYes Then
       Dim rs As New Recordset
       Dim strSql As String
       
       strSql = "DELETE FROM Detail7 "
       strSql = strSql & "WHERE Code=" & CodePaEt
       strSql = strSql & " AND Count0=" & Count0
       rs.Open strSql, CNS
       ''
       Call UpdateBaghimandeTable(r)
       ''Pak Kardane Satre Jadval
       Grid1.RemoveItem r
       ''Radif Kardane Count0 dar bank va Jadval
       strSql = "SELECT Count0 FROM Detail7 "
       strSql = strSql & "WHERE Code=" & CodePaEt
       rs.Open strSql, CNS, adOpenStatic, adLockOptimistic
       
       For i = 1 To Grid1.Rows - 1
           rs(0) = i
           Grid1.Cell(i, 14).Text = i
           Grid1.Cell(i, 15).Text = i
           rs.Update
           rs.MoveNext
           If rs.EOF Then Exit For
       Next
       rs.Close
       MsgBox "«ÿ·«⁄«  »«—‰«„Â Õ–› ‘œ", vbInformation, ""
       Set rs = Nothing
    End If
 Else
    MsgBox "·ÿ›« »—«Ì Õ–› »—‰«„Â Ìò ”ÿ— —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
 End If
End Sub

Private Sub CmdEditDetail_Click()
 Static r As Long
 Dim i As Integer
 
 If CmdEditDetail.Caption = "ÊÌ—«Ì‘ «ÿ·«⁄«  Ã«‰»Ì" Then
    r = Grid1.ActiveCell.Row
    If r > 0 Then
       Grid1.ReadOnly = False
       For i = 1 To Grid1.Rows - 1
          If i <> r Then
             Grid1.Range(i, 1, i, 15).Locked = True
          Else
             Grid1.Range(i, 1, i, 15).Locked = False
          End If
       Next
       Grid1.Range(r, 1, r, 13).BackColor = QBColor(11)
       Grid1.Cell(r, 15).Locked = True
       Grid1.Cell(r, Grid1.ActiveCell.Col).SetFocus
       CmdEditDetail.Caption = "À»   €ÌÌ—« "
       EditMode = True
       '
       CmdShowDetail.Enabled = False
       CmdEditMain.Enabled = False
       CmdFind.Enabled = False
       CmdDeleteDetail.Enabled = False
    Else
       MsgBox "·ÿ›« ”ÿ—Ì —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
       Exit Sub
    End If
 ElseIf CmdEditDetail.Caption = "À»   €ÌÌ—« " Then
    Dim rs As New Recordset
    Dim strSql As String
    strSql = SubMakeQuery(r)
    rs.Open strSql, CNS
    If CellTedadChange Or CellVaznChange Then UpdateBaghimandeTable (r)
    
    Grid1.Range(r, 1, r, 13).BackColor = vbWhite
    For i = 1 To Grid1.Rows - 1
        Grid1.Range(i, 1, i, 15).Locked = True
    Next
    Grid1.ReadOnly = True
    MsgBox "«ÿ·«⁄«   €ÌÌ— œ«œÂ ‘œÂ «’·«Õ ‘œ", vbInformation, ""
    CmdEditDetail.Caption = "ÊÌ—«Ì‘ «ÿ·«⁄«  Ã«‰»Ì"
    
    CmdShowDetail.Enabled = True
    CmdEditMain.Enabled = True
    CmdFind.Enabled = True
    CmdDeleteDetail.Enabled = True
    
    EditMode = False
    CellTedadChange = False
    CellVaznChange = False
    Set rs = Nothing
 End If
 
End Sub

Private Sub CmdFindDetail_Click()
 If Grid1.Rows > 1 Then
    Dim inp As Long
    Dim i As Integer
    Dim b As Boolean
    
    inp = CLng(Val(InputBox("·ÿ›« ‘„«—Â »«—‰«„Â —« Ê«—œ ‰„«ÌÌœ", "Ã” ÃÊ")))
    If inp = 0 Then Exit Sub
    b = False
    For i = 1 To Grid1.Rows - 1
        If CLng(Grid1.Cell(i, 12).Text) = inp Then
           b = True
           Exit For
        End If
    Next
    '
    If b Then
       Grid1.Cell(i, 12).SetFocus
       Grid1.SetFocus
    Else
       MsgBox "‘„«—Â »«—‰«„Â „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbInformation, ""
    End If
       
End If
    
End Sub

Private Sub CmdShowDetail_Click()
 Dim rs As New Recordset
 Dim strSql As String
 Dim i As Integer, r As Integer
 '
 strSql = "SELECT * FROM Detail7 "
 strSql = strSql & "WHERE Code=" & CodePaEt
 rs.Open strSql, CNS
 If rs.EOF Then
    MsgBox "«ÿ·«⁄«  —Ê“«‰Â ÊÃÊœ ‰œ«—œ", vbExclamation, ""
    TxtParvane.SetFocus
    SendKeys "{home}+{end}"
    rs.Close
 Else
    Grid1.Rows = 1
    Do While Not rs.EOF
       Grid1.AddItem ""
       r = Grid1.Rows - 1
       With Grid1
            .Cell(r, 1).Text = IIf(IsNull(rs("Mobile")), "", rs("Mobile"))
            .Cell(r, 2).Text = IIf(IsNull(rs("Keraye")), "", rs("Keraye"))
            .Cell(r, 3).Text = IIf(IsNull(rs("PayeKeraye")), "", rs("PayeKeraye"))
            .Cell(r, 4).Text = IIf(IsNull(rs("Size0")), "", rs("Size0"))
            .Cell(r, 5).Text = IIf(IsNull(rs("Tedad")), "", rs("Tedad"))
            .Cell(r, 6).Text = IIf(IsNull(rs("Vazn")), "", rs("Vazn"))
            .Cell(r, 7).Text = IIf(IsNull(rs("ShomareMashin")), "", rs("ShomareMashin"))
            '.Cell(r, 8).Text = IIf(IsNull(rs("Momayez2")), "", rs("Momayez2"))
            .Cell(r, 8).Text = IIf(IsNull(rs("DriverName")), "", rs("DriverName"))
            .Cell(r, 9).Text = IIf(IsNull(rs("Parvande")), "", rs("Parvande"))
            '.Cell(r, 9).Text = IIf(IsNull(rs("Havale")), "", rs("Havale"))Momayez
            .Cell(r, 10).Text = IIf(IsNull(rs("Address")), "", rs("Address"))
            .Cell(r, 11).Text = IIf(IsNull(rs("Tarikh")), "", rs("Tarikh"))
            .Cell(r, 12).Text = IIf(IsNull(rs("BarName")), "", rs("BarName"))
            .Cell(r, 13).Text = rs("Name")
            .Cell(r, 14).Text = rs("Count0")
            .Cell(r, 15).Text = r
            
       End With
       rs.MoveNext
    Loop
    rs.Close
    Grid1.Cell(1, 15).SetFocus
    Grid1.SetFocus
    '
    LblAlarm.Visible = True
 End If
 
 Set rs = Nothing
End Sub

Private Sub CmdEditMain_Click()
 Dim rs As New Recordset
 Dim strSql As String
 
 If CmdEditMain.Caption = "«’·«Õ «ÿ·«⁄«  «’·Ì" Then
    TxtParvane.Enabled = False
    CmdFind.Enabled = False
    CmdShowDetail.Enabled = False
    TxtEtebar.SetFocus
    CmdEditMain.Caption = "À»   €ÌÌ—« "
    TxtEtebar.Locked = False
    TxtPart.Locked = False
    TxtWeight.Locked = False
    CombMoshtari.Locked = False
 ElseIf CmdEditMain.Caption = "À»   €ÌÌ—« " Then
    TxtParvane.Enabled = True
    CmdFind.Enabled = True
    CmdShowDetail.Enabled = True
    
    strSql = "UPDATE Main7 SET "
    strSql = strSql & "Etebar='" & TxtEtebar & "',Part=" & Val(TxtPart)
    strSql = strSql & ",Weight=" & Val(TxtWeight)
    strSql = strSql & ",MoshtariCode=" & Val(CombMoshtariCode)
    strSql = strSql & ",KeshtiName='" & Trim(TxtKeshtiName) & "'"
    strSql = strSql & " WHERE Code=" & CodePaEt
    rs.Open strSql, CNS
    
    TxtEtebar.Locked = True
    TxtPart.Locked = True
    TxtWeight.Locked = True
    CombMoshtari.Locked = True
    CmdEditMain.Caption = "«’·«Õ «ÿ·«⁄«  «’·Ì"
    'ClearText Me
    MsgBox "«ÿ·«⁄«  «’·«Õ ‘œ", vbInformation, ""
 End If
 
End Sub

Private Sub CmdFind_Click()
 If TxtParvane = Empty Then
    MsgBox "‘„«—Â Å—Ê«‰Â —« Ê«—œ ò‰Ìœ", vbExclamation, ""
    TxtParvane.SetFocus
    Exit Sub
 End If
 
 Dim strSql As String
 Dim rs As New Recordset
 
 strSql = "SELECT * FROM Main7 "
 strSql = strSql & "WHERE Parvane='" & TxtParvane & "'"
 rs.Open strSql, CNS
 '
 If rs.EOF Then
    MsgBox "Å‹«—  »« «Ì‰ ‘„«—Â Å—Ê«‰Â ÊÃÊœ ‰œ«—œ", vbInformation, ""
    rs.Close
 Else
    CodePaEt = rs("Code")
    TxtEtebar = rs("Etebar")
    TxtPart = rs("Part")
    TxtKeshtiName = rs("KeshtiName")
    
    If rs("MoshtariCode") <> 0 Then
       Dim i As Integer
       For i = 0 To CombMoshtariCode.ListCount - 1
           If Val(CombMoshtariCode.List(i)) = rs("MoshtariCode") Then
              CombMoshtariCode.ListIndex = i
              CombMoshtari.ListIndex = i
              Exit For
           End If
       Next
    End If
    
    TxtWeight = IIf(IsNull(rs("Weight")), 0, rs("weight"))
    rs.Close
    CmdEditMain.Enabled = True
    CmdShowDetail.Enabled = True
 End If
 
 
End Sub

Private Sub CombMoshtari_Change()
 If CombMoshtari.Text = Empty Then CombMoshtariCode.ListIndex = -1
End Sub

Private Sub CombMoshtari_Click()
 CombMoshtariCode.ListIndex = CombMoshtari.ListIndex
End Sub

Private Sub CombMoshtari_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
End Sub

Private Sub CombMoshtari_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    SendKeys "{tab}"
    KeyCode = 0
 End If
End Sub

Private Sub CombMoshtari_KeyPress(KeyAscii As Integer)
  Dim sSearchText As String
  Dim lReturn As Long
  
  If KeyAscii = 13 Then
      CombMoshtari_Click
      KeyAscii = 0
  Else
      sSearchText = Left$(CombMoshtari.Text, CombMoshtari.SelStart) & Chr$(KeyAscii)
      lReturn = SendMessage(CombMoshtari.hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
      If lReturn <> CB_ERR Then
          mbIgnoreListClick = True
          CombMoshtari.ListIndex = lReturn
          mbIgnoreListClick = False
          CombMoshtari.Text = CombMoshtari.List(lReturn)
          CombMoshtari.SelStart = Len(sSearchText)
          CombMoshtari.SelLength = Len(CombMoshtari.Text)
          KeyAscii = 0
      End If
  End If
  '''

End Sub


Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackForms7.jpg")
 '
 TxtEtebar.Locked = True
 TxtPart.Locked = True
 TxtWeight.Locked = True
 CombMoshtari.Locked = True
 LblAlarm.Visible = False
 '
 CmdEditMain.Enabled = False
 CmdShowDetail.Enabled = False
 
 Call SetGrid
 '
 Dim rs As New Recordset
 rs.Open "SELECT * FROM Moshtari ", CNS
 Do While Not rs.EOF
    CombMoshtari.AddItem rs("MoshtariName")
    CombMoshtariCode.AddItem rs("MoshtariCODE")
    rs.MoveNext
 Loop
 rs.Close
 Set rs = Nothing
End Sub

Private Sub Grid1_CellChange(ByVal Row As Long, ByVal Col As Long)
 If EditMode Then
    If Col = 5 Then CellTedadChange = True
    If Col = 6 Then CellVaznChange = True
 End If
End Sub

Private Sub Grid1_GotFocus()
 CmdFind.Default = False
End Sub

Private Sub Grid1_LostFocus()
CmdFind.Default = True
End Sub

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub TxtEtebar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789/" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtPart_KeyPress(KeyAscii As Integer)
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtParvane_KeyPress(KeyAscii As Integer)
' Dim strValid As String
'   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
'   If InStr(strValid, Chr(KeyAscii)) = 0 Then
'      KeyAscii = 0
'   End If
End Sub

Private Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 16
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
      .Column(1).Width = 80
      .Column(2).Width = 85
      .Column(3).Width = 85
      .Column(4).Width = 60
      .Column(5).Width = 45
      .Column(6).Width = 60
      .Column(7).Width = 120
      .Column(8).Width = 90
      .Column(9).Width = 90
      .Column(10).Width = 100
      .Column(11).Width = 75
      .Column(12).Width = 85
      .Column(13).Width = 70
      .Column(14).Width = 0 ' Hidden For Count0
      .Column(15).Width = 40
      '
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      'Make Titr
      .Cell(0, 1).Text = " ·›‰ —«‰‰œÂ"
      .Cell(0, 2).Text = "ò‹—«Ì‹Â"
      .Cell(0, 3).Text = "Å«ÌÂ ò—«ÌÂ"
      .Cell(0, 4).Text = "”«Ì‹“"
      .Cell(0, 5).Text = " ⁄œ«œ"
      .Cell(0, 6).Text = "Ê“‰"
      .Cell(0, 7).Text = "‘„«—Â „«‘Ì‰"
      .Cell(0, 8).Text = "—«‰‰œÂ"
      .Cell(0, 9).Text = "ÕÊ«·Â"
      .Cell(0, 10).Text = "¬œ—”"
      .Cell(0, 11).Text = " «—ÌŒ Õ„·"
      .Cell(0, 12).Text = "‘„«—Â »«—‰«„Â"
      .Cell(0, 13).Text = "‰«„ »«—»—Ì"
      .Cell(0, 14).Text = "—œÌ›0" ' Count0 ast Code nemikham hamoon CODEPaEt
      .Cell(0, 15).Text = "—œÌ›"  ' Radife Jadval
      '
      .ReadOnly = True
      .ReadOnlyFocusRect = Solid
      '
      .SelectionMode = cellSelectionFree
      .Appearance = Flat
      
 End With

End Sub

Private Function SubMakeQuery(r As Long) As String
 Dim strSql As String
 
 With Grid1
   strSql = "UPDATE Detail7 SET "
   strSql = strSql & "Name='" & .Cell(r, 13).Text & "'"
   strSql = strSql & ",BarNamE=" & Val(.Cell(r, 12).Text)
   strSql = strSql & ",Tarikh='" & .Cell(r, 11).Text & "'"
   strSql = strSql & ",Address='" & .Cell(r, 10).Text & "'"
   strSql = strSql & ",Parvande='" & .Cell(r, 9).Text & "'"
   strSql = strSql & ",DriverName='" & .Cell(r, 8).Text & "'"
   strSql = strSql & ",ShomareMashin='" & .Cell(r, 7).Text & "'"
   strSql = strSql & ",Vazn=" & Val(.Cell(r, 6).Text)
   strSql = strSql & ",Tedad=" & Val(.Cell(r, 5).Text)
   strSql = strSql & ",Size0='" & .Cell(r, 4).Text & "'"
   strSql = strSql & ",PayeKeraye=" & Val(.Cell(r, 3).Text)
   strSql = strSql & ",Keraye=" & Val(.Cell(r, 2).Text)
   strSql = strSql & ",Mobile='" & .Cell(r, 1).Text & "' "
   'strSql = strSql & ",Parvande='" & .Cell(r, 1).Text & "' "
   strSql = strSql & "WHERE Code=" & CodePaEt
   strSql = strSql & " AND Count0=" & Val(.Cell(r, 14).Text)
   
   SubMakeQuery = strSql
 End With
End Function

Private Sub UpdateBaghimandeTable(r As Long)
 ' vaghti khaneye tedad taghir konad bayad dar Jadvale
  '' Baghimande ham taghirat asar begzarad
 Dim strSql As String
 Dim sumTedad As Long, sumVazn As Long
 Dim rs As New Recordset
 Dim TempRS As New Recordset
 Dim PrevCount0 As Long
 Dim Baghi As Integer
 Dim BaghiVazn As Long
 
 MsgBox "«ÿ·«⁄«   ⁄œ«œ Ê Ê“‰ »Â —Ê“ „Ì ‘Êœ", vbInformation, ""
 With Grid1
   strSql = "SELECT SUM(Tedad),SUM(Vazn) FROM Detail7 "
   strSql = strSql & "WHERE Code=" & CodePaEt
   strSql = strSql & " AND Tarikh='" & .Cell(r, 11).Text & "'"
   rs.Open strSql, CNS
   sumTedad = IIf(IsNull(rs(0)), 0, rs(0))
   sumVazn = IIf(IsNull(rs(1)), 0, rs(1))
   rs.Close
   ''
   strSql = "SELECT Count0 FROM Baghimande7 "
   strSql = strSql & "WHERE Code=" & CodePaEt
   strSql = strSql & " AND Tarikh='" & .Cell(r, 11).Text & "'"
   rs.Open strSql, CNS
   PrevCount0 = rs(0) - 1
   rs.Close
   ''
   If PrevCount0 = 0 Then ' yani taghir bar roye avalin satr anjam mishavad
      strSql = "UPDATE Baghimande7 SET "
      strSql = strSql & "Baghimande=" & Val(TxtPart) - sumTedad
      strSql = strSql & ",BaghiVazn=" & Val(TxtWeight) - sumVazn
      strSql = strSql & " WHERE Code=" & CodePaEt
      strSql = strSql & " AND Count0=1"
      rs.Open strSql, CNS
      '
      strSql = "SELECT * FROM Baghimande7 "
      strSql = strSql & "WHERE Code=" & CodePaEt
      rs.Open strSql, CNS, adOpenStatic, adLockOptimistic
      
      Do While Not rs.EOF
         Baghi = rs("Baghimande")
         BaghiVazn = rs("BaghiVazn")
         rs.MoveNext
         If rs.EOF Then Exit Do
         strSql = "SELECT SUM(Tedad),SUM(Vazn) FROM Detail7 "
         strSql = strSql & "WHERE Code=" & CodePaEt
         strSql = strSql & " AND Tarikh='" & rs("Tarikh") & "'"
         TempRS.Open strSql, CNS
         rs("Baghimande") = Baghi - TempRS(0)
         rs("BaghiVazn") = BaghiVazn - TempRS(1)
         rs.Update
         TempRS.Close
      Loop
      rs.Close
   Else
      ''
      strSql = "SELECT Baghimande,BaghiVazn FROM Baghimande7 "
      strSql = strSql & "WHERE Code=" & CodePaEt
      strSql = strSql & " AND Count0=" & PrevCount0
      rs.Open strSql, CNS
      Baghi = rs(0)
      BaghiVazn = rs(1)
      rs.Close
      '''===

      strSql = "UPDATE Baghimande7 SET "
      strSql = strSql & "Baghimande=" & Baghi - sumTedad
      strSql = strSql & ",BaghiVazn=" & BaghiVazn - sumVazn
      strSql = strSql & " WHERE Code=" & CodePaEt
      strSql = strSql & " AND Count0=" & PrevCount0 + 1
      rs.Open strSql, CNS
   End If
   Set rs = Nothing
   Set TempRS = Nothing
 End With
 
End Sub

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub
