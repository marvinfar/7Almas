VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmGetReportFalle 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
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
   ScaleHeight     =   5910
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CombAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmGetReportFalle.frx":0000
      Left            =   5040
      List            =   "FrmGetReportFalle.frx":0002
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton CmdOkAddress 
      Caption         =   "‰‘«‰ œ«œ‰ ¬œ—”Â«"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin HaftAlmas.TypeButton TypeButton1 
      Height          =   495
      Left            =   480
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Å«ò ò—œ‰  «—ÌŒ"
      Top             =   960
      Width           =   375
      _ExtentX        =   661
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
      MICON           =   "FrmGetReportFalle.frx":0004
      PICN            =   "FrmGetReportFalle.frx":0020
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
      Height          =   2295
      Left            =   1800
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4048
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
      EnterKeyMoveTo  =   1
   End
   Begin VB.CheckBox ChkParvane 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "ê“«—‘ »— «”«” ‘„«—Â Å—Ê«‰Â"
      Height          =   315
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3150
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   5040
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
      MICON           =   "FrmGetReportFalle.frx":3A8E
      PICN            =   "FrmGetReportFalle.frx":3AAA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdOK7 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   " ‹«Ì‹Ì‹œ"
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
      MICON           =   "FrmGetReportFalle.frx":7518
      PICN            =   "FrmGetReportFalle.frx":7534
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FarDate1.FarDate FarDate1 
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   960
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
   Begin FarDate1.FarDate FarDate2 
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   960
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
   Begin HaftAlmas.TypeButton CmdOkBank 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   5040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   " ‹«Ì‹Ì‹œ"
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
      MICON           =   "FrmGetReportFalle.frx":AB08
      PICN            =   "FrmGetReportFalle.frx":AB24
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdClearAddress 
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Å«ò ò—œ‰  «—ÌŒ"
      Top             =   4200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      MICON           =   "FrmGetReportFalle.frx":E0F8
      PICN            =   "FrmGetReportFalle.frx":E114
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
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      Height          =   405
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ê“«—‘ «“  «—ÌŒ "
      Height          =   405
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1410
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â œ—Ì«›   «—ÌŒ Ê ‘„«—Â Å—Ê«‰Â »—«Ì ‰„«Ì‘ ê“«—‘"
      Height          =   405
      Left            =   525
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   7110
   End
   Begin VB.Image ImgBackground 
      Height          =   5910
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "FrmGetReportFalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkParvane_Click()
 'TxtParvane.Enabled = CBool(ChkParvane.Value)
 If ChkParvane.Value = 1 Then
    Me.Height = 5910
    ImgBackground.Height = Me.Height
    CmdOK7.Top = 5040
    CmdOkBank.Top = CmdOK7.Top
    CmdClose.Top = CmdOK7.Top
    Grid1.Visible = True
    Grid1.SetFocus
    Grid1.Cell(1, 1).SetFocus
 Else
    Me.Height = 3660
    ImgBackground.Height = Me.Height
    CmdOK7.Top = 2640
    CmdOkBank.Top = CmdOK7.Top
    CmdClose.Top = CmdOK7.Top
    Grid1.Visible = False
 End If
End Sub


Private Sub CmdClearAddress_Click()
 CombAddress.ListIndex = -1
End Sub

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdOK7_Click()
 If ChkParvane.Value = 0 And Not CheckDate Then Exit Sub
 If ChkParvane.Value = 0 And FarDate1.Text = Empty Then Exit Sub
 If ChkParvane.Value = 1 And FarDate1.Text = Empty And Grid1.Cell(1, 1).Text = Empty Then Exit Sub
 '
 Dim strSql As String
 Dim rs As New Recordset
 Dim i As Integer, kRow As Integer
 Dim sumTedad As Long, sumVazn As Long, SumVaznTarikh As Long
 Dim CurrentParvane As String
 Dim SubStr As String
 Dim MainFalleCODE As Long
 '
 strSql = "SELECT Etebar,Parvane,DetailFalle.Part,BarName, "
 strSql = strSql & "Tarikh,Address,Parvande,DriverName,ShomareMashin, "
 strSql = strSql & "Vazn ,PayeKeraye,Keraye,Mobile, Count0,MainFalle.Code,KeshtiName "
 strSql = strSql & "FROM MainFalle INNER JOIN DetailFalle ON MainFalle.Code = DetailFalle.Code "
 strSql = strSql & "WHERE "
 
 If ChkParvane.Value = 1 And Grid1.Cell(1, 1).Text <> Empty Then
    If Grid1.Cell(Grid1.Rows - 1, 1).Text = Empty Then
       Grid1.RemoveItem Grid1.Rows - 1
    End If
    
    
    SubStr = "IN("
    For i = 1 To Grid1.Rows - 1
        SubStr = SubStr & "'" & Grid1.Cell(i, 1).Text & "',"
    Next
    SubStr = Left(SubStr, Len(SubStr) - 1)
    SubStr = SubStr & ")"
    
    
    strSql = strSql & "Parvane " & SubStr & " AND"
    
    If CombAddress.ListIndex <> -1 Then ' Namayeshe ye Addresse khas
       strSql = strSql & " Address='" & CombAddress.Text & "' AND"
    End If
 End If

 If (FarDate1.Text & FarDate1.Text) <> Empty Then
    strSql = strSql & " (((DetailFalle.Tarikh) BETWEEN '" & Mid(FarDate1.Text, 3) & "' "
    strSql = strSql & "AND '" & Mid(FarDate2.Text, 3) & "')) "
 End If
 
 If Right(strSql, 3) = "AND" Then strSql = Left(strSql, Len(strSql) - 3)

 strSql = strSql & "ORDER BY MainFalle.Code,DetailFalle.Count0 "
 
 rs.Open strSql, CNS
 If Not rs.EOF Then
    ''Entekhabe name Saheb Kala
    Dim ss As String
    Dim rs1 As New Recordset
    Dim Saheb As String
    ss = "SELECT DISTINCT Moshtari.MoshtariName, MainFalle.moshtaricode "
    ss = ss & "FROM Moshtari INNER JOIN MainFalle ON "
    ss = ss & "Moshtari.MoshtariCODE = MainFalle.moshtaricode "
    ss = ss & "WHERE Parvane='" & rs("Parvane") & "'"
    rs1.Open ss, CNS
    Saheb = IIf(IsNull(rs1(0)), "", rs1(0))
    rs1.Close
    
    Set rs1 = Nothing
    '''''''''''''''''
    Dim keshtiname As String
    With FrmPreview.Grid1
        .OpenFile App.Path & "\RepFalle.cel"
        kRow = 3
        sumTedad = 0: sumVazn = 0
        Do While Not rs.EOF
           CurrentParvane = rs("Parvane")
           MainFalleCODE = rs(14)
           keshtiname = IIf(IsNull(rs("KeshtiName")), "‰œ«—œ", rs("KeshtiName"))
           
           For i = 1 To 13
               .Cell(kRow, i + 1).Text = IIf(IsNull(rs(13 - i)), "", rs(13 - i))
           Next
           .Cell(kRow, 15).Text = kRow - 2
           
           sumVazn = sumVazn + rs("Vazn")
           SumVaznTarikh = sumVazn
           kRow = kRow + 1
           .InsertRow kRow, 1
           rs.MoveNext
           If rs.EOF Then GoTo ss:
           If CurrentParvane <> rs("Parvane") Then
ss:           FrmPreview.List1.AddItem sumVazn
              FrmPreview.LstCodeParvane.AddItem MainFalleCODE
              sumTedad = 0: sumVazn = 0
           End If
        Loop
        rs.Close
        Call Molahezat
        If CombAddress.ListIndex = -1 Then ' zamani ke address bashad niazi be namayeshe bandel nist
           Call TedadBaghimande(SubStr)
        End If
        .Cell(1, 5).Text = .Cell(1, 5).Text & " " & Saheb & "- ò‘ Ì " & keshtiname
        .Cell(1, 3).Text = .Cell(1, 3).Text & Space(8) & FarDate1.today
        .AddItem ""
        .Range(.Rows - 1, 1, .Rows - 1, 9).Merge
        .Range(.Rows - 1, 10, .Rows - 1, .Cols - 1).Merge
        .Range(.Rows - 2, 1, .Rows - 2, 9).Merge
        .Cell(.Rows - 2, 1).Alignment = cellCenterCenter
        .Cell(.Rows - 1, 1).Alignment = cellCenterCenter
        .Cell(.Rows - 1, 10).Alignment = cellCenterCenter
        
        .Range(.Rows - 2, 1, .Rows - 1, 2).FontName = "B Nazanin"
        .Range(.Rows - 2, 1, .Rows - 1, 2).FontBold = True
        .Range(.Rows - 2, 1, .Rows - 1, 2).FontSize = 14
        
        .RowHeight(.Rows - 1) = 35
        .RowHeight(.Rows - 2) = 35
        
       If CombAddress.ListIndex <> -1 Then
        .Cell(.Rows - 2, 1).Text = " Ã„⁄ ò· Ê“‰ «Ì‰ ÕÊ«·Â" & SumVaznTarikh & " òÌ·Ê ê—„ „Ì »«‘œ "
       End If

        
        .Cell(.Rows - 1, 1).Text = "«—”«·Ì «“ ‘—ò  Õ„· Ê ‰ﬁ· „Â—Ê—“«‰  —«»— »‰œ— «‰“·Ì-‘«Ì«‰ „Â—         ·›‰4-01344439880 "
        .Cell(.Rows - 1, 10).Text = "Email: Mehrvarzan.tarabar@gmail.com"
        
        Call PageSetupANDFooter
        
        If Grid1.Enabled And Grid1.Visible Then Grid1.SetFocus
        .PrintPreview 110
        FrmPreview.Show 1
    End With
 Else
    rs.Close
    MsgBox "ê“«—‘Ì »—«Ì  «—ÌŒ „Ê—œ ‰Ÿ— „ÊÃÊœ ‰„Ì »«‘œ", vbInformation, ""
 End If
End Sub


Private Sub CmdOkAddress_Click()
If Grid1.Rows > 1 And Grid1.Cell(1, 1).Text <> "" Then
   Dim Hnum As Long
   Hnum = Val(Grid1.Cell(1, 1).Text)
   Dim rs As New Recordset
   Dim strSql As String

   strSql = "SELECT distinct DetailFalle.Address "
   strSql = strSql & "FROM MainFalle INNER JOIN DetailFalle ON "
   strSql = strSql & "MainFalle.Code = DetailFalle.Code "
   strSql = strSql & "WHERE (((MainFalle.Parvane)='" & Hnum & "'))"
   rs.Open strSql, CNS
   CombAddress.Clear
    Do While Not rs.EOF
    'MsgBox Trim(rs(0))
       CombAddress.AddItem Trim(rs(0))
       rs.MoveNext
    Loop
    
End If
End Sub

Private Sub FarDate1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{tab}"
    KeyCode = 0
 End If
End Sub

Private Sub FarDate1_LostFocus()
 If FarDate2.Text = Empty Then
    FarDate2.Text = FarDate1.Text
 End If

End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackFormsHazine.jpg")
 '
 FarDate1.Text = FarDate1.today
 FarDate2.Text = FarDate2.today
 '
 ChkParvane.BackColor = RGB(221, 221, 221)
 '
 'Call LoadAccountName
 '
 Call SetGrid
 '
 Me.Height = 3660
 ImgBackground.Height = Me.Height
 CmdOK7.Top = 2640
 CmdOkBank.Top = CmdOK7.Top
 CmdClose.Top = CmdOK7.Top
 Grid1.Visible = False

End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
 With Grid1
 If KeyAscii = 13 Then
    If .Cell(.Rows - 1, 1).Text <> Empty Then
       .AddItem ""
       .Cell(.Rows - 1, 1).SetFocus
    End If
 ElseIf KeyAscii = 8 Then
    If .ActiveCell.Row > 1 And .ActiveCell.Text = Empty Then
       .RemoveItem .ActiveCell.Row
    Else
       .Cell(.ActiveCell.Row, 1).Text = ""
    End If
 End If
 End With
End Sub

Private Sub LblTitle_DblClick()
  CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Function CheckDate() As Boolean
  CheckDate = True
  If (FarDate2.Text < FarDate1.Text) Then
     MsgBox " «—ÌŒ œÊ„ ‰»«Ìœ «“  «—ÌŒ «Ê· ò„ — »«‘œ", vbExclamation, ""
     FarDate2.SetFocus
     CheckDate = False
     Exit Function
  End If
End Function


Private Sub Molahezat()
 Dim r As Integer
 With FrmPreview.Grid1
      r = .Rows - 2
     .Range(r, 4, r, 14).Merge
     .Cell(r, 4).Alignment = cellCenterCenter
     .Cell(r, 4).Font.Name = "B Titr"
     .Cell(r, 4).Font.Bold = True
     .Cell(r, 4).Font.Size = 12
     .Cell(r, 4).Text = ".„·«ÕŸ« : »« ”·«„ Ê Œ” Â ‰»«‘Ìœ " & _
               "Ã„⁄ ò· ÕÊ«·Â Â«Ì »«—êÌ—Ì œ—  «—ÌŒ " & Mid(FarDate1.Text, 3) & _
               " —« »Â Õ÷Ê— «‰  ﬁœÌ„ „Ìœ«—„"
 End With
End Sub

Private Sub PageSetupANDFooter()
 Dim L As String, C As String, r As String 'left , center , right
 With FrmPreview.Grid1.PageSetup
     .PrintGridlines = True
     .BlackAndWhite = True
     .CenterHorizontally = True
     .TopMargin = 1
     .BottomMargin = 1.9
     .LeftMargin = 0.5
     .RightMargin = 0.7
     .HeaderMargin = 0.7
     .FooterMargin = 0.9
     .FooterFont.Name = "B Zar"
     .FooterFont.Size = 14
     .FooterFont.Bold = True
     .FooterAlignment = cellCenter
     ''
     r = " ‰ŸÌ„ ò‰‰œÂ:"
     C = " «∆Ìœ ò‰‰œÂ:"
     L = "’›ÕÂ :" & "&P" & " «“ " & "&N"
     .Footer = r & Space(90) & C & Space(90) & L
 End With
 
End Sub

Private Sub TedadBaghimande(SubStr As String)
 Dim rs As New Recordset
 Dim strSql As String
 Dim r As Integer
 Dim strPrompt As String
 Dim TedadRafte As Long, VaznRafte As Long
 Dim j As Integer, p As Integer
 Dim BaghiCodeList As String
 '
    BaghiCodeList = "IN("
    For r = 0 To FrmPreview.LstCodeParvane.ListCount - 1
        BaghiCodeList = BaghiCodeList & FrmPreview.LstCodeParvane.List(r) & ","
    Next

    BaghiCodeList = Left(BaghiCodeList, Len(BaghiCodeList) - 1)
    BaghiCodeList = BaghiCodeList & ")"
 
'''''''''''''''''''''''''''''''''''''''''''''
  strSql = "SELECT MAX(baghimandeFalle.Count0) AS MaxOfCount0,code "
  strSql = strSql & "FROM baghimandeFalle "
  strSql = strSql & "WHERE (baghimandeFalle.code) " & BaghiCodeList
  If FarDate1.Text & FarDate2.Text <> Empty Then
  strSql = strSql & " AND ((baghimandeFalle.tarikh)>='" & Mid(FarDate1.Text, 3)
  strSql = strSql & "' AND (baghimandeFalle.tarikh)<='" & Mid(FarDate2.Text, 3) & "') "
  End If
  strSql = strSql & " GROUP BY baghimandeFalle.code "
  strSql = strSql & " ORDER BY baghimandeFalle.code "

  rs.Open strSql, CNS
  Dim ss As String
  Dim rs1 As New Recordset
  j = -1
  Do While Not rs.EOF
     ss = "SELECT BaghiVazn FROM baghimandeFalle "
     ss = ss & "WHERE Code=" & rs("Code") & " AND Count0=" & rs("MaxOfCount0")
     rs1.Open ss, CNS
     If Not rs1.EOF Then
        With FrmPreview.Grid1
              r = FrmPreview.Grid1.Rows - 1
              .Range(r, 1, r, 13).Merge
              j = j + 1
              .RowHeight(r) = 32
              .Cell(r, 1).Alignment = cellCenterCenter
              .Cell(r, 1).Font.Name = "B Titr"
              .Cell(r, 1).Font.Bold = True
              .Cell(r, 1).Font.Size = 13
               On Error Resume Next
               
               VaznRafte = CLng(FrmPreview.List1.List(j))
               '' Get Parvane and Part From Code
               Dim rs2 As New Recordset
               rs2.Open "SELECT Parvane,Part FROM MainFalle WHERE Code=" & rs("Code"), CNS
               strPrompt = "«“ Å—Ê«‰Â " & rs2(0)
               rs2.Close: Set rs2 = Nothing
               strPrompt = strPrompt & "»Â Ê“‰ " & VaznRafte & "òÌ·Êê—„ Œ«—Ã  "
               
                  If rs1("BaghiVazn") > 0 Then
                     strPrompt = strPrompt & " Ê »Â Ê“‰ " & rs1("BaghiVazn") & "òÌ·Êê—„ »«ﬁÌ„«‰œÂ «”  "
                  ElseIf rs1("BaghiVazn") < 0 Then
                     strPrompt = strPrompt & " Ê »Â Ê“‰ " & Abs(rs1("BaghiVazn")) & "òÌ·Êê—„ «÷«›Â Ê“‰ œ«—œ "
                  ElseIf rs1("BaghiVazn") = 0 Then
                     strPrompt = strPrompt & " Ê Ê“‰ " & " Å«Ì«Å«Ì „Ì »«‘œ "
                  End If
              ' Else
              '    strPrompt = strPrompt & " »‰œ· »Â Ê“‰ " & rs1("BaghiVazn") & "òÌ·Êê—„ »«ﬁÌ„«‰œÂ «”  "
              ' End If
               .Cell(r, 1).Text = strPrompt
               .AddItem ""
               'r = r + 1
        End With
        rs1.Close
      End If
    rs.MoveNext
    rs1.MoveNext
  Loop
  rs.Close
  
  Set rs = Nothing
  Set rs1 = Nothing
End Sub

Private Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 2
      .Rows = 2
      '
      .DefaultFont.Name = "B Traffic"
      .DefaultFont.Bold = True
      .DefaultFont.Size = 12
      
      .DefaultRowHeight = 25
      .AllowUserResizing = False
      '
      .BackColorBkg = RGB(207, 219, 183)
      '
      .Column(0).Width = 20
      .Column(1).Width = 170
      '
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      'Make Titr
      .Cell(0, 1).Text = "‘„«—Â Å—Ê«‰Â"
      '
      .ReadOnlyFocusRect = Solid
      '
      .SelectionMode = cellSelectionFree
      .Appearance = Flat
      
 End With
End Sub

Private Sub TypeButton1_Click()
 FarDate1.Text = "13__/__/__"
 FarDate2.Text = "13__/__/__"
End Sub
