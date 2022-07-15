VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmMoshtari 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
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
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox TxtPhone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   6255
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2655
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4683
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5160
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
      MICON           =   "FrmMoshtari.frx":0000
      PICN            =   "FrmMoshtari.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdSave 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "À»  «ÿ·«⁄« "
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
      MICON           =   "FrmMoshtari.frx":3A8A
      PICN            =   "FrmMoshtari.frx":3AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdDelete 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Õ–› «ÿ·«⁄« "
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
      MICON           =   "FrmMoshtari.frx":707A
      PICN            =   "FrmMoshtari.frx":7096
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdEdit 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ÊÌ—«Ì‘"
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
      MICON           =   "FrmMoshtari.frx":AC96
      PICN            =   "FrmMoshtari.frx":ACB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdReport 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   5160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ê“«—‘ »«—‰«„Â »—«”«” ’«Õ» ò«·«"
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
      MICON           =   "FrmMoshtari.frx":E3F6
      PICN            =   "FrmMoshtari.frx":E412
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Ã” ÃÊÌ ÕÊ«·Â"
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Ã” ÃÊ ‰«„"
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
      MICON           =   "FrmMoshtari.frx":11D07
      PICN            =   "FrmMoshtari.frx":11D23
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
      Caption         =   "‘„«—Â  ·›‰"
      Height          =   405
      Left            =   6900
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ’«Õ» ò«·«"
      Height          =   405
      Left            =   6660
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â „⁄—›Ì ’«Õ»«‰ ò«·« ÿ—› ﬁ—«—œ«œ »« œ› —                                                                              "
      Height          =   405
      Left            =   600
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   7425
   End
   Begin VB.Image ImgBackground 
      Height          =   5895
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "FrmMoshtari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdDelete_Click()
 Dim msg As Integer
 Dim L As Long
 
 L = Grid1.ActiveCell.Row
 If L > 0 Then
    msg = MsgBox("¬Ì« „ÿ„∆‰ »Â Õ–› Â” Ìœø", vbQuestion + vbYesNo, "")
    If msg = vbYes Then
       Dim strSql As String
       Dim rs As New Recordset
       Dim MoshtariCode As Byte
       MoshtariCode = Val(Grid1.Cell(L, 4).Text)
       ' Check no related record
       strSql = "SELECT MoshtariCode FROM Detail7 "
       strSql = strSql & "WHERE MoshtariCode=" & MoshtariCode
       rs.Open strSql, CNS
       If Not rs.EOF Then
          MsgBox "«Ì‰ ’«Õ» ò«·« »Â œ·Ì· œ«‘ ‰ «ÿ·«⁄«  œ— ÃœÊ· »«—‰«„Â Â« ﬁ«»· Õ–› ‰Ì” ", vbExclamation, ""
          rs.Close
          Exit Sub
       End If
       rs.Close
       '''' IF no problem
       strSql = "DELETE FROM Moshtari "
       strSql = strSql & "WHERE MoshtariCODE=" & MoshtariCode
       rs.Open strSql, CNS
       Grid1.RemoveItem L
       For L = 1 To Grid1.Rows - 1
           Grid1.Cell(L, 3).Text = L
       Next
       Set rs = Nothing
    End If
 Else
    MsgBox "»—«Ì Õ–› »«Ìœ Ìò ”ÿ— —« «‰ Œ«» ò‰Ìœ", vbExclamation, ""
 End If
End Sub

Private Sub CmdEdit_Click()
 Dim L As Long
 
 L = Grid1.ActiveCell.Row
 If CmdEdit.Caption = "ÊÌ—«Ì‘" Then
    If L > 0 Then
       TxtName = Grid1.Cell(L, 2).Text
       TxtPhone = Grid1.Cell(L, 1).Text
       TxtName.SetFocus
       SendKeys "{home}+{end}"
       '
       CmdSave.Enabled = False
       CmdDelete.Enabled = False
       CmdFind.Enabled = False
       CmdReport.Enabled = False
       Grid1.Enabled = False
       '
       CmdEdit.Caption = "À»   €ÌÌ—« "
    Else
       MsgBox "»—«Ì ÊÌ—«Ì‘ ”ÿ— „Ê—œ ‰Ÿ— —« «‰ Œ«» ò‰Ìœ", vbExclamation, ""
    End If
 ElseIf CmdEdit.Caption = "À»   €ÌÌ—« " Then
    Dim rs As New Recordset
    Dim strSql As String
    
    strSql = "UPDATE Moshtari SET "
    strSql = strSql & "MoshtariName='" & Trim(TxtName) & "',"
    strSql = strSql & "Phone='" & Trim(TxtPhone) & "' "
    strSql = strSql & "WHERE MoshtariCODE=" & Val(Grid1.Cell(L, 4).Text)
    rs.Open strSql, CNS
    
    Grid1.Cell(L, 2).Text = TxtName
    Grid1.Cell(L, 1).Text = TxtPhone
    '
    Grid1.Enabled = True
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CmdFind.Enabled = True
    CmdReport.Enabled = True

    '
    TxtName = ""
    TxtPhone = ""
    '
    CmdEdit.Caption = "ÊÌ—«Ì‘"
    '
    Set rs = Nothing
 End If
 
End Sub

Private Sub CmdFind_Click()
 If Grid1.Rows > 1 Then
    Dim inp As String
    Dim i As Integer
    Dim b As Boolean
    
    inp = InputBox("·ÿ›« ‰«„ ’«Õ» ò«·« —« Ê«—œ ‰„«ÌÌœ", "Ã” ÃÊ")
    If inp = "" Then Exit Sub
    b = False
    For i = 1 To Grid1.Rows - 1
        If Grid1.Cell(i, 2).Text = inp Then
           b = True
           Exit For
        End If
    Next
    '
    If b Then
       Grid1.Cell(i, 2).SetFocus
       Grid1.SetFocus
    Else
       MsgBox "’«Õ» ò«·« „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbInformation, ""
    End If
       
End If

End Sub

Private Sub CmdReport_Click()
 Dim L As Long
 
 L = Grid1.ActiveCell.Row
 If L > 0 Then
    Dim inpDate As String
    Dim strPrompt As String
    strPrompt = "·ÿ›«  «—ÌŒ —« Ê«—œ ‰„«ÌÌœ"
    strPrompt = strPrompt & vbNewLine & " ›—„ ’ÕÌÕ  «—ÌŒ »Â ’Ê—  “Ì— »«‘œ"
    strPrompt = strPrompt & vbNewLine & "yy/mm/dd <--->89/01/01"
    inpDate = InputBox(strPrompt, "", Mid(FrmGetReportINFO7.FarDate1.today, 3))
    If inpDate <> Empty Then
       If Left(inpDate, 2) = "13" Then inpDate = Mid(inpDate, 3)
       Call MakeReport(inpDate, CByte(Grid1.Cell(L, 4).Text), Grid1.Cell(L, 2).Text)
    End If
 Else
    MsgBox "»—«Ì ê“«—‘ »«Ìœ Ìò ”ÿ— «‰ Œ«» ‘Êœ", vbExclamation, ""
 End If

End Sub

Private Sub CmdSave_Click()
 Dim CodeMoshtari As Byte
 If TxtName <> Empty Then
    Dim rs As New Recordset
    Dim strSql As String
    '
    strSql = "SELECT MoshtariName FROM Moshtari "
    strSql = strSql & "WHERE MoshtariName='" & Trim(TxtName) & "'"
    rs.Open strSql, CNS
    If Not rs.EOF Then
       MsgBox "‰«„ ’«Õ» ò«·«  ò—«—Ì «” ", vbExclamation, ""
       TxtName.SetFocus
       SendKeys "{home}+{end}"
       rs.Close
       Exit Sub
    End If
    rs.Close
    '
    CodeMoshtari = CByte(MakeAutoNumber("Moshtari", "MoshtariCODE"))
    '
    strSql = "INSERT INTO Moshtari "
    strSql = strSql & "VALUES(" & CodeMoshtari & ",'"
    strSql = strSql & Trim(TxtName) & "','" & Trim(TxtPhone) & "')"
    rs.Open strSql, CNS
    '
    With Grid1
       .AddItem ""
       .Cell(.Rows - 1, 1).Text = Trim(TxtPhone)
       .Cell(.Rows - 1, 2).Text = Trim(TxtName)
       .Cell(.Rows - 1, 3).Text = .Rows - 1
       .Cell(.Rows - 1, 4).Text = CodeMoshtari ' hidden
    End With
    MsgBox "«ÿ·«⁄«  ’«Õ» ò«·« À»  ‘œ", vbInformation, ""
    
    ClearText Me
    TxtName.SetFocus
    '
    Set rs = Nothing
 End If
End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackForms7.jpg")
 '
 Call SetGrid
 Call LoadMoshtari
End Sub

Private Sub Grid1_DblClick()
 MsgBox Grid1.Cell(Grid1.ActiveCell.Row, 1).Text
End Sub

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 5
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
      .Column(1).Width = 120
      .Column(2).Width = 155
      .Column(3).Width = 50
      .Column(4).Width = 0 ' Code Moshtari
      '
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      'Make Titr
      .Cell(0, 1).Text = "‘„«—Â  ·›‰"
      .Cell(0, 2).Text = "‰«„ ’«Õ» ò«·«"
      .Cell(0, 3).Text = "—œÌ›"
      .Cell(0, 4).Text = ""
      '
      .ReadOnly = True
      .ReadOnlyFocusRect = Solid
      '
      .SelectionMode = cellSelectionFree
      .Appearance = Flat
      
 End With
End Sub

Private Sub TxtName_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 
End Sub

Private Sub TxtPhone_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If

End Sub

Private Sub LoadMoshtari()
 Dim rs As New Recordset
 Dim strSql As String
 Dim r As Long
 '
 strSql = "SELECT * FROM Moshtari "
 rs.Open strSql, CNS
 r = 1
 Do While Not rs.EOF
    Grid1.AddItem rs(2) & vbTab & rs(1) & vbTab & r & vbTab & rs(0)
    rs.MoveNext
    r = r + 1
 Loop
 rs.Close
 Set rs = Nothing
 
End Sub


Private Sub MakeReport(dateX As String, MoshtariCode As Byte, MoshtariName As String)
 Dim strSql As String
 Dim rs As New Recordset
 Dim i As Integer, kRow As Integer
 Dim sumVazn As Long
 '
 strSql = "SELECT Name,Etebar,Parvane,Part,BarName, "
 strSql = strSql & "Tarikh,Address,Havale,ShomareMashin, "
 strSql = strSql & "Vazn ,Tedad,Size0,Keraye,Mobile,Parvande, Count0,Main7.Code "
 strSql = strSql & "FROM Main7 INNER JOIN Detail7 ON Main7.Code = Detail7.Code "
 strSql = strSql & "WHERE Detail7.MoshtariCODE=" & MoshtariCode
 strSql = strSql & " AND Tarikh='" & dateX & "'"
 strSql = strSql & " ORDER BY Main7.Code,Count0 "
 
 rs.Open strSql, CNS
 
 If Not rs.EOF Then
    With FrmPreview.Grid1
        .OpenFile App.Path & "\Rep7Almas.cel"
        .Cell(1, 4).Text = MoshtariName & "-Õ„· —Ê“«‰Â"
        kRow = 3
        sumVazn = 0
        Do While Not rs.EOF
           For i = 0 To 14
               .Cell(kRow, i + 1).Text = IIf(IsNull(rs(14 - i)), "", rs(14 - i))
           Next
           sumVazn = sumVazn + rs("Vazn")
           .Cell(kRow, 16).Text = kRow - 2
           kRow = kRow + 1
           .InsertRow kRow, 1
           rs.MoveNext
        Loop
        rs.Close
        Call Molahezat(sumVazn, MoshtariName)
        Call PageSetupANDFooter

        .Cell(1, 2).Text = .Cell(1, 2).Text & Space(8) & FrmGetReportINFO7.FarDate1.today
        
        .PrintPreview 100
        FrmPreview.Show 1
    End With
Else
    MsgBox "ê“«—‘Ì »—«Ì «Ì‰ ’«Õ» ò«·« ÊÃÊœ ‰œ«—œ", vbExclamation, ""
End If

End Sub

Private Sub Molahezat(VazneVarede As Long, Moshtari As String)
 Dim r As Integer

 With FrmPreview.Grid1
      r = .Rows - 2
     .Range(r, 4, r, 14).Merge
     .Cell(r, 4).Alignment = cellCenterCenter
     .Cell(r, 4).Font.Name = "Arial"
     .Cell(r, 4).Font.Bold = True
     .Cell(r, 4).Font.Size = 13
     .Cell(r, 4).Text = ".„·«ÕŸ« : »—«Ì ’«Õ» ò«·« " & Moshtari & _
                        " »Â Ê“‰ " & VazneVarede & " òÌ·Êê—„ Œ«—Ã ‘œÂ «”  "
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


