VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmRepGardesh 
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
   Begin VB.ComboBox CombMoshtariCode 
      Height          =   525
      Left            =   2400
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CombMoshtari 
      Height          =   525
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "CombMoshtari"
      Top             =   1680
      Width           =   2895
   End
   Begin HaftAlmas.TypeButton TypeButton1 
      Height          =   495
      Left            =   480
      TabIndex        =   9
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
      MICON           =   "FrmRepGardesh.frx":0000
      PICN            =   "FrmRepGardesh.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   7
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
      MICON           =   "FrmRepGardesh.frx":3A8A
      PICN            =   "FrmRepGardesh.frx":3AA6
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
      TabIndex        =   4
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
      MICON           =   "FrmRepGardesh.frx":7514
      PICN            =   "FrmRepGardesh.frx":7530
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
      TabIndex        =   0
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
      TabIndex        =   1
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
   Begin FlexCell.Grid Grid1 
      Height          =   2295
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4048
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
      EnterKeyMoveTo  =   1
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â œ—Ì«›   «—ÌŒ Ê ‘„«—Â Å—Ê«‰Â »—«Ì ê“«—‘ ê—œ‘ ò«—"
      Height          =   405
      Left            =   2310
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   5325
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Å—Ê«‰Â"
      Height          =   405
      Left            =   6150
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Image ImgBackground 
      Height          =   5910
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’«Õ» ò«·«"
      Height          =   405
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      Height          =   405
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   8
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
      TabIndex        =   6
      Top             =   960
      Width           =   1410
   End
End
Attribute VB_Name = "FrmRepGardesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbIgnoreListClick As Boolean

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdOK7_Click()
' If ChkParvane.Value = 0 And Not CheckDate Then Exit Sub
' If ChkParvane.Value = 0 And FarDate1.Text = Empty Then Exit Sub
' If ChkParvane.Value = 1 And FarDate1.Text = Empty And Grid1.Cell(1, 1).Text = Empty Then Exit Sub
 '
 Dim strSql As String
 Dim rs As New Recordset
 Dim i As Integer, kRow As Integer
 Dim sumTedad As Long, sumVazn As Long
 Dim CurrentParvane As String
 Dim SubStr As String
 '
 If Grid1.Cell(Grid1.Rows - 1, 1).Text = Empty Then
    Grid1.RemoveItem Grid1.Rows - 1
 End If
 
 
 SubStr = "IN("
 For i = 1 To Grid1.Rows - 1
     SubStr = SubStr & "'" & Grid1.Cell(i, 1).Text & "',"
 Next
 SubStr = Left(SubStr, Len(SubStr) - 1)
 SubStr = SubStr & ")"
'''''''''''''
 strSql = "SELECT Gardesh.MoshtariCODE, Gardesh.Parvane, Gardesh.NoeKala, "
 strSql = strSql & "Gardesh.Tedad, Gardesh.Tonaj, Gardesh.Molahezat, "
 strSql = strSql & "(select Sum(tedad) From Detail7 where detail7.Code=main7.code "
 strSql = strSql & "AND (Tarikh BETWEEN '" & Mid(FarDate1.Text, 3) & "' AND '"
 strSql = strSql & Mid(FarDate2.Text, 3) & "')) AS TT,"
 strSql = strSql & "(select Sum(Vazn) From Detail7 where detail7.Code=main7.code "
 strSql = strSql & "AND (Tarikh BETWEEN '" & Mid(FarDate1.Text, 3) & "' AND '"
 strSql = strSql & Mid(FarDate2.Text, 3) & "')) AS TV "
 strSql = strSql & "FROM Gardesh INNER JOIN Main7 ON Gardesh.Parvane = Main7.Parvane"
 strSql = strSql & " WHERE ((Gardesh.Parvane " & SubStr & "))"
 rs.Open strSql, CNS
 '

   '''''''''''''''''
   If rs.EOF Then
      MsgBox "ê“«—‘ »—«Ì «Ì‰ „Ê—œ ÊÃÊœ ‰œ«—œ", vbExclamation, ""
      rs.Close
      Set rs = Nothing
      Exit Sub
   End If
    With FrmPreview.Grid1
        .OpenFile App.Path & "\RepGardeshParvane.cel"
        kRow = 4
        sumTedad = 0: sumVazn = 0
        .Cell(2, 7).Text = CombMoshtari
        .Cell(1, 2).Text = FarDate1.Text & " «  «—ÌŒ" & FarDate2.Text
        .Cell(1, 1).Text = " «—ÌŒ ’œÊ—  " & FarDate1.today
        On Error Resume Next
       For i = 1 To Grid1.Rows - 1
        .Cell(kRow, 1).Text = IIf(IsNull(rs(5)), "", rs(5))
        .Cell(kRow, 2).Text = rs(4) - IIf(IsNull(rs(7)), 0, rs(7)) ' Tonaj Baghimande
        .Cell(kRow, 3).Text = IIf(IsNull(rs(7)), 0, rs(7)) ' Tonaj Rafte
        .Cell(kRow, 4).Text = rs(4) ' Tonaj KOL
        .Cell(kRow, 5).Text = rs(3) - IIf(IsNull(rs(6)), 0, rs(6)) ' ' Tedad Baghimande
        .Cell(kRow, 6).Text = IIf(IsNull(rs(6)), 0, rs(6)) '  ' Tedad Rafte
        .Cell(kRow, 7).Text = rs(3) ' Tedad KOL
        .Cell(kRow, 8).Text = rs(2) ' Noe Kala
        .Cell(kRow, 9).Text = rs(1) ' Parvane
        kRow = kRow + 1
        .AddItem ""
        rs.MoveNext
       Next
        rs.Close
        'Call Molahezat
        'Call TedadBaghimande(SubStr)
        .Range(4, 1, kRow, .Cols - 1).Borders(cellInsideVertical) = cellThin
        .Range(4, 1, kRow, .Cols - 1).Borders(cellInsideHorizontal) = cellThin
        .Range(4, 1, kRow, .Cols - 1).Borders(cellEdgeBottom) = cellThick
        .Range(4, 1, kRow, .Cols - 1).Borders(cellEdgeTop) = cellThick
        .Range(4, 1, kRow, .Cols - 1).Borders(cellEdgeLeft) = cellThick
        .Range(4, 1, kRow, .Cols - 1).Borders(cellEdgeRight) = cellThick
        
        .Range(4, 1, kRow, .Cols - 1).Alignment = cellCenterCenter
        .Range(4, 1, kRow, .Cols - 1).FontName = "B Nazanin"
        .Range(4, 1, kRow, .Cols - 1).FontBold = True
        .Range(4, 1, kRow, .Cols - 1).FontSize = 12
        
        For i = 4 To .Rows - 1
            .RowHeight(i) = 25
        Next
        .AddItem ""
        .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).Merge
        .Cell(.Rows - 1, 1).Alignment = cellCenterCenter
        .Cell(.Rows - 1, 1).Font.Name = "B Titr"
        .Cell(.Rows - 1, 1).Font.Bold = True
        .Cell(.Rows - 1, 1).Font.Size = 10
        .RowHeight(.Rows - 1) = 35
        .Cell(.Rows - 1, 1).Text = "«—”«·Ì «“ «‰»«— ÅÊÌ«( »Õ—Ì )-‘«Ì‹«‰ „‹Â—"
        
        'Call PageSetupANDFooter
        

        .PrintPreview 110
        FrmPreview.Show 1
    End With
 'Else
 '   rs.Close
 '   MsgBox "ê“«—‘Ì »—«Ì  «—ÌŒ „Ê—œ ‰Ÿ— „ÊÃÊœ ‰„Ì »«‘œ", vbInformation, ""
 'End If
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
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackForms7.jpg")
 '
 FarDate1.Text = FarDate1.today
 FarDate2.Text = FarDate2.today
 '
 
 '
 Dim rs As New Recordset
 rs.Open "SELECT * FROM Moshtari ORDER BY MoshtariName", CNS
 Do While Not rs.EOF
    CombMoshtari.AddItem rs("MoshtariName")
    CombMoshtariCode.AddItem rs("MoshtariCODE")
    rs.MoveNext
 Loop
 rs.Close
 Set rs = Nothing
  '
 Call SetGrid

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
     .Cell(r, 4).Font.Name = "Arial"
     .Cell(r, 4).Font.Bold = True
     .Cell(r, 4).Font.Size = 13
     .Cell(r, 4).Text = ".„·«ÕŸ« : »« ”·«„ Ê Œ” Â ‰»«‘Ìœ " & _
               "Ã„⁄ ò· ÕÊ«·Â Â«Ì »«—êÌ—Ì œ— «Ì‰  «—ÌŒ " & Mid(FarDate1.Text, 3) & _
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

