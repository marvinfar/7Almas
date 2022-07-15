VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmGetReportDarAmad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
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
   ScaleHeight     =   3675
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtKeshti 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1680
      Width           =   2055
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2880
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
      MICON           =   "FrmGetReportDarAmad.frx":0000
      PICN            =   "FrmGetReportDarAmad.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdOKDarAmad 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2880
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
      MICON           =   "FrmGetReportDarAmad.frx":3A8A
      PICN            =   "FrmGetReportDarAmad.frx":3AA6
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
      Left            =   4080
      TabIndex        =   3
      Top             =   840
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
      Left            =   1200
      TabIndex        =   4
      Top             =   840
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
   Begin HaftAlmas.TypeButton CmdOkHazine 
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   2880
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
      MICON           =   "FrmGetReportDarAmad.frx":707A
      PICN            =   "FrmGetReportDarAmad.frx":7096
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò‘ Ì"
      Height          =   405
      Left            =   6765
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â  ê“«—‘ œ— ¬„œ »«—‘„«—Ì"
      Height          =   405
      Left            =   330
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   -50
      Width           =   7305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ê“«—‘ «“  «—ÌŒ "
      Height          =   405
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      Height          =   405
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   645
   End
   Begin VB.Image ImgBackground 
      Height          =   3660
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "FrmGetReportDarAmad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WichReport As Byte ' 1 Caspian -- 2 GOL
Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdOKDarAmad_Click()
 Call REPORTDarAmad
End Sub

Private Sub CmdOkHazine_Click()
 Dim PlaceName As String
 Dim strSql As String
 Dim rs As New Recordset
 Dim SumHazine As Currency
 Dim kRow As Integer
 Dim D1 As String * 8, D2 As String * 8
 '
 D1 = Mid(FarDate1.Text, 3)
 D2 = Mid(FarDate2.Text, 3)
 If WichReport = 1 Then PlaceName = "»«—‘„«—Ì ò‹«”ÅÌ‰ Œ“—"
 If WichReport = 2 Then PlaceName = "ò‘ Ì—«‰Ì ê‹·"
 '
 If Trim(D1 & D2) = Empty And TxtKeshti = Empty Then
    MsgBox "·ÿ›« ÌòÌ «“ ‘—«Ìÿ  «—ÌŒ Ì« ‘—Õ Â“Ì‰Â —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    FarDate1.SetFocus
    Exit Sub
 End If
 ''
 strSql = "SELECT * FROM Hazine "
 strSql = strSql & "WHERE CodePlace= " & WichReport & " "
 If TxtKeshti <> Empty Then
    strSql = strSql & "AND Description LIKE '%" & TxtKeshti & "%' "
 End If
 If Trim(D1 & D2) <> Empty Then
    strSql = strSql & "AND (Tarikh BETWEEN '" & D1 & "' AND '" & D2 & "') "
 End If
 strSql = strSql & " ORDER BY Count0 "
 ''''
 rs.Open strSql, CNS
 If rs.EOF Then
    MsgBox "ò“«—‘ Â‹“Ì‰Â „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation, ""
    rs.Close
 Else '''
 
    With FrmPreview.Grid1
      .OpenFile App.Path & "\RepHazine.cel"
      '
      .Cell(1, 7).Text = PlaceName
      If Trim(D1 & D2) <> Empty Then
         .Cell(1, 4).Text = "ê“«—‘ Â“Ì‰Â Â«Ì «‰Ã«„ ‘œÂ œ—  «—ÌŒ"
         .Cell(1, 3).Text = D1
         .Cell(1, 1).Text = D2
      Else
         .Cell(1, 4).Text = "ê“«—‘ Â“Ì‰Â Â«Ì «‰Ã«„ ‘œÂ »« ‘‹—Õ"
         .Cell(1, 3).Text = TxtKeshti
         .Cell(1, 2).Text = ""
      End If
      '
      kRow = 5: 'i = 1
      SumHazine = 0
      Do While Not rs.EOF
         .Cell(kRow, 9).Text = kRow - 4
         .Cell(kRow, 6).Text = IIf(IsNull(rs("Description")), "", rs("Description"))
         .Cell(kRow, 5).Text = IIf(IsNull(rs("Tarikh")), "", rs("Tarikh"))
         .Cell(kRow, 3).Text = Format(rs("Mablagh"), "#,#—Ì«·")
         SumHazine = SumHazine + rs("Mablagh")
         '
         kRow = kRow + 1
         .InsertRow kRow, 1
         .Range(kRow, 1, kRow, 2).Merge
         .Range(kRow, 3, kRow, 4).Merge
         .Range(kRow, 6, kRow, 8).Merge
         '
         rs.MoveNext
         'i = i + 1
      Loop
      .Cell(kRow, 3).Text = "Ã„⁄ Â“Ì‰Â Â«Ì «‰Ã«„ ‘œÂ"
      .Cell(kRow, 1).Text = Format(SumHazine, "#,#—Ì«·")
      rs.Close
      '
      .PrintPreview 100
      'FrmPreview.Show 1
   End With
 End If
 
 Set rs = Nothing
End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackFormsHazine.jpg")
 '
 FarDate1.Text = FarDate1.today
 FarDate2.Text = FarDate2.today
End Sub

Private Sub Label1_DblClick()
 CmdClose_Click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub REPORTDarAmad()
 If Not CheckDate Then Exit Sub
 '''
 Dim strSql As String
 Dim rs As New Recordset
 Dim D1 As String * 8, D2 As String * 8
 Dim Count0D1 As Long, Count0D2 As Long
 Dim Mande As Currency
 Dim repName As String
  
 '
 D1 = Mid(FarDate1.Text, 3): D2 = Mid(FarDate2.Text, 3)
 
 If TxtKeshti <> Empty Then
    strSql = "SELECT * FROM DarAmad "
    strSql = strSql & "WHERE CodeBarShomari=" & 1
    strSql = strSql & " AND Keshti='" & TxtKeshti & "' "
    strSql = strSql & "ORDER BY Count0 "
    rs.Open strSql, CNS
    repName = " ê“«—‘ œ—¬„‹œ »«—‘„«—Ì »—«”«” ò‘ Ì"
    If Not rs.EOF Then
       Call ReportBodyMaker(rs, repName, Mande, "", "", 0)
    Else
       rs.Close
       MsgBox "»—«Ì «Ì‰ »«—‘„«—Ì ê“«—‘Ì „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation, ""
    End If
    Exit Sub
 End If
 
 ''''''''''
 strSql = "SELECT MIN(Count0),MAX(Count0) FROM DarAmad "
 strSql = strSql & "WHERE (((Tarikh) BETWEEN '" & D1 & "' AND '" & D2 & "') "
 strSql = strSql & "AND ((CodeBarShomari)=" & 1 & "))"
 
 rs.Open strSql, CNS
 If Not rs.EOF Then  ' if found
    Count0D1 = IIf(IsNull(rs(0)), 0, rs(0))
    Count0D2 = IIf(IsNull(rs(1)), 0, rs(1))
    rs.Close
    If Count0D1 = 0 Or Count0D2 = 0 Then
       MsgBox "»—«Ì «Ì‰ »«—‘„«—Ì ê“«—‘Ì „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation, ""
       Set rs = Nothing
       Exit Sub
    End If
    
    Mande = CalcMande(Count0D1) ' function
    ''
    strSql = "SELECT * FROM DarAmad "
    strSql = strSql & "WHERE CodeBarShomari=" & 1
    strSql = strSql & " AND (Count0 BETWEEN " & Count0D1
    strSql = strSql & " AND " & Count0D2 & ") "
    strSql = strSql & " ORDER BY Count0 "
    rs.Open strSql, CNS
    repName = "ê“«—‘ œ—¬„‹œ »«—‘„«—Ì œ— „ÕœÊœÂ  «—ÌŒ "
    Call ReportBodyMaker(rs, repName, Mande, D1, D2, Count0D1)
 Else
    MsgBox "ê“«—‘Ì œ— „ÕœÊœÂ  «—ÌŒ œ«œÂ ‘œÂ „ÊÃÊœ ‰„Ì »«‘œ", vbExclamation, ""
    rs.Close
 End If
 
 Set rs = Nothing

End Sub

Private Function CheckDate() As Boolean
  CheckDate = True
  If FarDate2.Text < FarDate1.Text Then
     MsgBox " «—ÌŒ œÊ„ ‰»«Ìœ «“  «—ÌŒ «Ê· ò„ — »«‘œ", vbExclamation, ""
     FarDate2.SetFocus
     CheckDate = False
     Exit Function
  End If
End Function

Private Function CalcMande(C As Long) As Currency
 Dim strSql As String
 Dim rs As New Recordset
 
  If C = 1 Then ' az avval gozaresh giri,, mande mishavad avalin bestankari
     strSql = "SELECT Varizi FROM DarAmad "
     strSql = strSql & "WHERE CodeBarShomari=" & 1
     strSql = strSql & " AND Count0=1"
     rs.Open strSql, CNS
     CalcMande = rs(0)
     rs.Close
  Else
     strSql = "SELECT SUM(Varizi)-SUM(Daryafti) FROM DarAmad "
     strSql = strSql & "WHERE CodeBarShomari=" & 1
     strSql = strSql & " AND (Count0 BETWEEN 1 AND " & C - 1 & ") "
     rs.Open strSql, CNS
     CalcMande = rs(0)
     rs.Close
  End If
  Set rs = Nothing
End Function

Private Sub ReportBodyMaker(rs As Recordset, ReportName As String, ByRef Mande As Currency, D1 As String, D2 As String, Count0D1 As Long)
 Dim Daryafti As Currency, Varizi As Currency
 Dim SumDaryafti As Currency, SumVarizi As Currency
 Dim AccountName  As String

    With FrmPreview.Grid1
         Dim kRow As Integer
         AccountName = "ò‹«”ÅÌ‰ Œ‹“—"

         .OpenFile App.Path & "\RepDarAmad.cel"
         .Cell(1, 1).Text = AccountName
         .Cell(1, 5).Text = ReportName
         If Trim(D1 & D2) = Empty Then
            .Cell(1, 3).Text = TxtKeshti
         Else
            .Cell(1, 3).Text = D1 & " « " & D2
         End If
         
         If Count0D1 > 1 Then
            .InsertRow 4, 1
            .Cell(4, 1).Text = Format(Mande, "#,#")
            .Cell(4, 2).Text = "„«‰œÂ «“ ﬁ»·"
            kRow = 6
         Else
            '.Cell(5, 2).Text = ""
            kRow = 5
         End If
         
         SumDaryafti = 0: SumVarizi = 0
         Do While Not rs.EOF
            .InsertRow kRow, 1
            Daryafti = IIf(IsNull(rs("Daryafti")), 0, rs("Daryafti"))
            Varizi = IIf(IsNull(rs("Varizi")), 0, rs("Varizi"))
            
            If Count0D1 > 1 Then Mande = Mande + Varizi - Daryafti
            
            SumDaryafti = SumDaryafti + Daryafti
            SumVarizi = SumVarizi + Varizi
            
            .Cell(kRow, 1).Text = Format(Mande, "#,#")
            .Cell(kRow, 2).Text = Format(Daryafti, "#,#")
            .Cell(kRow, 3).Text = Format(Varizi, "#,#")
            .Cell(kRow, 4).Text = rs("Tarikh")
            .Cell(kRow, 5).Text = rs("BarNamEDarya")
            .Cell(kRow, 6).Text = rs("Keshti")
            .Cell(kRow, 7).Text = kRow - 4
            rs.MoveNext
            kRow = kRow + 1
            
            If Not rs.EOF Then ' baraye inke akharin bar Error nade
               Daryafti = IIf(IsNull(rs("Daryafti")), 0, rs("Daryafti"))
               Varizi = IIf(IsNull(rs("Varizi")), 0, rs("Varizi"))
               Mande = Mande + Varizi - Daryafti
            End If
         Loop
         rs.Close
         '
         .AddItem ""
         .Cell(.Rows - 1, 2).Text = Format(SumDaryafti, "#,#")
         .Cell(.Rows - 1, 3).Text = Format(SumVarizi, "#,#")
         .Cell(.Rows - 1, 4).Text = "”ÿ— „Ã„Ê⁄"
         .Cell(.Rows - 1, 4).Border(cellEdgeLeft) = cellThin
         ''
         .RowHeight(.Rows - 1) = 32
         .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).Alignment = cellCenterCenter
         .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).FontName = "B Nazanin"
         .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).FontBold = True
         .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).FontSize = 12
         .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).BackColor = QBColor(7)
         '
         .PrintPreview 100
    End With

End Sub
