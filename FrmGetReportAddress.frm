VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmGetReportAddress 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
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
   ScaleHeight     =   4110
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtHavale 
      Alignment       =   2  'Center
      Height          =   525
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox TxtAddress 
      Alignment       =   2  'Center
      Height          =   525
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2400
      Width           =   4815
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   3360
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
      MICON           =   "FrmGetReportAddress.frx":0000
      PICN            =   "FrmGetReportAddress.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdOK 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
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
      MICON           =   "FrmGetReportAddress.frx":3A8A
      PICN            =   "FrmGetReportAddress.frx":3AA6
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ÕÊ«·Â"
      Height          =   405
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò‹‹œ ¬œ—”"
      Height          =   405
      Left            =   6285
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ê“«—‘ «“  «—ÌŒ "
      Height          =   405
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1410
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â œ—Ì«›   «—ÌŒ Ê ¬œ—”  »—«Ì ‰„«Ì‘ ê“«—‘"
      Height          =   405
      Left            =   555
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   7080
   End
   Begin VB.Image ImgBackground 
      Height          =   4140
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      Height          =   405
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   645
   End
End
Attribute VB_Name = "FrmGetReportAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
  Unload Me
End Sub

Private Sub CmdOK7_Click()

End Sub

Private Sub CmdOK_Click()
 If Not CheckDate Then Exit Sub
 '
 Dim strSql As String
 Dim rs As New Recordset
 Dim i As Integer, kRow As Integer
 Dim sumTedad As Long, sumVazn As Long
 '
 If ((FarDate1.Text & FarDate1.Text) <> Empty) And (TxtAddress <> Empty) And (TxtHavale <> Empty) Then
    strSql = "SELECT Name,Etebar,Parvane,Part,BarName, "
    strSql = strSql & "Tarikh,Address,Havale,ShomareMashin, "
    strSql = strSql & "Vazn ,Tedad,Size0,Keraye,Mobile,Parvande, Count0,Main7.Code "
    strSql = strSql & "FROM Main7 INNER JOIN Detail7 ON Main7.Code = Detail7.Code "
    strSql = strSql & "WHERE (((Detail7.Tarikh) BETWEEN '" & Mid(FarDate1.Text, 3) & "' "
    strSql = strSql & "AND '" & Mid(FarDate2.Text, 3) & "')) "
    strSql = strSql & "AND Address='" & Trim(TxtAddress) & "' "
    strSql = strSql & "AND Havale='" & TxtHavale & "' "
    strSql = strSql & "ORDER BY Main7.Code,Count0 "
 End If

 
 rs.Open strSql, CNS
 If Not rs.EOF Then
    With FrmPreview.Grid1
        .OpenFile App.Path & "\Rep7Almas.cel"
        kRow = 3
        Do While Not rs.EOF
           For i = 0 To 14
               .Cell(kRow, i + 1).Text = IIf(IsNull(rs(14 - i)), "", rs(14 - i))
           Next
           .Cell(kRow, 16).Text = kRow - 2
           sumTedad = sumTedad + rs("Tedad")
           sumVazn = sumVazn + rs("Vazn")
           kRow = kRow + 1
           .InsertRow kRow, 1
           rs.MoveNext
        Loop
        rs.Close
        Call Molahezat(sumTedad, sumVazn, kRow - 3)
       ' Call TedadBaghimande
        Call PageSetupANDFooter
        
        .PrintPreview 110
        FrmPreview.Show 1
    End With
 Else
    rs.Close
    MsgBox "ê“«—‘Ì »—«Ì  «—ÌŒ „Ê—œ ‰Ÿ— „ÊÃÊœ ‰„Ì »«‘œ", vbInformation, ""
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
  If FarDate2.Text < FarDate1.Text Then
     MsgBox " «—ÌŒ œÊ„ ‰»«Ìœ «“  «—ÌŒ «Ê· ò„ — »«‘œ", vbExclamation, ""
     FarDate2.SetFocus
     CheckDate = False
     Exit Function
  End If
End Function

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

Private Sub Molahezat(Tedad As Long, Vazn As Long, Traily As Long)
 Dim r As Integer
 With FrmPreview.Grid1
      r = .Rows - 2
     .Range(r, 4, r, 14).Merge
     .Cell(r, 4).Alignment = cellCenterCenter
     .Cell(r, 4).Font.Name = "Arial"
     .Cell(r, 4).Font.Bold = True
     .Cell(r, 4).Font.Size = 11
     .Cell(r, 4).Text = " ⁄œ«œ " & Traily & " œ” ê«Â  —Ì·— ‘«„·  " & Tedad & " »‰œ· »Â Ê“‰ ò· " & Vazn & " òÌ·Êê—„ Õ„· ê—œÌœ"
 End With
End Sub

Private Sub TxtAddress_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtHavale_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789/-" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub
