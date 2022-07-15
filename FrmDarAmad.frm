VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmDarAmad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
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
   ScaleHeight     =   6405
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtVariz 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "B Yekan"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox TxtDaryafti 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "B Yekan"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox TxtBarNamEDarya 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox TxtKeshti 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5520
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
      MICON           =   "FrmDarAmad.frx":0000
      PICN            =   "FrmDarAmad.frx":001C
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
      Left            =   2160
      TabIndex        =   5
      Top             =   5520
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
      MICON           =   "FrmDarAmad.frx":3A8A
      PICN            =   "FrmDarAmad.frx":3AA6
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
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ﬁœ«— „»·€ Ê«—Ì“Ì"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3360
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ì‹‹‹‹«"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   570
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ﬁœ«— „»·€ œ—Ì«› Ì"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   2115
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ Ê«—Ì“ Ì« œ—Ì«› "
      Height          =   405
      Left            =   6450
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â »«—‰«„Â œ—Ì«ÌÌ"
      Height          =   405
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò‹‘ Ì"
      Height          =   405
      Left            =   7410
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   945
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â À»  «ÿ·«⁄«  œ—¬„œ »«—‘„«—Ì ò«”ÅÌ‰ Œ‹‹“—"
      Height          =   405
      Left            =   765
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   105
      Width           =   7710
   End
   Begin VB.Image ImgBackground 
      Height          =   6420
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "FrmDarAmad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()
 If Not CheckValidate Then Exit Sub
 
 Dim StrSql As String
 Dim rs As New Recordset
 Dim LastCount As Long, codeBarShomari As Byte
 '
 codeBarShomari = 1 ' chon barShomari digar nist pas CASPIAN KHAZAR 1 mishavad
 '
 StrSql = "SELECT MAX(Count0) FROM DarAmad "
 StrSql = StrSql & "WHERE CodeBarShomari=" & codeBarShomari
 rs.Open StrSql, CNS
 LastCount = IIf(IsNull(rs(0)), 1, rs(0) + 1)
 rs.Close
 '
 StrSql = "INSERT INTO DarAmad "
 StrSql = StrSql & "VALUES(" & codeBarShomari & "," & LastCount & ","
 StrSql = StrSql & "'" & Trim(TxtKeshti) & "',"
 StrSql = StrSql & "'" & TxtBarNamEDarya & "',"
 StrSql = StrSql & "'" & Mid(FarDate1.Text, 3) & "',"
 
 If Text2Currency(TxtVariz) > 0 Then
    StrSql = StrSql & Text2Currency(TxtVariz) & ",0)"
 ElseIf Text2Currency(TxtDaryafti) > 0 Then
    StrSql = StrSql & "0," & Text2Currency(TxtDaryafti) & ")"
 End If
 
 rs.Open StrSql, CNS
 '
 MsgBox "«ÿ·«⁄«  ‘„« À»  ‘œ", vbInformation, ""
 ClearText Me
 TxtKeshti.SetFocus
 Set rs = Nothing
End Sub

Private Sub FarDate1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    SendKeys "{Tab}"
    KeyCode = 0
 End If
End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 FarDate1.Text = FarDate1.today
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackFormsHazine.jpg")
End Sub
 
Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub TxtBarNamEDarya_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtDaryafti_GotFocus()
  TxtDaryafti = Format(TxtDaryafti)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtDaryafti_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtDaryafti_LostFocus()
 TxtDaryafti = Format(TxtDaryafti, "#,#—Ì«·")
End Sub

Private Sub TxtKeshti_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
End Sub

Private Sub TxtKeshti_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtVariz_GotFocus()
  TxtVariz = Format(TxtVariz)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtVariz_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Function CheckValidate() As Boolean
 CheckValidate = True
 If TxtKeshti = Empty Then
    MsgBox "·ÿ›« ‰«„ ò‘ Ì —«  «ÌÅ ‰„«ÌÌœ", vbExclamation, ""
    TxtKeshti.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If TxtBarNamEDarya = Empty Then
    MsgBox "·ÿ›« ‘„«—Â »«—‰«„Â œ—Ì«ÌÌ —« „‘Œ’ ò‰Ìœ", vbExclamation, ""
    TxtBarNamEDarya.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If FarDate1.Text = Empty Then
    MsgBox "·ÿ›«  «—ÌŒ —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    FarDate1.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If Text2Currency(TxtVariz) + Text2Currency(TxtDaryafti) = 0 Then
    MsgBox "ÌòÌ «“ ò«œ— Â«Ì œ—Ì«› Ì Ì« Ê«—Ì“Ì »«Ìœ ò«„· ‘Êœ", vbExclamation, ""
    TxtVariz.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If Text2Currency(TxtVariz) <> 0 And Text2Currency(TxtDaryafti) <> 0 Then
    MsgBox "·ÿ›« ›ﬁÿ ÌòÌ «“ ò«œ— Â«—« ò«„· ‰„«ÌÌœ", vbExclamation, ""
    TxtVariz.SetFocus
    CheckValidate = False
    Exit Function
 End If
 
End Function

Private Sub TxtVariz_LostFocus()
 TxtVariz = Format(TxtVariz, "#,#—Ì«·")
End Sub
