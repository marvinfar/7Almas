VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmTransactionBank 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
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
   ScaleHeight     =   5715
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtDescription 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   6855
   End
   Begin VB.TextBox TxtBedehkar 
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
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox TxtBestankar 
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
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3120
      Width           =   2895
   End
   Begin VB.ComboBox CombCodeBank 
      Height          =   525
      Left            =   3960
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.ComboBox CombAccountName 
      Height          =   525
      Left            =   4560
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   4920
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
      MICON           =   "FrmTransactionBank.frx":0000
      PICN            =   "FrmTransactionBank.frx":001C
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
      Left            =   2400
      TabIndex        =   4
      Top             =   4920
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
      BCOL            =   16761087
      BCOLO           =   12583104
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTransactionBank.frx":3A8A
      PICN            =   "FrmTransactionBank.frx":3AA6
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
      Left            =   600
      TabIndex        =   6
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
   Begin VB.Label LblMande 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„«‰œÂ Õ”«» - ‰«„‘Œ’"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1560
      Width           =   2325
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘—Õ ⁄„·Ì« "
      Height          =   405
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ﬁœ«— „»·€ »—œ«‘  «“ Õ”«»"
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
      Left            =   5820
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label3 
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
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ﬁœ«— „»·€ Ê«—Ì“Ì »Â Õ”«»"
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
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   3105
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «„—Ê“"
      Height          =   405
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Õ”«» Ì« »«‰ò"
      Height          =   405
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1560
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â À»   —«ò‰‘Â« Ê ⁄„·Ì«  »—œ«‘  Ê Ê«—Ì“ «“ Õ”«»"
      Height          =   405
      Left            =   4200
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   20
      Width           =   4935
   End
   Begin VB.Image ImgBackground 
      Height          =   5700
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "FrmTransactionBank"
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
 '
 Dim StrSql As String
 Dim rs As New Recordset
 Dim LastCount As Long, CodeBank As Byte
 '
 CodeBank = Val(CombCodeBank.Text)
 StrSql = "SELECT MAX(Count0) FROM TransactionBank "
 StrSql = StrSql & "WHERE CodeBank=" & CodeBank
 rs.Open StrSql, CNS
 LastCount = IIf(IsNull(rs(0)), 1, rs(0) + 1)
 rs.Close
 '
 StrSql = "INSERT INTO TransactionBank "
 StrSql = StrSql & "VALUES(" & CodeBank & "," & LastCount & ","
 StrSql = StrSql & "'" & Mid(FarDate1.Text, 3) & "',"
 StrSql = StrSql & "'" & Trim(TxtDescription) & "',"
 
 If Text2Currency(TxtBestankar) > 0 Then
    StrSql = StrSql & Text2Currency(TxtBestankar) & ",0)"
 ElseIf Text2Currency(TxtBedehkar) > 0 Then
    StrSql = StrSql & "0," & Text2Currency(TxtBedehkar) & ")"
 End If
 
 rs.Open StrSql, CNS
 '
 MsgBox "«ÿ·«⁄«  ‘„« À»  ‘œ", vbInformation, ""
 ClearText Me
 LblMande = "„«‰œÂ Õ”«» - ‰«„‘Œ’"
 CombAccountName.SetFocus
 Set rs = Nothing
End Sub

Private Sub CombAccountName_Click()
 CombCodeBank.ListIndex = CombAccountName.ListIndex
 If CombAccountName.ListIndex <> -1 Then
    Dim rs As New Recordset
    Dim StrSql As String
    '
    StrSql = "SELECT SUM(Bestankar)-SUM(Bedehkar) FROM TransactionBank "
    StrSql = StrSql & "WHERE CodeBank=" & Val(CombCodeBank.Text)
    rs.Open StrSql, CNS
    If Not rs.EOF Then
       LblMande = "„‹‹«‰‹‹œÂ Õ”‹‹«»  " & Format(IIf(IsNull(rs(0)), 0, rs(0)), "#,# —Ì«·")
    End If
    rs.Close
    Set rs = Nothing
 End If
End Sub

Private Sub CombAccountName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If

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
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackFormsBanki.jpg")
 FarDate1.Text = FarDate1.today
 '
 Call LoadAccountName
 '
 
End Sub

Private Sub LblTitle_DblClick()
  CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub LoadAccountName()
 Dim StrSql As String
 Dim rs As New Recordset
 '
 StrSql = "SELECT CodeBank,AccountName FROM DefBank "
 rs.Open StrSql, CNS
 
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

Private Function CheckValidate() As Boolean
 CheckValidate = True
 If CombAccountName.ListIndex = -1 Then
    MsgBox "·ÿ›« ‰«„ Õ”«» —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    CombAccountName.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If FarDate1.Text = Empty Then
    MsgBox "·ÿ›«  «—ÌŒ ⁄„·Ì«  —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    FarDate1.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If Trim(TxtDescription) = Empty Then
    MsgBox "·ÿ›« ‘—Õ ⁄„·Ì«  —« „‘Œ’ ò‰Ìœ", vbExclamation, ""
    TxtDescription.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If Text2Currency(TxtBestankar) + Text2Currency(TxtBedehkar) = 0 Then
    MsgBox "ÌòÌ «“ ò«œ— Â«Ì »” «‰ò«— Ì« »œÂò«— »«Ìœ ò«„· ‘Êœ", vbExclamation, ""
    TxtBestankar.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If Text2Currency(TxtBestankar) <> 0 And Text2Currency(TxtBedehkar) <> 0 Then
    MsgBox "·ÿ›« ›ﬁÿ ÌòÌ «“ ò«œ— Â«—« ò«„· ‰„«ÌÌœ", vbExclamation, ""
    TxtBestankar.SetFocus
    CheckValidate = False
    Exit Function
 End If
 
End Function

Private Sub TxtBedehkar_GotFocus()
  TxtBedehkar = Format(TxtBedehkar)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtBedehkar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
 End If
 
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtBedehkar_LostFocus()
 TxtBedehkar = Format(TxtBedehkar, "#,#—Ì«·")
End Sub

Private Sub TxtBestankar_GotFocus()
  TxtBestankar = Format(TxtBestankar)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtBestankar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
 End If
 
 Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub TxtBestankar_LostFocus()
 TxtBestankar = Format(TxtBestankar, "#,#—Ì«·")
End Sub

Private Sub TxtDescription_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtDescription_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If

End Sub
