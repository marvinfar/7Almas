VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmSabtHazine 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9735
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
   ScaleHeight     =   5085
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtMablagh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox TxtDescription 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2400
      Width           =   7335
   End
   Begin VB.ComboBox CombPlace 
      Height          =   525
      ItemData        =   "FrmSabtHazine.frx":0000
      Left            =   5040
      List            =   "FrmSabtHazine.frx":0002
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4320
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
      MICON           =   "FrmSabtHazine.frx":0004
      PICN            =   "FrmSabtHazine.frx":0020
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
      TabIndex        =   2
      Top             =   4320
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
      MICON           =   "FrmSabtHazine.frx":3A8E
      PICN            =   "FrmSabtHazine.frx":3AAA
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
      Left            =   5040
      TabIndex        =   5
      Top             =   1575
      Width           =   2775
      _ExtentX        =   4895
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
   Begin HaftAlmas.TypeButton CmdReport 
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   4320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ê‹“«—‘"
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
      MICON           =   "FrmSabtHazine.frx":707E
      PICN            =   "FrmSabtHazine.frx":709A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â À»  Â‹“Ì‹‰‹Â Â«Ì »‹«—‘„‹«—Ì ò‹«”‹Å‹Ì‹‰ Œ‹“—"
      Height          =   405
      Left            =   285
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   9300
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„»‹·€ Â‹“Ì‹‰Â"
      Height          =   405
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘‹—Õ Â‹“Ì‹‰Â"
      Height          =   405
      Left            =   8145
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰‹«„ ‘‹—ò‹ "
      Height          =   405
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ À»  Â‹“Ì‰Â"
      Height          =   405
      Left            =   7890
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1620
      Width           =   1575
   End
   Begin VB.Image ImgBackground 
      Height          =   5100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "FrmSabtHazine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdReport_Click()
 With FrmGetReportDarAmad
    .WichReport = CombPlace.ListIndex + 1
    .LblTitle = "‘—Õ Â‹“Ì‰Â"
    .TxtKeshti.ToolTipText = "„Ì  Ê«‰Ìœ ﬁ”„ Ì «“ ‘—Õ Â“Ì‰Â —« »—«Ì Ã” ÃÊ  «ÌÅ ‰„«ÌÌœ"
    .CmdOKDarAmad.Visible = False
    .CmdOkHazine.Visible = True
    .Show 1
 End With
End Sub

Private Sub CmdSave_Click()
 If Not CheckValidate Then Exit Sub
 '
 Dim codePlace As Byte
 Dim LastCount0 As Long
 Dim rs As New Recordset
 Dim StrSql As String
 Dim CodeBank As Byte
 
 codePlace = CombPlace.ListIndex + 1 ' Khazar=1 GOL=2
 ''' Sabte Hazine GOL ba tavajo be hesabe Tankhah
 If codePlace = 2 Then
    StrSql = "SELECT CodeBank FROM DefBank "
    StrSql = StrSql & "WHERE AccountName LIKE '%" & " ‰ŒÊ«Â" & "%'"
    rs.Open StrSql, CNS
    If Not rs.EOF Then
       CodeBank = rs(0)
       rs.Close
    Else
       MsgBox "·ÿ›« »—«Ì ”Ì” „ Ìò Õ”«»  ‰ŒÊ«Â «ÌÃ«œ ‰„«ÌÌœ", vbExclamation, ""
       rs.Close
       Exit Sub
    End If
    '
    StrSql = "SELECT MAX(Count0) FROM TransactionBank "
    StrSql = StrSql & "WHERE CodeBank=" & CodeBank
    rs.Open StrSql, CNS
    LastCount0 = IIf(IsNull(rs(0)), 1, rs(0) + 1)
    rs.Close
 
    StrSql = "INSERT INTO TransactionBank "
    StrSql = StrSql & "VALUES(" & CodeBank & "," & LastCount0 & ","
    StrSql = StrSql & "'" & Mid(FarDate1.Text, 3) & "',"
    StrSql = StrSql & "'" & Trim(TxtDescription) & "',"
    StrSql = StrSql & "0," & Text2Currency(TxtMablagh) & ")"
    rs.Open StrSql, CNS
 End If
 '''
 StrSql = "SELECT MAX(Count0) FROM Hazine "
 StrSql = StrSql & "WHERE CodePlace=" & codePlace
 rs.Open StrSql, CNS
 LastCount0 = IIf(IsNull(rs(0)), 1, rs(0) + 1)
 rs.Close
 '
 StrSql = "INSERT INTO Hazine "
 StrSql = StrSql & "VALUES(" & codePlace & "," & LastCount0
 StrSql = StrSql & ",'" & Trim(TxtDescription) & "',"
 StrSql = StrSql & "'" & Mid(FarDate1.Text, 3) & "',"
 StrSql = StrSql & Text2Currency(TxtMablagh) & ") "
 
 rs.Open StrSql, CNS
 '
 MsgBox "«ÿ·«⁄«  À»  ‘œ", vbInformation, ""
 Set rs = Nothing
 '
 TxtDescription = Empty
 TxtMablagh = Empty
 TxtDescription.SetFocus
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
 '
 CombPlace.AddItem "»‹«— ‘„«—Ì ò«”ÅÌ‰ Œ‹“—"
 CombPlace.AddItem "ò‘ ‹Ì—«‰Ì ê‹·"
 CombPlace.Enabled = False
 '
 
End Sub

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Function CheckValidate() As Boolean
 CheckValidate = True
 If CombPlace.ListIndex = -1 Then
    MsgBox "·ÿ›« ‰«„ ‘‹—ò‹  —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    CombPlace.SetFocus
    CheckValidate = False
    Exit Function
 End If
 '
 If Trim(TxtDescription) = Empty Then
    MsgBox "·ÿ›« ‘—Õ Â“Ì‰Â —«  ‹«ÌÅ ò‰Ìœ", vbExclamation, ""
    TxtDescription.SetFocus
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
 If Text2Currency(TxtMablagh) = 0 Then
    MsgBox "·ÿ›« „»·‹€ Â“Ì‰Â —« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    TxtMablagh.SetFocus
    CheckValidate = False
    Exit Function
 End If
 
End Function

Private Sub TxtDescription_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
End Sub

Private Sub TxtMablagh_GotFocus()
  TxtMablagh = Format(TxtMablagh)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtMablagh_KeyPress(KeyAscii As Integer)
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

Private Sub TxtMablagh_LostFocus()
 TxtMablagh = Format(TxtMablagh, "#,#—Ì«·")
End Sub
