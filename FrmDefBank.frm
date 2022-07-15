VERSION 5.00
Begin VB.Form FrmDefBank 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5310
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPool 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox TxtCardNumber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox TxtAccountNumber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox TxtAccountName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2895
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   6
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
      MICON           =   "FrmDefBank.frx":0000
      PICN            =   "FrmDefBank.frx":001C
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
      Left            =   2280
      TabIndex        =   4
      Top             =   4560
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
      MICON           =   "FrmDefBank.frx":3A8A
      PICN            =   "FrmDefBank.frx":3AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ﬁœ«— ÅÊ· »—«Ì «›  «Õ Õ”«»"
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   420
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1680
      Width           =   2505
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„ÊÃÊœÌ ÅÊ· «Ê·ÌÂ"
      Height          =   405
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1560
   End
   Begin VB.Image ImgAlarm 
      Height          =   480
      Left            =   7140
      Picture         =   "FrmDefBank.frx":707A
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label LblAlarm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Õ·Ì »—«Ì ‰„«Ì‘ ÅÌ€«„ Œÿ«"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4080
      Width           =   2445
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ Ì«—Ì"
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ Ì«—Ì"
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   2235
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2400
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ò«— "
      Height          =   405
      Left            =   6510
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3240
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Õ”«»"
      Height          =   405
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ Õ”«» Ì« »«‰ò"
      Height          =   405
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â „⁄—›Ì Õ”«»Â«Ì »«‰òÌ ‘„« "
      Height          =   405
      Left            =   1020
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image ImgBackground 
      Height          =   5340
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "FrmDefBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()
 If Trim(TxtAccountName) = Empty Then
    MsgBox "·ÿ›« ‰«„ Õ”«» —« Ê«—œ ‰„«ÌÌœ", vbExclamation, ""
    TxtAccountName.SetFocus
    Exit Sub
 End If
 '
 Dim rs As New Recordset
 Dim strSql As String
 Dim Code As Byte
 
 Code = MakeAutoNumber("DefBank", "CodeBank")
 strSql = "INSERT INTO DefBank(CodeBank,AccountName,Pool,AccountNumber,CardNumber) "
 strSql = strSql & "VALUES(" & Code & ",'" & TxtAccountName & "',"
 strSql = strSql & Text2Currency(TxtPool) & ","
 strSql = strSql & "'" & TxtAccountNumber & "','" & TxtCardNumber & "')"
 rs.Open strSql, CNS
 '
 If Text2Currency(TxtPool) > 0 Then Call SaveInTransaction(Code)
 '
 MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  –ŒÌ—Â ‘œ", vbInformation, ""
 ClearText Me
 TxtAccountName.SetFocus
 
 Set rs = Nothing

End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackFormsBanki.jpg")
 LblAlarm_Click
End Sub

Private Sub LblTitle_DblClick()
  CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub LblAlarm_Click()
 ImgAlarm.Visible = False
 LblAlarm.Visible = False
End Sub

Private Sub TxtPool_GotFocus()
  TxtPool = Format(TxtPool)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtPool_KeyPress(KeyAscii As Integer)
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

Private Sub TxtAccountName_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

End Sub

Private Sub TxtAccountName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If

End Sub

Private Sub TxtAccountName_LostFocus()
 If TxtAccountName <> Empty Then
    Dim strSql As String
    Dim rs As New Recordset
    '
    strSql = "SELECT * FROM DefBank "
    strSql = strSql & "WHERE AccountName LIKE '%" & Trim(TxtAccountName) & "%'"
    rs.Open strSql, CNS
    If Not rs.EOF Then
       ImgAlarm.Visible = True
       LblAlarm.Visible = True
       LblAlarm = "Ìò Õ”«» »« ‰«„[ " & rs(1) & "] »Â ‘„«—Â " & rs(2) & " ÊÃÊœ œ«—œ"
    Else
       LblAlarm_Click
    End If
    rs.Close
    Set rs = Nothing
 End If
 
End Sub

Private Sub TxtAccountNumber_KeyPress(KeyAscii As Integer)
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

Private Sub TxtCardNumber_KeyPress(KeyAscii As Integer)
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

Private Sub TxtPool_LostFocus()
 TxtPool = Format(TxtPool, "#,#—Ì«·")
End Sub

Private Sub SaveInTransaction(CodeBank As Byte)
 Dim strSql As String
 Dim rs As New Recordset
 Dim LastCount As Long
 '
 strSql = "SELECT MAX(Count0) FROM TransactionBank "
 strSql = strSql & "WHERE CodeBank=" & CodeBank
 rs.Open strSql, CNS
 LastCount = IIf(IsNull(rs(0)), 1, rs(0) + 1)
 rs.Close
 '
 strSql = "INSERT INTO TransactionBank "
 strSql = strSql & "VALUES(" & CodeBank & "," & LastCount & ","
 strSql = strSql & "'" & Mid(FrmDetail7.FarDate1.Text, 3) & "',"
 strSql = strSql & "'«›  «Õ Õ”«» Â‰ê«„ „⁄—›Ì »Â ”Ì” „'" & ","
 strSql = strSql & Text2Currency(TxtPool) & ",0)"
 rs.Open strSql, CNS
 '
 Set rs = Nothing
End Sub

