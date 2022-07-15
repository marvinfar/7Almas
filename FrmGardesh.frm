VERSION 5.00
Begin VB.Form FrmGardesh 
   BorderStyle     =   0  'None
   Caption         =   "›—„  ⁄—Ì› Å—Ê«‰Â Ê Å«—  ÃœÌœ"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
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
   ScaleHeight     =   5790
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtNoeKala 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   3375
   End
   Begin VB.ComboBox CombMoshtari 
      Height          =   525
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "CombMoshtari"
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox CombMoshtariCode 
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtMolahezat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4320
      Width           =   6015
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
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
      MICON           =   "FrmGardesh.frx":0000
      PICN            =   "FrmGardesh.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtTonaj 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox TxtBandel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox TxtParvane 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   2175
   End
   Begin HaftAlmas.TypeButton CmdSave 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   5040
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
      MICON           =   "FrmGardesh.frx":3A8A
      PICN            =   "FrmGardesh.frx":3AA6
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
      Caption         =   "‰Ê⁄ ò«·«"
      Height          =   405
      Left            =   7395
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’«Õ» ò«·«"
      Height          =   405
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ« "
      Height          =   405
      Left            =   7230
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4320
      Width           =   825
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê—Êœ «ÿ·«⁄«  ê—œ‘ ò«— œ— «‰»«—"
      Height          =   405
      Left            =   285
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   7740
   End
   Begin VB.Label LblAlarm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Õ·Ì »—«Ì ‰„«Ì‘ ÅÌ€«„ Œÿ«"
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   3705
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImgAlarm 
      Height          =   1680
      Left            =   240
      Picture         =   "FrmGardesh.frx":707A
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ‰«é Ê«—œÂ(òÌ·Êê—„)"
      Height          =   405
      Left            =   6285
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3600
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ »‰œ· Ê«—œÂ"
      Height          =   405
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "òœ «‰»«—"
      Height          =   405
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   675
   End
   Begin VB.Image ImgBackground 
      Height          =   5775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "FrmGardesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IF_ID_CHANGE  As Boolean
Dim mbIgnoreListClick  As Boolean

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()
 If Not CheckValidate Then Exit Sub
 '
 Dim rs As New Recordset
 Dim strSql As String
 Dim Code As Long
 'Check For No duplication in both of Parvane
 strSql = "SELECT Parvane FROM Gardesh "
 strSql = strSql & "WHERE Parvane='" & TxtParvane & "' "
 rs.Open strSql, CNS
 If Not rs.EOF Then 'if dup = yes
    MsgBox "òœ «‰»«—   ò—«—Ì «” ", vbCritical, ""
    TxtParvane.SetFocus
    rs.Close
 Else
     rs.Close
     '
     strSql = "INSERT INTO Gardesh "
     strSql = strSql & "VALUES(" & CombMoshtariCode & ",'" & TxtParvane & "',"
     strSql = strSql & "'" & TxtNoeKala & "'," & Val(TxtBandel) & ","
     strSql = strSql & Val(TxtTonaj) & ",'" & TxtMolahezat & "')"
     rs.Open strSql, CNS
     '
     MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  –ŒÌ—Â ‘œ", vbInformation, ""
     ClearText Me
     CombMoshtari.SetFocus
  End If
 
 Set rs = Nothing
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
 ImgAlarm.Visible = False
 LblAlarm.Visible = False
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackForms7.jpg")
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

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub TxtBandel_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 
  Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtMolahezat_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtNoeKala_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If

End Sub

Private Sub TxtParvane_Change()
 IF_ID_CHANGE = True
 ImgAlarm.Visible = False
 LblAlarm.Visible = False
End Sub

Private Sub TxtParvane_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
' Dim strValid As String
'   strValid = "0123456789/" + Chr(vbKeyBack) + Chr(vbKeyDelete)
'   If InStr(strValid, Chr(KeyAscii)) = 0 Then
'      KeyAscii = 0
'   End If

End Sub




Private Sub TxtParvane_LostFocus()
 If IF_ID_CHANGE Then  ' check code is duplicate or no
    Dim rs As New Recordset
    Dim Etebar As String
    
    rs.Open "SELECT Parvane FROM Gardesh " & _
            "WHERE Parvane ='" & TxtParvane & "'", CNS
    
    If Not rs.EOF Then   ' if Parvane is duplicate
       ImgAlarm.Visible = True
       LblAlarm.Visible = True
       LblAlarm = "«Ì‰ Å—Ê«‰Â ﬁ»·« œ— ”Ì” „ Ê«—œ ‘œÂ «” "
    End If
    rs.Close
    Set rs = Nothing
 End If

End Sub

Private Function CheckValidate() As Boolean
 CheckValidate = True
 
 If TxtNoeKala = Empty Then
    CheckValidate = False
    MsgBox "‰Ê⁄ ò«·« Œ«·Ì „Ì »«‘œ", vbExclamation, ""
    TxtNoeKala.SetFocus
    Exit Function
 End If
 
 If TxtParvane = Empty Then
    CheckValidate = False
    MsgBox "‘„«—Â òœ«‰»«— Œ«·Ì „Ì »«‘œ", vbExclamation, ""
    TxtParvane.SetFocus
    Exit Function
 End If
 '
 If CombMoshtari.ListIndex = -1 Then
    CheckValidate = False
    MsgBox "‰«„ ’«Õ» ò«·«—« «‰ Œ«» ‰„«ÌÌœ", vbExclamation, ""
    CombMoshtari.SetFocus
    Exit Function
 End If
 '
 If Val(TxtBandel) <= 0 Then
    CheckValidate = False
    MsgBox "»‰œ· Ê«—œ ‘œÂ ’ÕÌÕ ‰„Ì »«‘œ", vbExclamation, ""
    TxtBandel.SetFocus
    Exit Function
 End If
 '
 If Val(TxtTonaj) <= 0 Then
    CheckValidate = False
    MsgBox " ‰«é Ê«—œ ‘œÂ ’ÕÌÕ ‰„Ì »«‘œ", vbExclamation, ""
    TxtTonaj.SetFocus
    Exit Function
 End If
 
End Function


Private Sub TxtTonaj_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 
  Dim strValid As String
   strValid = "0123456789" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub
