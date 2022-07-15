VERSION 5.00
Begin VB.Form FrmMain7 
   BorderStyle     =   0  'None
   Caption         =   "›—„  ⁄—Ì› Å—Ê«‰Â Ê Å«—  ÃœÌœ"
   ClientHeight    =   5190
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
   ScaleHeight     =   5190
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtKeshtiName 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2760
      Width           =   3255
   End
   Begin VB.ComboBox CombMoshtari 
      Height          =   525
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "CombMoshtari"
      Top             =   3480
      Width           =   2895
   End
   Begin VB.ComboBox CombMoshtariCode 
      Height          =   525
      Left            =   3240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtWeight 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2760
      Width           =   2175
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   7
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
      MICON           =   "FrmMain7.frx":0000
      PICN            =   "FrmMain7.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtPart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox TxtEtebar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox TxtParvane 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin HaftAlmas.TypeButton CmdSave 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
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
      MICON           =   "FrmMain7.frx":3A8A
      PICN            =   "FrmMain7.frx":3AA6
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
      Caption         =   "‰«„ ò‘ Ì"
      Height          =   405
      Left            =   3615
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’«Õ» ò«·«"
      Height          =   405
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê“‰ ‰«Œ«·’"
      Height          =   405
      Left            =   7005
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â „⁄—›Ì Å—Ê«‰Â Ê Å«—  ÃœÌœ («» œ«Ì ò«—- ﬁ»· «“ Ê—Êœ «ÿ·«⁄«  —Ê“«‰Â)"
      Height          =   405
      Left            =   1440
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   6585
   End
   Begin VB.Label LblAlarm 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Õ·Ì »—«Ì ‰„«Ì‘ ÅÌ€«„ Œÿ«"
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   4185
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImgAlarm 
      Height          =   1680
      Left            =   240
      Picture         =   "FrmMain7.frx":707A
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‹‹«— "
      Height          =   405
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â «⁄ »«—"
      Height          =   405
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Å—Ê«‰Â"
      Height          =   405
      Left            =   6900
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   1155
   End
   Begin VB.Image ImgBackground 
      Height          =   5175
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "FrmMain7"
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
 'Check For No duplication in both of Parvane AND Etebar
 strSql = "SELECT Parvane,Etebar FROM Main7 "
 strSql = strSql & "WHERE Parvane='" & TxtParvane & "' "
 strSql = strSql & "AND Etebar='" & TxtEtebar & "'"
 rs.Open strSql, CNS
 If Not rs.EOF Then 'if dup = yes
    MsgBox "‘„«—Â Å—Ê«‰Â Ê «⁄ »«— Â—œÊ  ò—«—Ì Â” ‰œ", vbCritical, ""
    TxtParvane.SetFocus
    rs.Close
 Else
     rs.Close
     '
     Code = MakeAutoNumber("Main7", "Code")
    
     strSql = "INSERT INTO Main7 "
     strSql = strSql & "VALUES(" & Code & ",'" & TxtParvane & "',"
     strSql = strSql & "'" & TxtEtebar & "','" & TxtPart & "',"
     strSql = strSql & Val(TxtWeight) & "," & Val(CombMoshtariCode)
     strSql = strSql & ",'" & Trim(TxtKeshtiName) & "')"
     
     rs.Open strSql, CNS
     '
     MsgBox "«ÿ·«⁄«  »« „Ê›ﬁÌ  –ŒÌ—Â ‘œ", vbInformation, ""
     ClearText Me
     TxtParvane.SetFocus
  End If
 
 Set rs = Nothing
End Sub

Private Sub Command1_Click()
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
 rs.Open "SELECT * FROM Moshtari ORDER BY MoshtariName ", CNS
 Do While Not rs.EOF
    CombMoshtari.AddItem rs("MoshtariName")
    CombMoshtariCode.AddItem rs("MoshtariCODE")
    rs.MoveNext
 Loop
 rs.Close
 Set rs = Nothing
 
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub TxtKeshtiName_KeyPress(KeyAscii As Integer)
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

Private Sub Text2_Change()

End Sub

Private Sub TxtEtebar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789/" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If

End Sub

Private Sub Text3_Change()

End Sub

Private Sub TxtPart_KeyPress(KeyAscii As Integer)
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

Private Sub TxtParvane_LostFocus()
 If IF_ID_CHANGE Then  ' check code is duplicate or no
    Dim rs As New Recordset
    Dim Etebar As String
    
    rs.Open "SELECT Parvane,Etebar FROM Main7 " & _
            "WHERE Parvane ='" & TxtParvane & "'", CNS
    
    If Not rs.EOF Then   ' if Parvane is duplicate
       Etebar = IIf(IsNull(rs(1)), "", rs(1))
       ImgAlarm.Visible = True
       LblAlarm.Visible = True
       LblAlarm = "«Ì‰ Å—Ê«‰Â ﬁ»·« »« ‘„«—Â «⁄ »«— " & Etebar & " œ— ”Ì” „ Ê«—œ ‘œÂ «” "
    End If
    rs.Close
    Set rs = Nothing
 End If

End Sub

Private Function CheckValidate() As Boolean
 CheckValidate = True
 If TxtParvane = Empty Then
    CheckValidate = False
    MsgBox "‘„«—Â Å—Ê«‰Â Œ«·Ì „Ì »«‘œ", vbExclamation, ""
    TxtParvane.SetFocus
    Exit Function
 End If
 '
 If TxtEtebar = Empty Then
    CheckValidate = False
    MsgBox "‘„«—Â «⁄ »«— Œ«·Ì „Ì »«‘œ", vbExclamation, ""
    TxtEtebar.SetFocus
    Exit Function
 End If
 '
 If Val(TxtPart) <= 0 Then
    CheckValidate = False
    MsgBox "Å«—  Ê«—œ ‘œÂ ’ÕÌÕ ‰„Ì »«‘œ", vbExclamation, ""
    TxtPart.SetFocus
    Exit Function
 End If
 '
 If Val(TxtWeight) <= 0 Then
    CheckValidate = False
    MsgBox "Ê“‰ Ê«—œ ‘œÂ ’ÕÌÕ ‰„Ì »«‘œ", vbExclamation, ""
    TxtWeight.SetFocus
    Exit Function
 End If
 
End Function

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)
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
