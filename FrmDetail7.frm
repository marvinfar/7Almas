VERSION 5.00
Object = "{9DBDC544-49CA-11D7-B1ED-C2237039C523}#1.1#0"; "FarDate.Ocx"
Begin VB.Form FrmDetail7 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7350
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtKeraye 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6480
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox TxtDriverName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6240
      MaxLength       =   11
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3480
      Width           =   2295
   End
   Begin VB.ComboBox CombMoshtariCode 
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CombMoshtari 
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   "CombMoshtari"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox TxtParvande 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox TxtMobile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox TxtPayeKeraye 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox TxtSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox TxtTedad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox TxtVazn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox TxtKamioon 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox TxtAddress 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2640
      Width           =   4815
   End
   Begin FarDate1.FarDate FarDate1 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
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
   Begin VB.TextBox TxtBarNamE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox TxtPart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   7200
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox CombEtebar 
      Height          =   525
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox CombParvane 
      Height          =   525
      Left            =   3720
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "���� �����"
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
      MICON           =   "FrmDetail7.frx":0000
      PICN            =   "FrmDetail7.frx":001C
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
      TabIndex        =   17
      Top             =   6600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "��� �������"
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
      MICON           =   "FrmDetail7.frx":3A8A
      PICN            =   "FrmDetail7.frx":3AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �������"
      Height          =   405
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� ������"
      Height          =   405
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3480
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� ����"
      Height          =   765
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   5040
      Width           =   705
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblAlarm2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� ���� ����� ����� ���"
      ForeColor       =   &H00FF00FF&
      Height          =   405
      Left            =   6660
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   6600
      Width           =   2445
   End
   Begin VB.Label LblAlarm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� ���� ����� ����� ���"
      ForeColor       =   &H00FF00FF&
      Height          =   405
      Left            =   6660
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   6120
      Width           =   2445
   End
   Begin VB.Image ImgAlarm 
      Height          =   480
      Left            =   9120
      Picture         =   "FrmDetail7.frx":707A
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   480
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� �����"
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� ������"
      Height          =   405
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   5160
      Width           =   1020
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� �������"
      Height          =   405
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   405
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4320
      Width           =   420
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      Height          =   405
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� ��Ә��"
      Height          =   405
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� �����"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5145
      TabIndex        =   28
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   405
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� ���"
      Height          =   405
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� �������"
      Height          =   405
      Left            =   5805
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1800
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   405
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� ������"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   2460
      MouseIcon       =   "FrmDetail7.frx":A1EE
      MousePointer    =   99  'Custom
      RightToLeft     =   -1  'True
      TabIndex        =   23
      ToolTipText     =   "�� ��� �� ������ ����� ������ �� ����� ����"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� ������"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   5880
      MouseIcon       =   "FrmDetail7.frx":A4F8
      MousePointer    =   99  'Custom
      RightToLeft     =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "�� ��� �� ������ ����� ������ �� ����� ����"
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��� ������"
      Height          =   405
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   840
      Width           =   915
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����� ���� ������� ���� ���� ��� ��� ������"
      Height          =   405
      Left            =   990
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   20
      ToolTipText     =   "�� ����� ��� ����� ��� ��� ����� ���� �� ���"
      Top             =   120
      Width           =   8595
   End
   Begin VB.Image ImgBackground 
      Height          =   7380
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "FrmDetail7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodePaEt As Long ' codi ke dar jadvale main be ezaye Parvane va Etebar Sabt shode
Dim RadifVarede As Long  ' baraye namayeshe akharin radif
Dim VazneNakhales As Long
Dim mbIgnoreListClick  As Boolean
Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdSave_Click()
 If Not CheckValidate Then Exit Sub
 '
 Dim strSql As String
 Dim rs As New Recordset
 Dim LastCount0 As Integer
 '
 strSql = "SELECT MAX(Count0) FROM Detail7 "
 strSql = strSql & "WHERE Code=" & CodePaEt
 rs.Open strSql, CNS
 LastCount0 = IIf(IsNull(rs(0)), 1, rs(0) + 1)
 rs.Close
 ''
 strSql = "INSERT INTO Detail7 "
 strSql = strSql & "VALUES(" & CodePaEt & "," & LastCount0 & "," & Val(TxtBarNamE)
 strSql = strSql & ",'" & Mid(FarDate1.Text, 3) & "','" & Trim(TxtAddress) & "',"
 strSql = strSql & "'" & "����" & "','" & TxtKamioon & TxtSerial & "',"
 strSql = strSql & Val(TxtVazn) & "," & Val(TxtTedad) & ",'" & TxtSize & "',"
 strSql = strSql & Text2Currency(TxtKeraye) & ",'" & TxtMobile & "',"
 strSql = strSql & "'" & TxtParvande & "','" & TxtName & "',"
 strSql = strSql & Val(CombMoshtariCode) & ",'" & "����2" & "',"
 strSql = strSql & Text2Currency(TxtPayeKeraye) & ",'" & TxtDriverName & "')"
 rs.Open strSql, CNS
 ''Kasr Az Kole PART
 Call KasrAzPART
 '
 Set rs = Nothing
 ''
 Dim r As Byte, g As Byte, b As Byte
 Randomize
 r = Rnd(2.55) * 100
 g = Rnd(2.55) * 100
 b = Rnd(2.55) * 100
 TxtBarNamE = Val(TxtBarNamE) + 1
 RadifVarede = RadifVarede + 1
 LblAlarm2.Visible = True
 LblAlarm2 = "����� ���� ���� ��� ����� " & RadifVarede & " �� ����"

 TxtBarNamE.BackColor = RGB(r, g, b)
 TxtBarNamE.ForeColor = vbWhite
 
 TxtKamioon = Empty
 TxtSerial = Empty
 TxtVazn = Empty
 TxtMobile = Empty
 TxtDriverName = Empty
 ''
 TxtBarNamE.SetFocus
 SendKeys "{home}+{end}"
End Sub

Private Sub CmdSave_GotFocus()
 CmdSave.ForeColor = vbRed
End Sub

Private Sub CmdSave_LostFocus()
 CmdSave.BackColor = &H8000000F
End Sub

Private Sub CombEtebar_Click()
 Dim rs As New Recordset
 Dim strSql As String
 'Find Code of Selected Parvane and Etebar
 strSql = "SELECT Code,Part,Weight,MoshtariCode FROM Main7 "
 strSql = strSql & "WHERE Parvane='" & CombParvane & "' "
 strSql = strSql & "AND Etebar='" & CombEtebar & "'"
 rs.Open strSql, CNS
 If Not rs.EOF Then
    CodePaEt = rs(0)
    TxtPart = rs(1)
    VazneNakhales = rs(2)
    If rs(3) <> 0 Then
       Dim i As Integer
       For i = 0 To CombMoshtariCode.ListCount - 1
           If Val(CombMoshtariCode.List(i)) = rs(3) Then
              CombMoshtariCode.ListIndex = i
              CombMoshtari.ListIndex = i
              Exit For
           End If
       Next
    End If
 End If
 rs.Close
 'find Last Barname
 Dim SubQuery As String
 SubQuery = "SELECT MAX(Count0) FROM Detail7 "
 SubQuery = SubQuery & "WHERE Code=" & CodePaEt
  ' ' '
 strSql = "SELECT BarName,Count0 FROM Detail7 "
 strSql = strSql & "WHERE Code=" & CodePaEt
 strSql = strSql & " AND Count0=(" & SubQuery & ")"
 
 rs.Open strSql, CNS
 If Not rs.EOF Then
    LblAlarm = "����� ������� ���� ��� �� ����� " & rs(0) & " �� ����"
    LblAlarm2 = "����� ���� ���� ��� ����� " & rs(1) & " �� ����"
    RadifVarede = 0
    RadifVarede = rs(1)
    LblAlarm.Visible = True
    ImgAlarm.Visible = True
    LblAlarm2.Visible = True
 Else
    LblAlarm_Click
 End If
 rs.Close
''
 Set rs = Nothing
 TxtBarNamE.SetFocus
End Sub

Private Sub CombEtebar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
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

Private Sub CombParvane_Click()
 CombEtebar.Clear
 CombMoshtari.ListIndex = -1
 CombMoshtariCode.ListIndex = -1
 '
 If CombParvane.ListIndex = -1 Then Exit Sub
 
 Dim rs As New Recordset
 Dim strSql As String
 '
 strSql = "SELECT Etebar FROM Main7 "
 strSql = strSql & "WHERE Parvane='" & CombParvane & "'"
 rs.Open strSql, CNS
 If rs.EOF Then
    MsgBox "����� ������ ���� ��� ������ ��� ���� ���" & vbNewLine & _
           "���� �� ��� ����� ���� ���� �ѐ�� ����", vbExclamation, ""
 Else
    Do While Not rs.EOF
       CombEtebar.AddItem rs(0)
       rs.MoveNext
    Loop
    CombEtebar.Enabled = True
 End If
 rs.Close
 Set rs = Nothing
 
End Sub

Private Sub CombParvane_KeyPress(KeyAscii As Integer)
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
 '
End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 '
 LblAlarm.Visible = False
 ImgAlarm.Visible = False
 LblAlarm2.Visible = False
 '
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackForms7.jpg")
 '
 TxtName = "��������"
 FarDate1.Text = FarDate1.today
 '
 Call LoadParvane
 CombEtebar.Enabled = False
 '
 TxtPart = Empty
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

Private Sub ImgAlarm_Click()
 LblAlarm.Visible = False
 ImgAlarm.Visible = False
 LblAlarm2.Visible = False
End Sub

Private Sub Label1_Click()
  FindCombo CombParvane
End Sub

Private Sub Label2_Click()
  FindCombo CombEtebar
End Sub

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub LblAlarm_Click()
 LblAlarm.Visible = False
 ImgAlarm.Visible = False
 LblAlarm2.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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

Private Sub TxtAddress_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If
  
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtBarNamE_Change()
 LblAlarm_Click
End Sub

Private Sub TxtBarNamE_KeyPress(KeyAscii As Integer)
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

Private Sub TxtBarNamE_LostFocus()
 Dim strSql As String
 Dim rs As New Recordset
 '
 strSql = "SELECT BarName FROM Detail7 "
 strSql = strSql & "WHERE BarName=" & Val(TxtBarNamE)
 strSql = strSql & " AND Code=" & CodePaEt
 rs.Open strSql, CNS
 If Not rs.EOF Then
    ImgAlarm.Visible = True
    LblAlarm.Visible = True
    LblAlarm = "��� ����� ������� ���� ���� ��� ���"
    MsgBox "��� ����� ������� ���� ���� ��� ���", vbExclamation, ""
 Else
    LblAlarm_Click
 End If
 rs.Close
 Set rs = Nothing
End Sub

Private Sub TxtHavale_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtHavale_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
' Dim strValid As String
'   strValid = "0123456789/-" + Chr(vbKeyBack) + Chr(vbKeyDelete)
'   If InStr(strValid, Chr(KeyAscii)) = 0 Then
'      KeyAscii = 0
'   End If
End Sub

Private Sub TxtDriverName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtKamioon_Change()
  If Len(TxtKamioon) = 3 Then
     TxtKamioon = TxtKamioon & "�"
     SendKeys "{End}"
  End If
  '
  If Len(TxtKamioon) = 6 Then TxtSerial.SetFocus

End Sub

Private Sub TxtKamioon_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtKeraye_GotFocus()
  TxtKeraye = Format(TxtKeraye)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtKeraye_KeyPress(KeyAscii As Integer)
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

Private Sub TxtKeraye_LostFocus()
 TxtKeraye = Format(TxtKeraye, "#,#����")
End Sub

Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
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

Private Sub TxtMomayez2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtName_GotFocus()
 Dim oldKB As Long

  oldKB = GetKeyboardLayout(0)
  'Change keyboard farsi
  If oldKB = 67699721 Then 'keyboard is English
     ActivateKeyboardLayout HKL_NEXT, ByVal 0&
  End If

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

Private Sub TxtParvande_GotFocus()
 TxtParvande.BackColor = vbRed
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtParvande_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
End Sub

Private Sub TxtParvande_LostFocus()
 TxtParvande.BackColor = &HE0E0E0
End Sub

Private Sub TxtPayeKeraye_GotFocus()
  TxtKeraye = Format(TxtKeraye)
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtPayeKeraye_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtSerial_GotFocus()
  TxtSerial = "�����"
  SendKeys "{End}"
End Sub

Private Sub TxtSerial_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub TxtSize_GotFocus()
  SendKeys "{home}+{end}"

End Sub

Private Sub TxtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If
 '
 Dim strValid As String
   strValid = "0123456789*" + Chr(vbKeyBack) + Chr(vbKeyDelete)
   If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub

Private Sub TxtTedad_GotFocus()
 SendKeys "{Home}+{end}"
End Sub

Private Sub TxtTedad_KeyPress(KeyAscii As Integer)
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

Private Sub TxtVazn_KeyPress(KeyAscii As Integer)
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

Private Sub LoadParvane()
 Dim rs As New Recordset
 Dim strSql As String
 Dim Max As Long
 '
 ''faghat 6 ta parvaneye Akhar Load shavad
 strSql = "SELECT MAX(Code) FROM Main7 "
 rs.Open strSql, CNS
 Max = IIf(IsNull(rs(0)), 0, rs(0))
 rs.Close
 
 strSql = "SELECT DISTINCT Parvane FROM Main7 "
 strSql = strSql & "WHERE Code BETWEEN " & Max - 5 & " AND " & Max
 rs.Open strSql, CNS

 Do While Not rs.EOF
    CombParvane.AddItem rs(0)
    rs.MoveNext
 Loop
 
 rs.Close
 Set rs = Nothing
End Sub

Private Function CheckValidate() As Boolean
 CheckValidate = True
 If CombParvane.ListIndex = -1 Then
    MsgBox "����� ������ ���� ���", vbExclamation, ""
    CheckValidate = False
    CombParvane.SetFocus
    Exit Function
 End If
 '
 If CombEtebar.ListIndex = -1 Then
    MsgBox "����� ������ ���� ���", vbExclamation, ""
    CheckValidate = False
    If CombEtebar.Enabled Then CombEtebar.SetFocus
    Exit Function
 End If
 '
 If Val(TxtPart) <= 0 Then
    MsgBox "���� ���� ��� ������ ���", vbExclamation, ""
    CheckValidate = False
    TxtPart.SetFocus
    Exit Function
 End If
 '
 If Val(TxtBarNamE) <= 0 Then
    MsgBox "����� ������� �� ���� ������", vbExclamation, ""
    CheckValidate = False
    TxtBarNamE.SetFocus
    Exit Function
 End If
 '
 If FarDate1.Text = Empty Then
    MsgBox "����� ��� ������ �� ����", vbExclamation, ""
    CheckValidate = False
    FarDate1.SetFocus
    Exit Function
 End If
 '
 If Val(TxtVazn) = 0 Then
    MsgBox "��� ������� �� ���� ����", vbExclamation, ""
    CheckValidate = False
    TxtVazn.SetFocus
    Exit Function
 End If

End Function

Private Sub FindCombo(C As ComboBox)
 Dim inp As String
 Dim i As Long
 Dim b As Boolean
 '
 inp = InputBox("����� �� ���� ������", "�����")
 b = False
 For i = 0 To C.ListCount - 1
     If C.List(i) = inp Then
        b = True
        Exit For
     End If
 Next
 If b Then
    C.ListIndex = i
 Else ' Agar dar List peyda nashod dar bank migardad
    Dim strSql As String
    Dim rs As New Recordset
    '
    strSql = "SELECT Parvane FROM Main7 "
    strSql = strSql & "WHERE Parvane='" & inp & "'"
    rs.Open strSql, CNS
    If Not rs.EOF Then
       C.AddItem rs(0), 0
       C.ListIndex = 0
    Else
       C.ListIndex = -1
       If LCase(C.Name) = "combparvane" Then CombEtebar.Enabled = False
    End If
    
    rs.Close
    Set rs = Nothing
 End If

End Sub

Private Sub KasrAzPART()
 Dim strSql As String
 Dim rs As New Recordset
 Dim LastCount As Byte
 Dim BaghimandeyeRoozeGhabl As Integer
 Dim BaghiVazneGhabl As Long
 
 strSql = "SELECT * FROM Baghimande7 "
 strSql = strSql & "WHERE Code=" & CodePaEt
 strSql = strSql & " AND Tarikh='" & Mid(FarDate1.Text, 3) & "'"
 rs.Open strSql, CNS
 If rs.EOF Then ' yani avalin bar ast
    strSql = "SELECT MAX(Count0) FROM Baghimande7 "
    strSql = strSql & "WHERE Code=" & CodePaEt
    rs.Close
    rs.Open strSql, CNS
    LastCount = IIf(IsNull(rs(0)), 1, rs(0) + 1)
    rs.Close
    '''
    strSql = "INSERT INTO Baghimande7 "
    strSql = strSql & "VALUES(" & CodePaEt & "," & LastCount & ",'"
    strSql = strSql & Mid(FarDate1.Text, 3) & "',"
    If LastCount = 1 Then ' avalin rooz baraye sabte parvane
       strSql = strSql & Val(TxtPart) - Val(TxtTedad) & ","
       strSql = strSql & VazneNakhales - Val(TxtVazn) & ")"
    Else
       Dim tempStr As String ' jostejooye baghimandeye rooze ghabl
       tempStr = "SELECT Baghimande,BaghiVazn FROM Baghimande7 "
       tempStr = tempStr & "WHERE Code=" & CodePaEt
       tempStr = tempStr & " AND Count0=" & LastCount - 1
       rs.Open tempStr, CNS
       BaghimandeyeRoozeGhabl = rs(0)
       BaghiVazneGhabl = rs(1)
       rs.Close
       '
       strSql = strSql & BaghimandeyeRoozeGhabl - Val(TxtTedad) & ","
       strSql = strSql & BaghiVazneGhabl - Val(TxtVazn) & ")"
    End If
    
    rs.Open strSql, CNS
 Else
    strSql = "UPDATE Baghimande7 SET "
    strSql = strSql & "Baghimande = Baghimande - " & Val(TxtTedad)
    strSql = strSql & ",BaghiVazn= BaghiVazn - " & Val(TxtVazn)
    strSql = strSql & " WHERE Code=" & CodePaEt
    strSql = strSql & " AND Tarikh='" & Mid(FarDate1.Text, 3) & "'"
    rs.Close
    rs.Open strSql, CNS
 End If
 Set rs = Nothing
 
End Sub

