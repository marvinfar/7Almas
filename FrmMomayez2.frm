VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form FrmMomayez2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
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
   ScaleHeight     =   6825
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CombMoshtari 
      Height          =   525
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "CombMoshtari"
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox CombMoshtariCode 
      Height          =   525
      Left            =   1320
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtTedad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   2295
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3735
      Left            =   2400
      TabIndex        =   12
      Top             =   2280
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.TextBox TxtWeight 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox TxtHavale 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   2295
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   6120
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
      MICON           =   "FrmMomayez2.frx":0000
      PICN            =   "FrmMomayez2.frx":001C
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
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
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
      MICON           =   "FrmMomayez2.frx":3A8A
      PICN            =   "FrmMomayez2.frx":3AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdDelete 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Õ–› «ÿ·«⁄« "
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
      MICON           =   "FrmMomayez2.frx":707A
      PICN            =   "FrmMomayez2.frx":7096
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdEdit 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ÊÌ—«Ì‘"
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
      MICON           =   "FrmMomayez2.frx":AC96
      PICN            =   "FrmMomayez2.frx":ACB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdReportAddress 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Ìò ”ÿ— —« «‰ Œ«» ò‰Ìœ"
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ê“«—‘ ¬œ—”"
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
      MICON           =   "FrmMomayez2.frx":E3F6
      PICN            =   "FrmMomayez2.frx":E412
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdReportSelection 
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      ToolTipText     =   "”ÿ—Â«Ì „Ê—œ ‰Ÿ— —«  Ìò œ«— ò‰Ìœ"
      Top             =   6120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ê“«—‘  Ã„⁄Ì «“ ÕÊ«·Â Â«Ì «‰ Œ«»Ì"
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
      MICON           =   "FrmMomayez2.frx":11D07
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdTikAll 
      Height          =   375
      Left            =   9660
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "«‰ Œ«»  „«„ ÕÊ«·Â Â«"
      Top             =   2640
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      MICON           =   "FrmMomayez2.frx":11D23
      PICN            =   "FrmMomayez2.frx":11D3F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdunTikALL 
      Height          =   375
      Left            =   9660
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Œ«—Ã ò—œ‰  „«„ ÕÊ«·Â Â« «“ «‰ Œ«»"
      Top             =   3120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      MICON           =   "FrmMomayez2.frx":156CE
      PICN            =   "FrmMomayez2.frx":156EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdFind 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Ã” ÃÊÌ ÕÊ«·Â"
      Top             =   4440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "Ã” ÃÊ "
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
      MICON           =   "FrmMomayez2.frx":18DFC
      PICN            =   "FrmMomayez2.frx":18E18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdTafsili 
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      ToolTipText     =   "Ìò ”ÿ— —« «‰ Œ«» ò‰Ìœ"
      Top             =   6120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ê“«—‘  ›’Ì·Ì"
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
      MICON           =   "FrmMomayez2.frx":1C717
      PICN            =   "FrmMomayez2.frx":1C733
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œ«œ"
      Height          =   405
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "’«Õ» ò«·«"
      Height          =   405
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„„Ì“ œÊ„ «‰»«—"
      Height          =   405
      Left            =   8460
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label LblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Å‰Ã—Â „⁄—›Ì Ê“‰ »—«Ì  „„Ì“ œÊ„ «‰»«—                                                                              "
      Height          =   405
      Left            =   3195
      MousePointer    =   15  'Size All
      RightToLeft     =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "»« œÊ»«— ò·Ìò »——ÊÌ «Ì‰ »Œ‘ Å‰Ã—Â »” Â „Ì ‘Êœ"
      Top             =   0
      Width           =   6870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ê“‰ „„Ì“ œÊ„ «‰»«—"
      Height          =   405
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Image ImgBackground 
      Height          =   6855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "FrmMomayez2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldHavale As String, NewHavale As String 'For Edit
Dim mbIgnoreListClick  As Boolean

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdDelete_Click()
 Dim msg As Integer
 Dim L As Long
 
 L = Grid1.ActiveCell.Row
 If L > 0 Then
    msg = MsgBox("¬Ì« „ÿ„∆‰ »Â Õ–› Â” Ìœø", vbQuestion + vbYesNo, "")
    If msg = vbYes Then
       Dim strSql As String
       Dim rs As New Recordset
       
       strSql = "DELETE FROM Momayez2 "
       strSql = strSql & "WHERE Momayez2='" & Grid1.Cell(L, 5).Text & "'"
       rs.Open strSql, CNS
       Grid1.RemoveItem L
       For L = 1 To Grid1.Rows - 1
           Grid1.Cell(L, 6).Text = L
       Next
       Set rs = Nothing
    End If
 Else
    MsgBox "»—«Ì Õ–› »«Ìœ Ìò ”ÿ— —« «‰ Œ«» ò‰Ìœ", vbExclamation, ""
 End If
End Sub

Private Sub CmdEdit_Click()
 Dim L As Long
 Dim i As Integer
 
 L = Grid1.ActiveCell.Row
 If CmdEdit.Caption = "ÊÌ—«Ì‘" Then
    If L > 0 Then
       OldHavale = Grid1.Cell(L, 5).Text
       TxtHavale = OldHavale
       TxtWeight = Grid1.Cell(L, 4).Text
       For i = 0 To CombMoshtariCode.ListCount - 1
           If Grid1.Cell(L, 1).Text = CombMoshtariCode.List(i) Then
              CombMoshtari.ListIndex = i
              Exit For
           End If
       Next
       TxtTedad = Grid1.Cell(L, 2).Text
       TxtHavale.SetFocus
       SendKeys "{home}+{end}"
       '
       CmdSave.Enabled = False
       CmdDelete.Enabled = False
       CmdFind.Enabled = False
       CmdReportAddress.Enabled = False
       CmdTafsili.Enabled = False
       Grid1.Enabled = False
       '
       CmdEdit.Caption = "À»   €ÌÌ—« "
    Else
       MsgBox "»—«Ì ÊÌ—«Ì‘ ”ÿ— „Ê—œ ‰Ÿ— —« «‰ Œ«» ò‰Ìœ", vbExclamation, ""
    End If
 ElseIf CmdEdit.Caption = "À»   €ÌÌ—« " Then
    Dim rs As New Recordset
    Dim strSql As String
    NewHavale = TxtHavale
    On Error GoTo ErrLbl:
    strSql = "UPDATE Momayez2 SET "
    strSql = strSql & "Havale='" & NewHavale & "',"
    strSql = strSql & "Weight=" & Val(TxtWeight) & ","
    strSql = strSql & "Tedad=" & Val(TxtTedad) & ","
    strSql = strSql & "MoshtariCode=" & Val(CombMoshtariCode) & " "
    
    strSql = strSql & "WHERE Havale='" & OldHavale & "'"
    rs.Open strSql, CNS
ErrLbl:
    If Err.Number <> 0 Then MsgBox "„„Ì“2 «‰»«—  ò—«—Ì „Ì »«‘œ", vbCritical, ""
    Grid1.Cell(L, 1).Text = CombMoshtariCode
    Grid1.Cell(L, 2).Text = TxtTedad
    Grid1.Cell(L, 3).Text = CombMoshtari
    Grid1.Cell(L, 4).Text = TxtWeight
    Grid1.Cell(L, 5).Text = TxtHavale
    '
    Grid1.Enabled = True
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CmdFind.Enabled = True
    CmdReportAddress.Enabled = True
    CmdTafsili.Enabled = True
    
    '
    TxtHavale = ""
    TxtWeight = ""
    TxtTedad = ""
    CombMoshtari.ListIndex = -1
    '
    CmdEdit.Caption = "ÊÌ—«Ì‘"
    '
    NewHavale = ""
    OldHavale = ""
    Set rs = Nothing
 End If
End Sub

Private Sub CmdFindDetail_Click()

End Sub

Private Sub CmdFind_Click()
 If Grid1.Rows > 1 Then
    Dim inp As String
    Dim i As Integer
    Dim b As Boolean
    
    inp = InputBox("·ÿ›« „„Ì“2 «‰»«— —« Ê«—œ ‰„«ÌÌœ", "Ã” ÃÊ")
    If inp = "" Then Exit Sub
    b = False
    For i = 1 To Grid1.Rows - 1
        If Grid1.Cell(i, 5).Text = inp Then
           b = True
           Exit For
        End If
    Next
    '
    If b Then
       Grid1.Cell(i, 5).SetFocus
       Grid1.SetFocus
    Else
       MsgBox "„„Ì“2 «‰»«— „Ê—œ ‰Ÿ— ÅÌœ« ‰‘œ", vbInformation, ""
    End If
       
End If

End Sub

Private Sub CmdReportAddress_Click()
 Dim L As Long
 
 L = Grid1.ActiveCell.Row
 If L > 0 Then
    MakeHavaleAddress (Grid1.Cell(L, 2).Text)
 Else
    MsgBox "»—«Ì ê“«—‘ »«Ìœ Ìò ”ÿ— «‰ Œ«» ‘Êœ", vbExclamation, ""
 End If
End Sub

Private Sub CmdReportSelection_Click()
 Dim i As Long
 Dim Havale As String
 Dim VazneKol As Long
 Dim TedadKol As Long
 Dim strSql As String
 Dim rs As New Recordset
 Dim SubStr As String, SubStrT As String
 Dim TempV As Long, TempT As Long
 Dim Radif As Long, kRow As Long
 
 Radif = 0
 kRow = 3
 FrmPreview.Grid1.OpenFile App.Path & "\RepMomayez2TajamoEE.cel"
 
 FrmPreview.Grid1.Cell(1, 1).Text = FrmGetReportINFO7.FarDate1.today
 FrmPreview.Grid1.Cell(1, 3).Text = Grid1.Cell(Grid1.ActiveCell.Row, 3).Text
 
 For i = 1 To Grid1.Rows - 1
     If Grid1.Cell(i, 7).Text = "1" Then
        Havale = Grid1.Cell(i, 5).Text
        VazneKol = Val(Grid1.Cell(i, 4).Text) ' vazneHavale
        TedadKol = Val(Grid1.Cell(i, 2).Text) ' TedadHavale
        Radif = Radif + 1
        
        strSql = "SELECT Momayez2.Havale,SUM(Vazn),SUM(Detail7.Tedad) "
        strSql = strSql & "FROM Momayez2 INNER JOIN Detail7 ON "
        strSql = strSql & "Momayez2.Havale = Detail7.Momayez2 "
        strSql = strSql & "GROUP BY Momayez2.Havale "
        strSql = strSql & "HAVING (((Momayez2.Havale)='" & Havale & "'))"
        rs.Open strSql, CNS
        On Error Resume Next
        TempV = VazneKol - rs(1)
        TempT = TedadKol - rs(2)
        Select Case TempV
            Case Is > 0: SubStr = "„ﬁœ«— " & TempV & "òÌ·Êê—„ »«ﬁÌ „«‰œÂ «” "
            Case Is < 0: SubStr = "„ﬁœ«— " & Abs(TempV) & "òÌ·Êê—„ «÷«›Â Ê“‰ œ«—œ"
            Case 0: SubStr = " ÕÊ«·Â Å«Ì«Å«Ì „Ì »«‘œ"
        End Select
        '
        Select Case TempT
            Case Is > 0: SubStrT = " ⁄œ«œ " & TempT & "»‰œ· »«ﬁÌ „«‰œÂ «” "
            Case Is < 0: SubStrT = " ⁄œ«œ " & Abs(TempT) & "»‰œ· «÷«›Â œ«—œ"
            Case 0: SubStrT = " ÕÊ«·Â Å«Ì«Å«Ì „Ì »«‘œ"
        End Select
        
        With FrmPreview.Grid1
             Dim tv As String, tt As String
             .Cell(kRow, 1).WrapText = True
             If VazneKol > 0 Then
             tv = "«“ —œÌ› " & Havale & " »Â Ê“‰ ò· " & _
                            VazneKol & "òÌ·Êê—„ „ﬁœ«— " & rs(1) & "òÌ·Êê—„ Œ«—Ã Ê " & _
                            SubStr & " "
             End If
             If TedadKol > 0 Then
             tt = ". «“ —œÌ› " & Havale & "  ⁄œ«œ ò· " & _
                            TedadKol & "»‰œ·  ⁄œ«œ " & rs(2) & "»‰œ· Œ«—Ã Ê" & _
                            " " & SubStrT
             End If
             .Cell(kRow, 1).Text = tv & vbCrLf & tt
             .RowHeight(kRow) = 50
             .Cell(kRow, 9).Text = Radif
             kRow = kRow + 1
             .InsertRow kRow, 1
             .Range(kRow, 1, kRow, 8).Merge
        End With
        rs.Close
     End If
 Next
 
 FrmPreview.Grid1.PrintPreview 100
 Set rs = Nothing
 
End Sub

Private Sub CmdSave_Click()
 Dim ts As String, i As Integer
 Dim bl As Boolean
 
 ts = CombMoshtari
 bl = False
 For i = 0 To CombMoshtari.ListCount - 1
     If ts = CombMoshtari.List(i) Then
        bl = True
        Exit For
     End If
 Next
 '
 If Not bl Then
    MsgBox "Œÿ«:’«Õ» ò«·«  ⁄—Ì› ‰‘œÂ «” ", vbCritical, ""
    CombMoshtari = ""
    CombMoshtariCode.ListIndex = -1
    Exit Sub
 End If
 '
 If TxtHavale <> Empty Then
    Dim rs As New Recordset
    Dim strSql As String
    '
    strSql = "SELECT Havale FROM Momayez2 "
    strSql = strSql & "WHERE Havale='" & TxtHavale & "' "
    rs.Open strSql, CNS
    If rs.EOF Then ' no duplicate
       rs.Close
       strSql = "INSERT INTO Momayez2 "
       strSql = strSql & "VALUES('" & TxtHavale & "'," & Val(TxtWeight)
       strSql = strSql & "," & Val(TxtTedad) & "," & Val(CombMoshtariCode) & ")"
       rs.Open strSql, CNS
       '
       Grid1.AddItem CombMoshtariCode & vbTab & TxtTedad & vbTab & CombMoshtari & vbTab & _
                     TxtWeight & vbTab & TxtHavale & vbTab & Grid1.Rows
       '
       TxtHavale = Empty
       TxtWeight = Empty
       TxtTedad = Empty
       CombMoshtari.ListIndex = -1
       TxtHavale.SetFocus
    Else
       rs.Close
       MsgBox "„„Ì“2 «‰»«—  ò—«—Ì «” ", vbExclamation, ""
       TxtHavale.SetFocus
       SendKeys "{home}+{end}"
    End If
    Set rs = Nothing
 Else
    MsgBox "·ÿ›« „„Ì“2 «‰»«— —« Ê«—œ ‰„«ÌÌœ", vbExclamation, ""
    TxtHavale.SetFocus
    SendKeys "{home}+{end}"
 End If
End Sub

Private Sub CmdTafsili_Click()
 Dim L As Long
 
 L = Grid1.ActiveCell.Row
 If L > 0 Then
    MakeHavaleReport (Grid1.Cell(L, 5).Text)
 Else
    MsgBox "»—«Ì ê“«—‘ »«Ìœ Ìò ”ÿ— «‰ Œ«» ‘Êœ", vbExclamation, ""
 End If

End Sub

Private Sub CmdTikAll_Click()
 Dim i As Long
 For i = 1 To Grid1.Rows - 1
     Grid1.Cell(i, 7).Text = "1"
 Next
End Sub

Private Sub CmdunTikALL_Click()
 Dim i As Long
 For i = 1 To Grid1.Rows - 1
     Grid1.Cell(i, 7).Text = "0"
 Next

End Sub

Private Sub CombMoshtari_Change()
 If CombMoshtari.Text = Empty Then CombMoshtariCode.ListIndex = -1
End Sub

Private Sub CombMoshtari_Click()
 CombMoshtariCode.ListIndex = CombMoshtari.ListIndex
 If CombMoshtari.ListIndex = CombMoshtari.ListCount - 1 Then
     FrmMoshtari.Show 1
     CombMoshtari.Clear
     CombMoshtariCode.Clear
     Call LoadMoshtari
 End If
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

End Sub

Private Sub Form_Load()
 CenterForm Me
 ClearText Me
 ImgBackground.Picture = LoadPicture(App.Path & "\Images\BackForms7.jpg")
 '
 Call SetGrid
 Call LoadHavale
 Call LoadMoshtari
 
End Sub

Private Sub Grid1_Click()
 Dim C As Integer, r As Long
 C = Grid1.ActiveCell.Col
 r = Grid1.ActiveCell.Row
 If C = 7 Then
    If Grid1.Cell(r, C).Text = "0" Then
       Grid1.Cell(r, C).Text = "1"
    Else
       Grid1.Cell(r, C).Text = "0"
    End If
 End If
End Sub

Private Sub LblTitle_DblClick()
 CmdClose_Click
End Sub

Private Sub LblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

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

Private Sub SetGrid()
 Dim i As Integer
 With Grid1
      .Cols = 8
      .Rows = 1
      '
      .DefaultFont.Name = "B Traffic"
      .DefaultFont.Bold = True
      .DefaultFont.Size = 11
      
      .DefaultRowHeight = 24
      .AllowUserResizing = False
      '.AllowUserSort = True
      '
      .BackColor1 = RGB(111, 200, 81)
      .BackColor2 = RGB(108, 181, 100)
      '.BackColorBkg = vbBlack
      '.BackColorFixed = RGB(255, 215, 179)
      '.BackColorScrollBar = &H80FF&    'RGB(255, 125, 199)
      '
      .Column(0).Width = 15
      .Column(1).Width = 0
      .Column(2).Width = 50
      .Column(3).Width = 120
      .Column(4).Width = 100
      .Column(5).Width = 100
      .Column(6).Width = 45
      .Column(7).Width = 20
      '
      .Column(7).CellType = cellCheckBox
      For i = 1 To .Cols - 1
          .Column(i).Alignment = cellCenterCenter
      Next
      'Make Titr
      .Cell(0, 1).Text = "" 'Code Saheb kala
      .Cell(0, 2).Text = " ⁄œ«œ"
      .Cell(0, 3).Text = "’«Õ» ò«·«"
      .Cell(0, 4).Text = "Ê“‰ „„Ì“2 «‰»«—"
      .Cell(0, 5).Text = "„„Ì“2 «‰»«—"
      .Cell(0, 6).Text = "—œÌ›"
      .Cell(0, 7).Text = ""
      '
      .ReadOnly = True
      .ReadOnlyFocusRect = Solid
      '
      .SelectionMode = cellSelectionFree
      '.Appearance = Flat
      
 End With
End Sub

Private Sub LoadHavale()
 Dim rs As New Recordset
 Dim strSql As String
 Dim r As Long
 Dim Saheb As String
 '
 strSql = "SELECT Havale,Weight,Tedad,MoshtariCODE,"
 strSql = strSql & "(SELECT MoshtariName FROM Moshtari WHERE "
 strSql = strSql & "Moshtari.MoshtariCODE=Momayez2.MoshtariCode )AS MC "
 strSql = strSql & "FROM Momayez2 "
 rs.Open strSql, CNS
 r = 1
 Do While Not rs.EOF
    Grid1.AddItem rs(3) & vbTab & rs(2) & vbTab & _
                  rs(4) & vbTab & rs(1) & vbTab & rs(0) & vbTab & r
    rs.MoveNext
    r = r + 1
 Loop
 
 rs.Close
 Set rs = Nothing
 'Set rs1 = Nothing
End Sub

Private Sub MakeHavaleAddress(Havale As String)
 Dim strSql As String
 Dim rs As New Recordset
 Dim i As Integer
 Dim sumTedad As Long, sumVazn As Long
 Dim CurrentAddress As String
 Dim strPrompt As String
 
 '
 strSql = "SELECT Name,Etebar,Parvane,Part,BarName, "
 strSql = strSql & "Tarikh,Address,Havale,Momayez2,ShomareMashin, "
 strSql = strSql & "Vazn ,Tedad,Size0,Keraye,Mobile,Parvande, Count0,Main7.Code "
 strSql = strSql & "FROM Main7 INNER JOIN Detail7 ON Main7.Code = Detail7.Code "
 strSql = strSql & "WHERE Havale='" & Havale & "' "
 strSql = strSql & "ORDER BY Address "
 
 rs.Open strSql, CNS
 If Not rs.EOF Then
    With FrmPreview.Grid1
        .OpenFile App.Path & "\RepMomayez2TajamoEE.cel"
        i = 0
        Do While Not rs.EOF
           CurrentAddress = rs("Address")
           sumTedad = sumTedad + rs("Tedad")
           sumVazn = sumVazn + rs("Vazn")
           i = i + 1
           rs.MoveNext
           If rs.EOF Then GoTo ss:
           If CurrentAddress <> rs("Address") Then
              strPrompt = CurrentAddress & "  ⁄œ«œ " & sumTedad & " »‰œ·  Ê”ÿ " & i & " œ” ê«Â ò«„ÌÊ‰ " & " »Â Ê“‰ " & sumVazn & " Œ«—Ã ‘œÂ «”  "
ss:           FrmPreview.List1.AddItem strPrompt
              sumTedad = 0: sumVazn = 0: i = 0
           End If
           
        Loop
        rs.Close
        Call MolahezatAddress
       ' Call TedadBaghimande
        Call PageSetupANDFooter
        
        .PrintPreview 110
        FrmPreview.Show 1
    End With
 Else
    rs.Close
    MsgBox "ê“«—‘Ì »—«Ì ÕÊ«·Â „Ê—œ ‰Ÿ— „ÊÃÊœ ‰„Ì »«‘œ", vbInformation, ""
 End If

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

Private Sub MolahezatAddress()
 Dim r As Integer
 Dim i As Integer
 With FrmPreview.Grid1
      .Cell(1, 3).Text = "»Â  ›òÌò ¬œ—” "
      For i = 0 To FrmPreview.List1.ListCount - 1
          r = .Rows - 2
          .AddItem ""
         .Range(r, 1, r, 8).Merge
         .RowHeight(r) = 32
         .Range(r, 1, r, 9).Alignment = cellRightCenter
         .Range(r, 1, r, 9).FontName = "B Nazanin"
         .Range(r, 1, r, 9).FontBold = True
         .Range(r, 1, r, 9).FontSize = 12
         .Cell(r, 1).Text = FrmPreview.List1.List(i)
         '
         .Cell(r, 9).Text = i + 1
      Next
      '
      .RemoveItem (r)
 End With
 
 FrmPreview.List1.Clear
End Sub

Private Sub MakeHavaleReport(Havale As String)

 Dim strSql As String
 Dim rs As New Recordset
 Dim i As Integer, kRow As Integer
 Dim sumVazn As Long
 Dim sumTedad As Long
 '
 strSql = "SELECT Name,Etebar,Parvane,Part,BarName, "
 strSql = strSql & "Tarikh,Address,Havale,Momayez2,ShomareMashin, "
 strSql = strSql & "Vazn ,Tedad,Size0,Keraye,Mobile,Parvande, Count0,Main7.Code "
 strSql = strSql & "FROM Main7 INNER JOIN Detail7 ON Main7.Code = Detail7.Code "
 strSql = strSql & "WHERE Momayez2='" & Havale & "' "
 strSql = strSql & "ORDER BY Main7.Code,Count0 "
 
 rs.Open strSql, CNS
 
 If Not rs.EOF Then
    Dim ss As String
    Dim rs1 As New Recordset
    Dim Saheb As String
    ss = "SELECT moshtaricode FROM Main7 "
    ss = ss & "WHERE Parvane='" & rs("Parvane") & "'"
    rs1.Open ss, CNS
    Saheb = IIf(IsNull(rs1(0)), "", rs1(0))
    rs1.Close
    ss = "SELECT MoshtariName FROM Moshtari "
    ss = ss & "WHERE moshtaricode=" & Val(Saheb)
    rs1.Open ss, CNS
    Saheb = IIf(IsNull(rs1(0)), "", rs1(0))
    rs1.Close
 
    With FrmPreview.Grid1
        .OpenFile App.Path & "\Rep7Almas.cel"
        kRow = 3
        sumVazn = 0
        sumTedad = 0
        Do While Not rs.EOF
           For i = 0 To 15
               .Cell(kRow, i + 1).Text = IIf(IsNull(rs(15 - i)), "", rs(15 - i))
           Next
           sumVazn = sumVazn + rs("Vazn")
           sumTedad = sumTedad + rs("Tedad")
           .Cell(kRow, 17).Text = kRow - 2
           kRow = kRow + 1
           .InsertRow kRow, 2
           rs.MoveNext
        Loop
        rs.Close
        
        .Cell(1, 4).Text = "ê“«—‘ ò·Ì «“ ÕÊ«·Â ‘—ò  " & " " & Saheb
        .Cell(1, 2).Text = .Cell(1, 2).Text & Space(8) & FrmGetReportINFO7.FarDate1.today
        
        Call Molahezat(sumVazn, Havale, sumTedad)
        Call PageSetupANDFooter
        
        .PrintPreview 100
        FrmPreview.Show 1
    End With
Else
    MsgBox "ê“«—‘Ì »—«Ì «Ì‰ ÕÊ«·Â ÊÃÊœ ‰œ«—œ", vbExclamation, ""
End If

End Sub

Private Sub Molahezat(VazneVarede As Long, Havale As String, TedadeVarede As Long)
'tafsili
 Dim r As Integer
 Dim VazneKolHavale As Long
 Dim TedadKolHavale As Long
 Dim SubStr As String, SubStrT As String
 Dim TempV As Long, TempT As Long
 
 VazneKolHavale = Val(Grid1.Cell(Grid1.ActiveCell.Row, 4).Text)
 TedadKolHavale = Val(Grid1.Cell(Grid1.ActiveCell.Row, 2).Text)
 
 TempV = VazneKolHavale - VazneVarede
 TempT = TedadKolHavale - TedadeVarede
 
 Select Case TempV
     Case Is > 0: SubStr = "„ﬁœ«— " & TempV & "òÌ·Êê—„ »«ﬁÌ „«‰œÂ «” "
     Case Is < 0: SubStr = "„ﬁœ«— " & Abs(TempV) & "òÌ·Êê—„ «÷«›Â Ê“‰ œ«—œ"
     Case 0: SubStr = " ÕÊ«·Â Å«Ì«Å«Ì „Ì »«‘œ"
 End Select
 
 Select Case TempT
     Case Is > 0: SubStrT = " ⁄œ«œ " & TempT & "»‰œ· »«ﬁÌ „«‰œÂ «” "
     Case Is < 0: SubStrT = " ⁄œ«œ " & Abs(TempT) & "»‰œ· «÷«›Â œ«—œ"
     Case 0: SubStrT = " ÕÊ«·Â Å«Ì«Å«Ì „Ì »«‘œ"
 End Select
 
 With FrmPreview.Grid1
      r = .Rows - 2
     .Range(r, 4, r, 14).Merge
     If VazneKolHavale > 0 Then
     .Cell(r - 1, 4).Alignment = cellCenterCenter
     .Cell(r - 1, 4).Font.Name = "Arial"
     .Cell(r - 1, 4).Font.Bold = True
     .Cell(r - 1, 4).Font.Size = 13
     .Cell(r - 1, 4).Text = ".„·«ÕŸ« : «“ —œÌ› " & Havale & " »« Ê“‰ ò· " & _
                        VazneKolHavale & "òÌ·Êê—„ „ﬁœ«— " & VazneVarede & "òÌ·Êê—„ Œ«—Ã Ê" & _
                        " " & SubStr
     End If
     '
     .Range(r - 1, 4, r - 1, 14).Merge
     If TedadKolHavale > 0 Then
     .Cell(r, 4).Alignment = cellCenterCenter
     .Cell(r, 4).Font.Name = "Arial"
     .Cell(r, 4).Font.Bold = True
     .Cell(r, 4).Font.Size = 13
     .Cell(r, 4).Text = ".„·«ÕŸ« : «“ —œÌ› " & Havale & "  ⁄œ«œ ò· " & _
                        TedadKolHavale & "»‰œ·  ⁄œ«œ " & TedadeVarede & "»‰œ· Œ«—Ã ‘œÂ «” " & _
                        " " & SubStrT
     End If
 End With
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

Private Sub TxtWeight_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
 End If

End Sub

Private Sub LoadMoshtari()
 Dim rs As New Recordset
 rs.Open "SELECT * FROM Moshtari ORDER BY MoshtariName ", CNS
 Do While Not rs.EOF
    CombMoshtari.AddItem rs("MoshtariName")
    CombMoshtariCode.AddItem rs("MoshtariCODE")
    rs.MoveNext
 Loop
 CombMoshtari.AddItem "<.....>"
 CombMoshtariCode.AddItem "n"
 rs.Close
 Set rs = Nothing

End Sub
