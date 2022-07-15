VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›—„ ÅÌ‘ ‰„«Ì‘"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "B Nazanin"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   14160
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstCodeParvane 
      Height          =   1230
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog ComDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FlexCell.Grid Grid1 
      Align           =   1  'Align Top
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   17806
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin HaftAlmas.TypeButton CmdClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   10200
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPreview.frx":15162
      PICN            =   "FrmPreview.frx":1517E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdExcel 
      Height          =   495
      Left            =   11400
      TabIndex        =   2
      Top             =   10200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "–ŒÌ—Â »Â ›—„ «ò”·"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   4210688
      BCOLO           =   32896
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPreview.frx":18BEC
      PICN            =   "FrmPreview.frx":18C08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdPreview 
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   10200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ÅÌ‘ ‰„«Ì‘ œÊ»«—Â"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPreview.frx":1C524
      PICN            =   "FrmPreview.frx":1C540
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HaftAlmas.TypeButton CmdComment 
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   10200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "«÷«›Â ò—œ‰  Ê÷ÌÕ« "
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPreview.frx":1FDF8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
 Unload Me
End Sub

Private Sub CmdComment_Click()
 With Grid1
    .AddItem ""
    .Range(.Rows - 1, 1, .Rows - 1, .Cols - 1).Merge
    .Cell(.Rows - 1, 1).Alignment = cellCenterCenter
    .Cell(.Rows - 1, 1).Font.Name = "B Titr"
    .Cell(.Rows - 1, 1).Font.Bold = True
    .Cell(.Rows - 1, 1).Font.Size = 10
    .RowHeight(.Rows - 1) = 32
 End With
End Sub

Private Sub CmdExcel_Click()
 With ComDialog1
     .DialogTitle = "„”Ì—Ì —« »—«Ì –ŒÌ—Â ”«“Ì «‰ Œ«» ò‰Ìœ"
     .Filter = "Excel Files(*.xls)|*.xls"
     .ShowSave
     If .filename <> Empty Then
        Grid1.ExportToExcel .filename
        Call ConvertToExcel(.filename)
     End If
 End With
End Sub

Private Sub CmdPreview_Click()
 Grid1.PrintPreview 95
End Sub

Private Sub Form_Load()
 Me.BackColor = RGB(207, 219, 183)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'''' Dim msg As Integer
'''' Dim initPath As String
'''' Dim DoSave As Boolean
'''' Dim AskSave As Boolean ' if false then Save in Default path
''''
'''' If Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "DoSave", "0")) <> Trim("0") Then
''''    If Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "DoSave", "0")) = "1" Then
''''       Exit Sub
''''    Else
''''       DoSave = True
''''    End If
'''' End If
''''
'''' If DoSave Then
''''    If Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "AskSave", "0")) <> Trim("0") Then
''''       If Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "AskSave", "0")) = "1" Then
''''          AskSave = False
''''          If Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "FilePath", "")) <> Trim("") Then
''''             initPath = Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "FilePath", ""))
''''          End If
''''          If initPath <> Empty Then
''''             Dim inp As String
''''             inp = InputBox("‰«„Ì »—«Ì ò“«—‘  «ÌÅ ò‰Ìœ", "–ŒÌ—Â ê“«—‘ ")
''''             On Error Resume Next
''''             Grid1.SaveFile initPath & "\" & inp & ".cel"
''''             MsgBox "œ— „”Ì— ÅÌ‘ ›—÷ –ŒÌ—Â ‘œ", vbInformation, ""
''''          End If
''''       Else
''''          msg = MsgBox("¬Ì« „«Ì· »Â –ŒÌ—Â ò—œ‰ ›«Ì· ê“«—‘ „Ì »«‘Ìœø", vbQuestion + vbYesNo, "")
''''          If msg = vbYes Then
''''             With ComDialog1
''''                  If Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "FilePath", "")) <> Trim("") Then
''''                     initPath = Trim(GetSetting("HKEY_CURRENT_MACHINE", "Nafis", "FilePath", ""))
''''                  End If
''''                  .InitDir = initPath
''''                  .Filter = "Grid Files(*.cel)|*.Cel"
''''                  .ShowSave
''''                  If .FileName <> Empty Then Grid1.SaveFile .FileName
''''                  initPath = .FileName
''''                  SaveSetting "HKEY_CURRENT_MACHINE", "Nafis", "FilePath", initPath
''''             End With
''''          End If ' msg
''''       End If ' Else
''''    End If
'''' End If ' Do save
End Sub

Private Sub ConvertToExcel(filename As String)
 Dim X_Excel As Excel.Application
 Dim X_WorkBook As Excel.Workbook
 Dim X_WorkSheet As Excel.Worksheet
 
 Set X_Excel = New Excel.Application
 Set X_WorkBook = X_Excel.Workbooks.Open(filename)
 Set X_WorkSheet = X_WorkBook.Worksheets(1)
 
 With X_WorkSheet.PageSetup
   .HeaderMargin = 19.6
   .FooterMargin = 25.31
   .TopMargin = 28.57
   .BottomMargin = 53.88
   .LeftMargin = 13.89
   .RightMargin = 19.6
   .Orientation = xlLandscape
   .Zoom = 86
   .BlackAndWhite = True
   .PrintGridlines = True
   .CenterHorizontally = True
   .LeftFooter = "’›ÕÂ :" & "&P" & " «“ " & "&N"
   .CenterFooter = " «ÌÌœ ò‰‰œÂ:"
   .RightFooter = " ‰ŸÌ„ ò‰‰œÂ:"
     
 End With

 X_WorkBook.Save
 X_Excel.Quit
 
 Set X_Excel = Nothing
 Set X_WorkBook = Nothing
 Set X_WorkSheet = Nothing
 
 MsgBox "–ŒÌ—Â ”«“Ì «‰Ã«„ ‘œ", vbInformation
 
End Sub
