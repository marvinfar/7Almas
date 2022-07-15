Attribute VB_Name = "ModMain"
Option Explicit


Public CNS As String
Public isEnglish As Boolean
''''
Public Declare Function ActivateKeyboardLayout Lib "User32" (ByVal HKL As Long, ByVal Flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "User32" (ByVal dwLayout As Long) As Long
Public Const HKL_NEXT = 1
Public Const HKL_PREV = 0




Public Sub Main()
'Dim HardSerial As Long
' HardSerial = GetHardSerial("d:")
'
'   If HardSerial <> 6502413 And HardSerial <> 407336769 Then
'     MsgBox "‘„« «Ã«“Â «” ›«œÂ «“ »—‰«„Â —« ‰œ«—Ìœ" & vbCrLf & _
'            "·ÿ›« »—«Ì «ÿ·«⁄«  »Ì‘ — »« ‘„«—Â 09112320258  „«” »êÌ—Ìœ", vbExclamation
'     End
'   End If
  '
  
  '
  CNS = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db7Almas.mdb;Persist Security Info=False"
  '
  'Call ChangeResolution(1024, 768)
  FrmBackGround.Show
  
  '''
End Sub

Public Sub CenterForm(XForm As Form)
   XForm.Left = (Screen.Width - XForm.ScaleWidth) / 2
   XForm.Top = (Screen.Height - XForm.ScaleHeight) / 2
End Sub

Public Sub setFormsProperties(Frm As Form, Optional Title As String = Empty, Optional WinState As Byte = 0)
  If Not TypeOf Frm Is MDIForm Then
     Frm.KeyPreview = True
  End If
   'Frm.RightToLeft = True
   Frm.Caption = Title
   Frm.WindowState = WinState
End Sub

Public Sub ClearText(Frm As Form)
 Dim oCtrl As Object
 
  'Clear Text Boxes
   For Each oCtrl In Frm.Controls
       If TypeOf oCtrl Is TextBox Then
          If Not oCtrl.Locked Then
             oCtrl.Text = Empty
          End If
       End If
       '
       If TypeOf oCtrl Is CheckBox Then oCtrl.Value = 0
       If TypeOf oCtrl Is ComboBox Then oCtrl.ListIndex = -1
       
       If TypeOf oCtrl Is ComboBox Then
          If oCtrl.Style = 0 Then
             oCtrl.Text = Empty
          End If
       End If
   Next
  '
End Sub

Function SearchCombo(Comb1, Comb2 As ComboBox, key As Variant) As String
 Dim i As Integer
   
   SearchCombo = ""
   For i = 0 To Comb1.ListCount - 1
       If Val(key) = Val(Comb1.List(i)) Then
          SearchCombo = Comb2.List(i)
          Exit Function
       End If
   Next
End Function


Public Function CompactDB(pFileName As String) As Boolean
On Error GoTo ErrH
Dim CONN As New JRO.JetEngine
Dim ConnstringSorg As String, ConnstringDest As String

' Ensure file is not read only
SetAttr pFileName, vbNormal
ConnstringSorg = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
pFileName & ";User ID=;Password=;"
ConnstringDest = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
App.Path & "\Temp.mdb" & ";Jet OLEDB:Engine Type=5;"

Screen.MousePointer = vbHourglass
CONN.CompactDatabase ConnstringSorg, ConnstringDest
Screen.MousePointer = vbDefault

'Copia il file compattato.
Kill pFileName
FileCopy App.Path & "\Temp.mdb", pFileName
Kill App.Path & "\Temp.mdb"

Set CONN = Nothing
CompactDB = True
Exit Function
ErrH:
Screen.MousePointer = vbDefault
'MsgBox Err.Description, "DELTA"
End Function

Public Function Trunc(X As Double, Num As Integer) As Double
  Dim Dec As Integer  ' Decimal Places
    
    Dec = 10 ^ Num
    
    Trunc = Int(X * Dec) / Dec
    
End Function

Public Function DigitGrouping(Number As Currency) As String
 Dim s$, t$
 Dim i%, L%
 
 s = CStr(Number)
 L = Len(s)
 
 For i = 1 To L \ 3
     t = "," & Right(s, 3) & t
     s = Left(s, Len(s) - 3)
 Next
 If L Mod 3 <> 0 Then t = s & t
 If Left(t, 1) = "," Then t = Mid(t, 2)
 
 DigitGrouping = t
End Function

Function MakeAutoNumber(TableName As String, Field As String) As Long
  Dim rs As New Recordset
  Dim cnt As Long
  
  rs.Open "SELECT MAX(" & Field & ") FROM " & TableName, CNS
  '
  cnt = IIf(IsNull(rs(0)), 1, rs(0) + 1)
  '
  MakeAutoNumber = cnt
  '
  rs.Close
  Set rs = Nothing
End Function

Function SearchInGrid(ByRef StartRow As Long, ByRef rowFind As Long, xGrid As Grid, key As Variant, Col As Integer) As Boolean
 Dim i As Long
 Dim b As Boolean
 
  With xGrid
     b = False
     For i = StartRow To .Rows - 1
         If LCase(.Cell(i, Col).Text) = LCase(key) Then
            b = True
            StartRow = i + 1
            Exit For
         End If
     Next
     '
     rowFind = i
     SearchInGrid = b
  End With
End Function

Function Text2Currency(Text As String) As Currency
  If Format(Text) = " —Ì«·" Or Format(Text) = "" Then
     Text2Currency = 0
  Else
     Text2Currency = CCur(Text)
  End If
End Function

Function PlusToFarsiDate(FarsiDate As String, n As Integer) As String
 Dim d As Byte, m As Byte, Y As Integer
 Dim t As Integer
 Dim lDay As Byte
 Dim im As Integer, iy As Integer
 Dim hm As Integer, hy As Integer
 Dim HlpMah As Integer
 Dim MakedDate As String
 '
 Y = Val(Mid(FarsiDate, 1, 4))
 m = Val(Mid(FarsiDate, 6, 2))
 d = Val(Mid(FarsiDate, 9, 2))
 '
 If n >= 0 Or 1 = 1 Then
    t = d + n
    If m <= 6 Then
       lDay = 31
    Else
       lDay = 30
    End If
    '
    im = 0: iy = 0
    If t > lDay Then
       hm = t - lDay
       im = im + 1
       Do While hm > lDay
          hm = hm - lDay
          im = im + 1
          HlpMah = m + im
          Select Case HlpMah / 6
              Case 0 To 1: lDay = 31
              Case 1.1 To 2: lDay = 30
              Case 2.2 To 3: lDay = 31
              Case 3.3 To 4: lDay = 30
          End Select
       Loop
       If m + im = 12 And (d + hm) > 30 Then
          Y = Y + 1
          m = 1
          d = 30 - (d + hm)
          MakedDate = Format(CStr(Y) + "/" + CStr(m) + "/" + CStr(d), "yyyy/mm/dd")
       End If
       '
       If m + im > 12 Then
          hy = (m + im) - 12
          iy = 1
          Do While hy > 12
             hy = hy - 12
             iy = iy + 1
          Loop
       End If
       '
       Y = Y + iy
       m = (m + im) - hy
       d = hm
       MakedDate = Format(CStr(Y) + "/" + CStr(m) + "/" + CStr(d), "yyyy/mm/dd")
   Else ' agar tedade roozha kamtar az yek mah bashad
       MakedDate = Format(CStr(Y) + "/" + CStr(m) + "/" + CStr(t), "yyyy/mm/dd")
   End If
 End If
 PlusToFarsiDate = MakedDate
End Function


