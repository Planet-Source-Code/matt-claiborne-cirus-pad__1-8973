VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Cirus Pad"
   ClientHeight    =   6465
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text1 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9763
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Form1.frx":0442
   End
   Begin MSComctlLib.StatusBar sb2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   6210
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8520
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu sdfasdfdsf 
         Caption         =   "-"
      End
      Begin VB.Menu psetup 
         Caption         =   "Print Setup"
      End
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu fsdg 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu dsf 
      Caption         =   "&Edit"
      Begin VB.Menu undo 
         Caption         =   "Undo"
      End
      Begin VB.Menu asdasd 
         Caption         =   "-"
      End
      Begin VB.Menu cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu sdfsadfasdf 
         Caption         =   "-"
      End
      Begin VB.Menu selall 
         Caption         =   "Select All"
      End
      Begin VB.Menu td 
         Caption         =   "Time/Date"
      End
   End
   Begin VB.Menu format 
      Caption         =   "&Format"
      Begin VB.Menu color 
         Caption         =   "Color.."
      End
      Begin VB.Menu font 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu cp 
      Caption         =   "&Cirus Pad"
      Begin VB.Menu uptime 
         Caption         =   "UpTime?"
      End
      Begin VB.Menu dclc 
         Caption         =   "Disable Char/Line Counting"
      End
      Begin VB.Menu wordcount 
         Caption         =   "Word Count"
      End
      Begin VB.Menu dghfg 
         Caption         =   "-"
      End
      Begin VB.Menu encrypt 
         Caption         =   "Encrypt"
      End
      Begin VB.Menu decrypt 
         Caption         =   "DeEncrypt"
      End
      Begin VB.Menu gddddgd 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About...."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CharCount As Boolean
Option Explicit
Private Sub about_Click()
Form2.Show
End Sub

Private Sub color_Click()
On Error Resume Next
cd1.ShowColor
Text1.SelColor = cd1.Color
End Sub


Private Sub Command1_Click()
MsgBox Text1.Text & vbTab & Len(Text1.Text)
End Sub

Private Sub copy_Click()
If Text1.SelLength > 0 Then SendKeys ("^c")
End Sub

Private Sub cut_Click()
If Text1.SelLength > 0 Then SendKeys ("^x")
End Sub

Private Sub dclc_Click()
If CharCount = True Then
CharCount = False
Chars_Lines
dclc.Caption = "Enable Char/Line Counting"
Else
CharCount = True
Chars_Lines
dclc.Caption = "Disable Char/Line Counting"
End If

End Sub

Private Sub decrypt_Click()
Dim AsciiOf As Integer
Dim NewText As String
Dim OldText As String
Dim x As Long
OldText = Text1.Text
Form1.Caption = "Cirus Pad - DeEncrypting..."
Text1.Text = "DeEncrypting..."

For x = 1 To Len(OldText)
    DoEvents
    AsciiOf = Asc(Mid(OldText, x, 1))
    If AsciiOf <= 25 Then AsciiOf = AsciiOf + 255
    NewText = NewText & Chr(AsciiOf - 25)
Next

Text1.Text = NewText
Form1.Caption = "Cirus Pad"
Call Chars_Lines
End Sub

Private Sub delete_Click()
If Text1.SelLength > 0 Then SendKeys "{DEL}"
End Sub

Private Sub encrypt_Click()
Dim Letter1 As String
Dim AsciiOf As Integer
Dim NewText As String
Dim MemText As String
Dim x As Long
MemText = Text1.Text
Form1.Caption = "Cirus Pad - Encrypting...."
Text1.Text = "Encrypting..."
For x = 1 To Len(MemText)
DoEvents
Letter1 = Mid(MemText, x, 1)
AsciiOf = Asc(Letter1)
AsciiOf = AsciiOf + 25
If AsciiOf > 255 Then AsciiOf = AsciiOf - 255
NewText = NewText & Chr(AsciiOf)
Next
Text1.Text = NewText
Form1.Caption = "Cirus Pad"
Call Chars_Lines
End Sub

Private Sub exit_Click()
Unload Me
End
End Sub

Private Sub ExitMe()
Dim a As Integer
If Text1.Text <> "" Then
    a = MsgBox("Would you like to save before exiting?", vbYesNoCancel, "Save?")
    If a = vbYes Then
        Call save_Click
    End If
    If a = vbCancel Then
        Exit Sub
    End If
End If

Unload Me
Unload Form2
End

End Sub

Private Sub font_Click()
cd1.Flags = cdlCFScreenFonts
cd1.ShowFont
Text1.SelFontName = cd1.FontName
Text1.SelBold = cd1.FontBold
Text1.SelItalic = cd1.FontItalic
Text1.SelFontSize = cd1.FontSize
Text1.SelStrikeThru = cd1.FontStrikethru
Text1.SelUnderline = cd1.FontUnderline
End Sub

Private Sub Form_Load()
Me.Show
Me.Refresh
Load Form2

sb2.SimpleText = "Char:0/0  Line:1/1"
If Form1.Caption = "Cirus Pad" Then
Else
Form1.Caption = "Cirus Pad"
End If
CharCount = True
sb2.Panels(1) = "hi"
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form1.WindowState = 0 Then
Text1.Height = Form1.Height - 970
Text1.Width = Form1.Width - 120
End If
If Form1.WindowState = 2 Then
Text1.Height = Form1.Height - 970
Text1.Width = Form1.Width - 120
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ExitMe
Cancel = 1
End Sub

Private Sub open_Click()
On Error GoTo err
Dim i As Long
Form1.Caption = "Cirus Pad - Opening..."
Text1.Text = ""
cd1.Filter = "Txt (*.txt)|*.txt|Any File (*.*)|*.*"
cd1.ShowOpen
If cd1.FileName <> "" Then
Dim t As Long
i = FreeFile
Open cd1.FileName For Input As #i
CharCount = True
If Int(LOF(i) / 1000) > 300 Then
    If MsgBox("File is > 300kb would you like to disable Charactor coutning?", vbYesNo, "Large FIle") = vbYes Then
    CharCount = False
    dclc.Caption = "Enable Char/Line Counting"
    Chars_Lines
    End If
End If

Text1.Text = Input(LOF(i), i)
Close #i
Else
Form1.Caption = "Cirus Pad"
Exit Sub
End If


Form1.Caption = "Cirus Pad"
Call Chars_Lines
Exit Sub



err:
Form1.Caption = "Cirus Pad - Opening in binary..."
Close #i
Open cd1.FileName For Binary As #i
Text1.Text = Input(LOF(i), i)
Close #i
Form1.Caption = "Cirus Pad"
Exit Sub




End Sub

Private Sub paste_Click()
SendKeys ("^v")
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub Print_Click()
Printer.Print Text1.Text
End Sub

Private Sub psetup_Click()
cd1.ShowPrinter

End Sub

Private Sub save_Click()
'On Error GoTo err
Form1.Caption = "Cirus Pad - Saving..."
Dim a As String
cd1.Filter = "Txt (*.txt)|*.txt|Html File (*.Html)|*.Html|Any File (*.*)|*.*"
cd1.ShowSave
If cd1.FileName <> "" Then
Open cd1.FileName For Output As #1
Print #1, Text1.Text
Close 1
End If
Form1.Caption = "Cirus Pad"
Exit Sub

err:
Form1.Caption = "Cirus Pad"
MsgBox "Error saving file"
End Sub

Private Sub selall_Click()
Dim a As String
Text1.SelStart = 0

Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub td_Click()
SendKeys (Now)
End Sub




Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Call Chars_Lines

End Sub

Private Sub Chars_Lines()
If CharCount = True Then
Dim Lines, Chars As String
Dim blah() As String
Dim bleh() As String
Dim Curline As String
Dim CurChar, TotalChar As String

Curline = Mid(Text1.Text, 1, Text1.SelStart)
blah() = Split(Curline, Chr$(10))
bleh() = Split(Text1.Text, Chr$(10))

If Text1.SelStart = 0 Then
CurChar = 0
Curline = 1

If Len(Text1.Text) = 0 Then
TotalChar = 0
Else
TotalChar = Len(Text1.Text) - (UBound(bleh) * 2)

End If

Else
CurChar = Text1.SelStart - (UBound(blah) * 2)
Curline = UBound(blah) + 1
TotalChar = Len(Text1.Text) - (UBound(bleh) * 2)
End If



Lines = "Line:" & Curline & "/" & SendMessage(Text1.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
Chars = "Char:" & CurChar & "/" & TotalChar

sb2.SimpleText = Chars & "  " & Lines

Else
sb2.SimpleText = "Off"
End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Chars_Lines
End Sub

Private Sub Text1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim numFiles As Integer
    numFiles = Data.Files.Count
    If numFiles = 1 Then
    'Add all dropped files into the list

        'File or directory?
        If (GetAttr(Data.Files(1)) And vbDirectory) = vbDirectory Then
            Else



On Error GoTo err
Dim i As Long
Form1.Caption = "Cirus Pad - Opening..."
Text1.Text = ""
i = FreeFile
Open Data.Files(1) For Input As #i
CharCount = True
If Int(LOF(i) / 1000) > 300 Then
    If MsgBox("File is > 300kb would you like to disable Charactor coutning?", vbYesNo, "Large FIle") = vbYes Then CharCount = False: dclc.Caption = "Enable Char/Line Counting"
End If

Text1.Text = Input(LOF(i), i)
Close #i




Form1.Caption = "Cirus Pad"
Call Chars_Lines
Exit Sub



err:
Form1.Caption = "Cirus Pad - Opening in binary..."
Close #i
Open Data.Files(1) For Binary As #i
Text1.Text = Input(LOF(i), i)
Close #i
Form1.Caption = "Cirus Pad"
Exit Sub




End If

    End If

  
End Sub

Private Sub undo_Click()
SendKeys ("^z")

End Sub

Private Sub uptime_Click()
    Dim Secs, Mins, Hours, Days As Long
    Dim TotalMins, TotalHours, TotalSecs, TempSecs As Long
    Dim CaptionText As String
    TotalSecs = Int(GetTickCount / 1000)
    Days = Int(((TotalSecs / 60) / 60) / 24)
    TempSecs = Int(Days * 86400)
    TotalSecs = TotalSecs - TempSecs
    TotalHours = Int((TotalSecs / 60) / 60)
    TempSecs = Int(TotalHours * 3600)
    TotalSecs = TotalSecs - TempSecs
    TotalMins = Int(TotalSecs / 60)
    TempSecs = Int(TotalMins * 60)
    TotalSecs = (TotalSecs - TempSecs)


    If TotalHours > 23 Then
        Hours = (TotalHours - 23)
    Else
        Hours = TotalHours
    End If


    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If
    CaptionText = "Your Computer has been up: " & Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds" & vbCrLf

    MsgBox CaptionText, vbOKOnly, "Up Time"
    Clipboard.Clear
    Clipboard.SetText CaptionText
End Sub

Private Sub wordcount_Click()
Dim a() As String
Dim b() As String
Dim wordcount As Long
Dim x As Long
Form1.Caption = "Cirus Pad - Counting Words..."

a() = Split(Text1.Text, " ")
wordcount = UBound(a)
For x = 0 To UBound(a)
If a(x) = "" Then
wordcount = wordcount - 1
End If
Next

b() = Split(Text1.Text, Chr$(10))
wordcount = wordcount + UBound(b)
For x = 0 To UBound(b)
If b(x) = "" Then
wordcount = wordcount - 1
End If
Next
If wordcount = -2 Then wordcount = -1
Form1.Caption = "Cirus Pad"
MsgBox "There are: " & wordcount + 1 & " Words", vbOKOnly, "Word"


End Sub
