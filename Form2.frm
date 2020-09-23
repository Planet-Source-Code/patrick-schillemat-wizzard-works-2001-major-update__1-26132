VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "Wizzard HTML 2001"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7335
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12938
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":0000
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "&Save As"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuundo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuins 
      Caption         =   "&Insert"
      Begin VB.Menu mnulink 
         Caption         =   "&Link"
      End
      Begin VB.Menu mnupic 
         Caption         =   "&Picture"
      End
      Begin VB.Menu mnubr 
         Caption         =   "<BR>"
      End
      Begin VB.Menu mnub 
         Caption         =   "<B></B>"
      End
      Begin VB.Menu mnuu 
         Caption         =   "<U></U>"
      End
      Begin VB.Menu mnutimedate 
         Caption         =   "&Time/Date"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
RichTextBox1.Width = Form2.ScaleWidth
RichTextBox1.Height = Form2.ScaleHeight
End Sub

Private Sub mnub_Click()
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "<b></b>"
End Sub

Private Sub mnubr_Click()
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "<br>"
End Sub

Private Sub mnucopy_Click()
If (TypeOf Screen.ActiveControl Is TextBox) Then
        Clipboard.Clear
        Clipboard.SetText Screen.ActiveControl.SelText
    End If
       mnupaste.Enabled = True
End Sub
Private Sub mnucut_Click()
    If (TypeOf Screen.ActiveControl Is TextBox) Then
        Clipboard.Clear
        Clipboard.SetText Screen.ActiveControl.SelText
        Screen.ActiveControl.SelText = ""
    End If
    mnupaste.Enabled = True
End Sub
Private Sub mnulink_Click()
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "<a href=""></a>"
End Sub
Private Sub mnupic_Click()
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "<img src="">"
End Sub
Private Sub mnupaste_Click()
 If (TypeOf Screen.ActiveControl Is TextBox) Then
        Screen.ActiveControl.SelText = Clipboard.GetText()
    End If
End Sub
Private Sub mnuopen_Click()
RichTextBox1.Visible = True
Dim Buffer1 As String, Buffer2 As String, CRLF As String
Dim FileNum As Integer
CRLF = Chr$(13) + Chr$(10)
OpenDlg.Filter = "Text Files (*.txt)|*.txt|Html Files (*.html)|*.html|All Files (*.*)|*.*"
OpenDlg.ShowOpen
If OpenDlg.FileName = "" Then Exit Sub
FileName = OpenDlg.FileName
FileNum = FreeFile
Open FileName For Input As FileNum
Do While Not EOF(FileNum)
Line Input #FileNum, Buffer1
Buffer2 = Buffer2 & Buffer1 & CRLF
Loop
Close FileNum
RichTextBox1.Text = Buffer2
RichTextBox1.Text = Buffer2
Form1.Caption = "Wizzard Works 2001" & " - [ " & OpenDlg.FileName & " ]"
Saved = False
End Sub
Private Sub mnusaveas_Click()
On Error GoTo nada
SaveDlg.CancelError = False
SaveDlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
SaveDlg.ShowSave
If SaveDlg.FileName = "" Then
SaveDlg.FileName = "c:\" & "Error.tmp"
Else
Open SaveDlg.FileName For Output As 1
Print #1, Text1.Text
Close
Saved = True
End If
nada:
Exit Sub
End Sub
Private Sub mnutimedate_Click()
Dim Text As String
Dim SelStart As Long
If RichTextBox1.SelLength > 0 Then
End If
Text = RichTextBox1.Text
SelStart = RichTextBox1.SelStart
RichTextBox1.Text = Left(Text, SelStart) & Now & _
        Right(Text, Len(Text) - SelStart)
RichTextBox1.SelStart = SelStart

End Sub

Private Sub mnuu_Click()
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "<u></u>"
End Sub

Private Sub mnuundo_Click()

Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtxtEdit.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub mnuprint_Click()
Dim bcancel As Boolean
Dim ncopy As Integer
On Error GoTo errorhandler

bcancel = False

CDL1.Flags = cdlPDHidePrintToFile Or _
        cdlPDNoSelection Or cdlPDNoPageNums _
        Or cdlPDCollate

CDL1.CancelError = True
CDL1.PrinterDefault = True
CDL1.Copies = 1
CDL1.ShowPrinter

If bcancel = False Then
    PrintRTF RichTextBox1, 1440, 1440, 1440, 1440
    For ncopy = 1 To CDL1.Copies
    Next ncopy
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
bcancel = True
Resume Next
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim X%
X% = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit program")
If X% = vbNo Then
Cancel = 1
Exit Sub

End If
End Sub
