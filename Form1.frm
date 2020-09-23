VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Wizzard Works 2001"
   ClientHeight    =   9615
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9735
   FillColor       =   &H00C0C0C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.Image Image1 
         Height          =   135
         Left            =   5760
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.ComboBox Sizes 
      Height          =   315
      ItemData        =   "Form1.frx":212A
      Left            =   2460
      List            =   "Form1.frx":2155
      TabIndex        =   2
      Text            =   "Sizes"
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":218B
      Left            =   90
      List            =   "Form1.frx":219B
      TabIndex        =   1
      Text            =   "Font"
      Top             =   360
      Width           =   2325
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   8925
      Left            =   30
      TabIndex        =   0
      Top             =   705
      Visible         =   0   'False
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   15743
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":21CD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog SaveDlg 
      Left            =   7710
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog OpenDlg 
      Left            =   8190
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   7230
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6735
      Top             =   -150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7500
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":229A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23AC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24BE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25D0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26E2
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27F4
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2906
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A18
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B2A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C3C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D4E
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E60
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F72
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3084
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save &As"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnufont 
      Caption         =   "Font"
      Begin VB.Menu mnubold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "Italics"
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "Underline"
      End
      Begin VB.Menu mnusize 
         Caption         =   "Size"
         Begin VB.Menu mnu12 
            Caption         =   "12"
         End
         Begin VB.Menu mnu24 
            Caption         =   "24"
         End
         Begin VB.Menu mnu36 
            Caption         =   "36"
         End
      End
      Begin VB.Menu mnufontbox 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuinsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuimage 
         Caption         =   "Image"
      End
      Begin VB.Menu mnusep 
         Caption         =   "Seperator"
      End
      Begin VB.Menu mnutimedate 
         Caption         =   "Time/Date"
      End
   End
   Begin VB.Menu mnubuilt 
      Caption         =   "&Built In"
      Begin VB.Menu mnuhtml 
         Caption         =   "&HTML editor"
      End
      Begin VB.Menu mnubrowser 
         Caption         =   "&Web Browser"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close1_Click()
Frame1.Visible = False
End Sub

Private Sub close2_Click()
Frame3.Visible = False
End Sub

Private Sub Fonts_Click()
ActiveForm.rtfText.SelFontName = Fonts.Text
End Sub

Private Sub Combo1_Change()
RichTextBox1.SelFontName = Fonts.Text
End Sub

Private Sub Form_Resize()
RichTextBox1.Width = Form1.ScaleWidth
RichTextBox1.Height = Form1.ScaleHeight
End Sub

Private Sub mnu12_Click()
RichTextBox1.Font.Size = 12
End Sub

Private Sub mnu24_Click()
RichTextBox1.Font.Size = 24
End Sub
Private Sub mnu36_click()
RichTextBox1.Font.Size = 36
End Sub
Private Sub mnubold_Click()
If RichTextBox1.SelBold Then
    RichTextBox1.SelBold = False
    mnubold.Checked = False
Else
  RichTextBox1.SelBold = True
    mnubold.Checked = True
End If
End Sub

Private Sub mnubrowser_Click()
frmBrowser.Visible = True
End Sub

Private Sub mnuclose_Click()
RichTextBox1.Text = ""
RichTextBox1.Visible = False
mnuprint.Enabled = False
mnusaveas.Enabled = False
End Sub

Private Sub mnucopy_Click()
On Error Resume Next
    Clipboard.SetText RichTextBox1.SelRTF
    mnupaste.Enabled = True
End Sub

Private Sub mnucut_Click()
     On Error Resume Next
    Clipboard.SetText RichTextBox1.SelRTF
    RichTextBox1.SelText = vbNullString
    mnupaste.Enabled = True
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnufontbox_Click()
    ' Shows the Font dialogue box and sets the current font.
CDL1.Flags = cdlCFBoth Or cdlCFEffects
CDL1.ShowFont

With RichTextBox1
    .SelFontName = CDL1.FontName
    .SelFontSize = CDL1.FontSize
    .SelBold = CDL1.FontBold
    .SelItalic = CDL1.FontItalic
    .SelStrikeThru = CDL1.FontStrikethru
    .SelUnderline = CDL1.FontUnderline
    .SelColor = CDL1.Color
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim X%
X% = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit program")
If X% = vbNo Then
Cancel = 1
Exit Sub

End If
End Sub

Private Sub mnuhtml_Click()
Form2.Visible = True
End Sub

Private Sub mnuimage_Click()
On Error Resume Next
frmImageOpen.Show vbModal, Me
End Sub

Private Sub mnuitalic_Click()
If RichTextBox1.SelItalic Then
    RichTextBox1.SelItalic = False
    mnuItaclic.Checked = False
Else
  RichTextBox1.SelItalic = True
    mnuItalic.Checked = True
End If
End Sub

Private Sub mnunew_Click()
RichTextBox1.Text = ""
RichTextBox1.Visible = True
mnuprint.Enabled = True
mnusaveas.Enabled = True
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
OpenDlg.Filter = "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|Html Files (*.html)|*.html|All Files (*.*)|*.*"
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
mnuprint.Enabled = True
mnusaveas.Enabled = True
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
Private Sub mnusaveas_Click()
On Error GoTo nada
SaveDlg.CancelError = False
SaveDlg.Filter = "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
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
Private Sub mnusecond_Click()
Frame2.Visible = True
End Sub

Private Sub mnusep_Click()
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "______________________________________"
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

Private Sub mnuUnderline_Click()
If RichTextBox1.SelUnderline Then
    RichTextBox1.SelUnderline = False
    mnuUnderline.Checked = False
Else
  RichTextBox1.SelUnderline = True
    mnuUnderline.Checked = True
End If
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

Private Sub Sizes_Click()
RichTextBox1.SelFontSize = Sizes.Text
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnunew_Click
        Case "Open"
            mnuopen_Click
        Case "Save"
            mnusaveas_Click
        Case "Print"
            mnuprint_Click
        Case "Cut"
            mnucut_Click
        Case "Copy"
            mnucopy_Click
        Case "Paste"
            mnupaste_Click
        Case "Bold"
            mnubold_Click
        Case "Italic"
            mnuitalic_Click
        Case "Underline"
            mnuUnderline_Click
        Case "Align Left"
            RichTextBox1.SelAlignment = rtfLeft
        Case "Center"
            RichTextBox1.SelAlignment = rtfCenter
        Case "Align Right"
            RichTextBox1.SelAlignment = rtfRight
    End Select
End Sub
