VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   Caption         =   "Wizzard Web 2001"
   ClientHeight    =   10020
   ClientLeft      =   3060
   ClientTop       =   3630
   ClientWidth     =   9480
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   585
      Width           =   9435
      ExtentX         =   16642
      ExtentY         =   16563
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8130
      Top             =   90
   End
   Begin VB.PictureBox picAddress 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   75
      ScaleHeight     =   555
      ScaleWidth      =   17040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -15
      Width           =   17040
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6435
         Picture         =   "frmBrowser.frx":1CFA
         ScaleHeight     =   375
         ScaleWidth      =   360
         TabIndex        =   9
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5985
         Picture         =   "frmBrowser.frx":20C3
         ScaleHeight     =   360
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   90
         Width           =   375
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5565
         Picture         =   "frmBrowser.frx":233D
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   90
         Width           =   375
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5175
         Picture         =   "frmBrowser.frx":26E2
         ScaleHeight     =   360
         ScaleWidth      =   300
         TabIndex        =   6
         Top             =   90
         Width           =   300
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   4815
         Picture         =   "frmBrowser.frx":2A69
         ScaleHeight     =   360
         ScaleWidth      =   300
         TabIndex        =   5
         Top             =   90
         Width           =   300
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   4455
         Picture         =   "frmBrowser.frx":2D17
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   4
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   4065
         Picture         =   "frmBrowser.frx":2F2C
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   3
         Top             =   90
         Width           =   360
      End
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Text            =   "Http://"
         Top             =   135
         Width           =   3795
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Begin VB.Menu mnunewwindow 
            Caption         =   "New &Window"
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnufav 
      Caption         =   "&Favorites"
      Begin VB.Menu mnufavlist 
         Caption         =   "Favorites &List"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    Form_Resize
    brwWebBrowser.GoHome
    cboAddress.Move 50
    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    brwWebBrowser.Width = Me.ScaleWidth
    brwWebBrowser.Height = Me.ScaleHeight
End Sub

Private Sub mnuprint_Click()
On Error GoTo ErrHandler
brwWebBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "Wizzard Web 2001"
End Sub

Private Sub Picture2_Click()
brwWebBrowser.GoBack
End Sub

Private Sub Picture3_Click()
brwWebBrowser.GoForward
End Sub

Private Sub Picture4_Click()
timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub Picture5_Click()
brwWebBrowser.Refresh
End Sub

Private Sub Picture6_Click()
brwWebBrowser.GoHome
End Sub

Private Sub Picture7_Click()
brwWebBrowser.GoSearch
End Sub

Private Sub Picture8_Click()
On Error GoTo ErrHandler
brwWebBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT: Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "Wizzard Web 2001"
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Please Stand By..."
    End If
End Sub
