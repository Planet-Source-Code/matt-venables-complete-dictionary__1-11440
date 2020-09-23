VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "msn dictionary"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2640
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/12/2000"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7620
            Text            =   "welcome to the msn dictionary"
            TextSave        =   "welcome to the msn dictionary"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "loop up"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblseconddef 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label lblsecond 
      BackStyle       =   0  'Transparent
      Caption         =   "second definition:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label lblfirstdef 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblfirst 
      BackStyle       =   0  'Transparent
      Caption         =   "first definition:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   120
      Picture         =   "Form1.frx":0ECA
      Top             =   5040
      Width           =   2355
   End
   Begin VB.Image Image3 
      Height          =   4965
      Left            =   0
      Picture         =   "Form1.frx":1D44
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5820
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "Form1.frx":1FF3
      Top             =   0
      Width           =   1725
   End
   Begin VB.Menu mnu_1 
      Caption         =   "mnu_1"
      Visible         =   0   'False
      Begin VB.Menu show 
         Caption         =   "show dictionary"
      End
      Begin VB.Menu break 
         Caption         =   "-"
      End
      Begin VB.Menu end 
         Caption         =   "exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DefineWord(word As String)
On Error GoTo err:
Text1.Text = Trim(Text1.Text)
StatusBar1.Panels(2).Text = "contacting msn dictionary"
Inet1.URL = "http://dictionary.msn.com/find/entry.asp?search=" & Text1.Text
def = Inet1.OpenURL(Inet1.URL)
StatusBar1.Panels(2).Text = "opening dictionary page"
If InStr(1, def, "No matches found for") Then GoTo nomatch:
beginspot = InStr(1, def, "<div class='dictionary'>") + 24
EndSpot = InStr(beginspot, def, "Encarta® World English Dictionary")
spot1 = beginspot
spot2 = beginspot + 2
StatusBar1.Panels(2).Text = "getting definitions"
' Get 2 Defs
   spot1 = InStr(spot1, def, "-1")
   spot1 = InStr(spot1, def, "<b>") + 3
   spot2 = InStr(spot1, def, "</b>")
   lblfirst.Caption = Mid$(def, spot1, spot2 - spot1)
   If InStr(1, lblfirst.Caption, "<") Then GoTo nomatch:
   spot1 = spot2 + 1
   spot1 = InStr(spot1, def, "</font>") + 7
   spot2 = InStr(spot1, def, "<br />")
   If InStr(spot1, def, "<img") < InStr(spot1, def, "<br />") Then
        spot2 = InStr(spot1, def, "<img")
   End If
   If InStr(spot1, def, "<i>") < spot2 Then
        spot2 = InStr(spot1, def, "<i>") - 1
    End If
    
   lblfirstdef.Caption = Mid$(def, spot1, spot2 - spot1)
' Second Def
   spot1 = InStr(spot1, def, "-1")
   spot1 = InStr(spot1, def, "<b>") + 3
   spot2 = InStr(spot1, def, "</b>")
   lblsecond.Caption = Mid$(def, spot1, spot2 - spot1)
   spot1 = spot2 + 1
   spot1 = InStr(spot1, def, "</font>") + 7
   spot2 = InStr(spot1, def, "<br />")
   If InStr(spot1, def, "<img") < InStr(spot1, def, "<br />") Then
        spot2 = InStr(spot1, def, "<img")
   End If
      If InStr(spot1, def, "<i>") < spot2 Then
        spot2 = InStr(spot1, def, "<i>") - 1
    End If
    
   lblseconddef.Caption = Mid$(def, spot1, spot2 - spot1)
err:
    StatusBar1.Panels(2).Text = "welcome to the dictionary"
    For a = 1395 To 6375
        Form1.Height = a
    Next a
    
    Exit Sub
nomatch:
    lblfirst.Caption = "first definition"
    lblfirstdef.Caption = "No matches found for your word or word is sensored"
        StatusBar1.Panels(2).Text = "welcome to the msn dictionary"
    For a = 1395 To 6375
        Form1.Height = a
    Next a
    Exit Sub
    
End Sub

Private Sub Command1_Click()

Me.Height = 1395
If Trim(Text1.Text) = "" Then Exit Sub
lblfirst.Caption = "first definition"
lblfirstdef.Caption = ""
lblsecond.Caption = "second definition"
lblseconddef.Caption = ""

DefineWord Text1.Text

End Sub


Private Sub end_Click()
End

End Sub

Private Sub Form_Load()
lblfirst.Caption = "first definition"
lblfirstdef.Caption = ""
lblsecond.Caption = "second definition"
lblseconddef.Caption = ""
Me.Height = 1395
    Me.show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " Click Right Mouse Button " & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result As Long
    Dim msg As Long
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP '514 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.show
        Case WM_LBUTTONDBLCLK '515 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.show
        Case WM_RBUTTONUP '517 display popup menu
        Result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mnu_1
    End Select
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub show_Click()
Me.SetFocus

End Sub
