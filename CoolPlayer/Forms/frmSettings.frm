VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options and preferences..."
   ClientHeight    =   2895
   ClientLeft      =   3645
   ClientTop       =   2565
   ClientWidth     =   6510
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGen 
      Caption         =   "Plugins"
      Height          =   1815
      Index           =   4
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Plugins"
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ListBox lstPlugs 
         Height          =   1035
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdRegPlug 
         Caption         =   "&Reg plugin"
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         ToolTipText     =   "Click to register the Misc_v1.dll"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdUnreg 
         Caption         =   "&Unreg plugin"
         Height          =   255
         Left            =   4560
         TabIndex        =   33
         ToolTipText     =   "Click to unregister the Misc_v1.dll"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlugAb 
         Caption         =   "&About"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblPlug 
         Caption         =   "Note: If you unregister the plugin, the program will register it on startup, if find it. Else the program will terminate."
         Height          =   420
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   5775
      End
   End
   Begin VB.Frame fraGen 
      Caption         =   "Registry"
      Height          =   1815
      Index           =   3
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Registry"
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.ListBox lstTypes 
         Height          =   1035
         ItemData        =   "frmSettings.frx":0442
         Left            =   120
         List            =   "frmSettings.frx":045B
         MultiSelect     =   1  'Simple
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "A&ssociate"
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         ToolTipText     =   "Click to associate the selected types with the CoolPlayer"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "&All"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         ToolTipText     =   "Select all items"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox chkAss 
         Caption         =   "Associate all files on start"
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "&None"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblAdv 
         Caption         =   "Note: Associate the selected file types with CoolPlayer."
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   1080
         Width           =   4200
      End
   End
   Begin MSComctlLib.ImageList Iml 
      Left            =   120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0483
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0D5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":11AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":14C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":17E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1AFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":27D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":34B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3A4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3D65
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraGen 
      Caption         =   "Graphics"
      Height          =   1815
      Index           =   2
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Graphics"
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin MSComctlLib.Slider sliT 
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Max             =   250
         TickFrequency   =   25
      End
      Begin VB.CheckBox chkTop 
         Caption         =   "Always on top"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chkGraph 
         Caption         =   "Graph effect on exit"
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkSnap 
         Caption         =   "Snap windows to screen edges"
         Height          =   195
         Left            =   3120
         TabIndex        =   18
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkScroll 
         Caption         =   "Scroll track name"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblTInfo 
         Caption         =   "(For Win2K/WinXP only.)"
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lblLev 
         Caption         =   "Transparency :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraGen 
      Caption         =   "Playlist"
      Height          =   1815
      Index           =   1
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Playlist"
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkLoop 
         Caption         =   "Loop the playlist"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkDubs 
         Caption         =   "Remove double entries"
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkRand 
         Caption         =   "Random play"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkSingl 
         Caption         =   "Single click on playlist"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Sort playlist on load"
         Height          =   195
         Left            =   3360
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "A&pply"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      ToolTipText     =   "Just save options"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sa&ve"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      ToolTipText     =   "Close the settings dialog and save options"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Close the settings dialog without saving options"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame fraGen 
      Caption         =   "General"
      Height          =   1815
      Index           =   0
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "General setings"
      Top             =   480
      Width           =   6015
      Begin VB.CheckBox chkStart 
         Caption         =   "Start on StartUp"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkMin 
         Caption         =   "Start minimized"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkInst 
         Caption         =   "Allow double instances"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkMute 
         Caption         =   "Mute"
         Height          =   190
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkSplash 
         Caption         =   "Show splash screen"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin MSComctlLib.Slider sliI 
         Height          =   1455
         Left            =   4440
         TabIndex        =   37
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2566
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Image TrayI 
         Height          =   495
         Left            =   5040
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblTray 
         Caption         =   "Tray icon:"
         Height          =   255
         Left            =   4920
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.TabStrip Tab 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4048
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Playlist"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gr&aphics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Registry"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P&lugins"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuReg 
      Caption         =   "Registry"
      Visible         =   0   'False
      Begin VB.Menu mnuRegExist 
         Caption         =   "Register existing"
      End
      Begin VB.Menu mnuENS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegNew 
         Caption         =   "Register &new"
      End
   End
   Begin VB.Menu mnuUReg 
      Caption         =   "Registry"
      Visible         =   0   'False
      Begin VB.Menu mnuUExist 
         Caption         =   "Unregister &existing"
      End
      Begin VB.Menu mnuUENS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUNew 
         Caption         =   "Unregister &new"
      End
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkDubs_Click()
    chkDubs.Value = IIf(chkDubs.Value = 1, 1, 0)
End Sub
Private Sub chkGraph_Click()
    chkGraph.Value = IIf(chkGraph.Value = 1, 1, 0)
End Sub
Private Sub chkInst_Click()
    chkInst.Value = IIf(chkInst.Value = 1, 1, 0)
End Sub
Private Sub chkLoop_Click()
    chkLoop.Value = IIf(chkLoop.Value = 1, 1, 0)
End Sub
Private Sub chkMute_Click()
    chkMute.Value = IIf(chkMute.Value = 1, 1, 0)
End Sub
Private Sub chkTop_Click()
    chkTop.Value = IIf(chkTop.Value = 1, 1, 0)
End Sub
Private Sub chkRand_Click()
    chkRand.Value = IIf(chkRand.Value = 1, 1, 0)
End Sub
Private Sub chkScroll_Click()
    chkScroll.Value = IIf(chkScroll.Value = 1, 1, 0)
End Sub
Private Sub chkSingl_Click()
    chkSingl.Value = IIf(chkSingl.Value = 1, 1, 0)
End Sub
Private Sub chkSnap_Click()
    chkSnap.Value = IIf(chkSnap.Value = 1, 1, 0)
End Sub
Private Sub chkSort_Click()
    chkSort.Value = IIf(chkSort.Value = 1, 1, 0)
End Sub
Private Sub chkSplash_Click()
    chkSplash.Value = IIf(chkSplash.Value = 1, 1, 0)
End Sub
Private Sub chkStart_Click()
    chkStart.Value = IIf(chkStart.Value = 1, 1, 0)
End Sub
Private Sub chkMin_Click()
    chkMin.Value = IIf(chkMin.Value = 1, 1, 0)
End Sub
Private Sub cmdAll_Click()
    Call SelList(True)
End Sub
Private Sub cmdApply_Click()
    Call SaveOptions
End Sub
Private Sub cmdCancel_Click()

    Call DisableForms(True)
    Unload frmSet

End Sub
Private Sub cmdNone_Click()
    Call SelList(False)
End Sub

Private Sub cmdPlugAb_Click()
    Call File.AboutPlugin(CI.bTop)
End Sub
Private Sub cmdReg_Click()

    On Error GoTo RError
    Dim intI As Integer

    For intI = 0 To lstTypes.ListCount - 1
     If lstTypes.Selected(intI) = True Then
      Call Reg.PublicReg(intI + 1, App.Path, App.EXEName)
     End If
    Next intI

RError:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub cmdRegPlug_Click()
    Call PopupMenu(mnuReg, True, cmdRegPlug.Left + 2080, cmdRegPlug.Top + 1470)
End Sub
Private Sub cmdSave_Click()

    On Error Resume Next
    Call SaveOptions
    Call DisableForms(True)
    Unload frmSet

End Sub

Private Sub cmdUnreg_Click()
    Call PopupMenu(mnuUReg, True, cmdUnreg.Left + 2230, cmdUnreg.Top + 1470)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call DisableForms(True)
End Sub

Private Sub mnuRegExist_Click()
    Call CreateKeys
End Sub

Private Sub mnuUExist_Click()
    Call DeleteKeys
End Sub
Private Sub sliI_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call SaveOptions
End Sub
Private Sub sliI_Scroll()

    With sliI
     CI.iIcon = .Value
     Call ChangeIcon
    End With

End Sub
Private Sub sliT_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call SaveOptions
End Sub
Private Sub Tab_Click()
    Call ShowTab(frmSet.Tab.SelectedItem.Index)
End Sub
