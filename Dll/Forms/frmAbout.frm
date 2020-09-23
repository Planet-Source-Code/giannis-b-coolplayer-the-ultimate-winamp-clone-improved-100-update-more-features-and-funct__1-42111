VERSION 5.00
Begin VB.Form frmAb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Misc_v1.dll"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "Misc_v1.dll by John"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         ToolTipText     =   "Close the dialog"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "The date..."
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblAb 
         Caption         =   "This plugin contains all the runtime functions needed for CoolPlayer. You can use it in your programs. Enjoy it..."
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "About"
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmAb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
    Unload frmAb
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload frmAb
End Sub
