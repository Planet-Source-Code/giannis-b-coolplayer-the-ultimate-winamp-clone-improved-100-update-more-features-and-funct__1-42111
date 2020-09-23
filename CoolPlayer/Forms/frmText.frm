VERSION 5.00
Begin VB.Form frmTxt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text search"
   ClientHeight    =   1215
   ClientLeft      =   5280
   ClientTop       =   3285
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFing 
      Caption         =   "&Search"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Search for it"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Close the dialog"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Caption         =   "Search..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Your word"
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblStat 
         Caption         =   "Enter a phrase or a word:"
         Height          =   255
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   "Enter a phrase or a word"
         Top             =   240
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()

    Call DisableForms(True)
    Unload frmTxt

End Sub
Private Sub cmdFing_Click()

    Call Lst.findString(txtSearch.Text, frmPl.l)
    Call SetMax

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call DisableForms(True)
End Sub
Private Sub txtSearch_Change()
    Call TextChange(txtSearch.Text)
End Sub
Private Sub txtSearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      Call Lst.findString(txtSearch.Text, frmPl.l)
      Call SetMax
    End If

End Sub
