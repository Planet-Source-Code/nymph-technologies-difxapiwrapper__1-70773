VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DIFxAPI Demo"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Uninstall"
      Height          =   720
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PreIinstall"
      Height          =   720
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Install"
      Height          =   720
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim Retval As String
Dim DTool As New DIFxAPI_Wrapper

Retval = DTool.Install("C:\Test Project\PlugAndPlay\toastpkg.inf", Force, "test", "test", "test", "test", False)

MsgBox Retval, vbInformation, "DIFxAPI Demo"

End Sub

Private Sub Command2_Click()

Dim Retval As String
Dim DTool As New DIFxAPI_Wrapper

Retval = DTool.Preinstall("C:\Test Project\PlugAndPlay\toastpkg.inf", Normal)

MsgBox Retval, vbInformation, "DIFxAPI Demo"

End Sub

Private Sub Command3_Click()

Dim Retval As String
Dim DTool As New DIFxAPI_Wrapper

Retval = DTool.UnInstall("C:\Test Project\PlugAndPlay\toastpkg.inf", Uninstall_Force, "test", "test", "test", "test", False)

MsgBox Retval, vbInformation, "DIFxAPI Demo"

End Sub
