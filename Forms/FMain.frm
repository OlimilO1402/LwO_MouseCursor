VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "FMain"
   ScaleHeight     =   3195
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "ComplicatedOperation"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   $"FMain.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Don't worry, the various mouse cursors here are intentional and will be restored after closing the program, I promise!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private mc As TMouseCursor

Private Sub Form_Load()
    
    Me.Caption = "MouseCursor v" & App.Major & "." & App.Minor & "." & App.Revision
    New_MouseCursor(mc, Me) = EMouseCursor.MCCrosshair
    
End Sub

Private Sub Command1_Click()

    Dim mc As TMouseCursor: New_MouseCursor(mc, Me) = EMouseCursor.MCArrowHourglass
    Sleep 500
    
    ReadAHugeFile
    Sleep 500
    
    ReadAHugeFile
    Sleep 500
    
End Sub

Public Sub ReadAHugeFile()
    
    Dim mc As TMouseCursor: New_MouseCursor(mc, Me) = EMouseCursor.MCHourglass
    Sleep 2000
    
End Sub

