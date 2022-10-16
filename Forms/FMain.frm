VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "FMain"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "ComplicatedOperation"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"FMain.frx":0000
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3975
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

