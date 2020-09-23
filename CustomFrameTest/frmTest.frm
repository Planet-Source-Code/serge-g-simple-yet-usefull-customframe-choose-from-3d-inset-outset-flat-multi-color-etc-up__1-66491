VERSION 5.00
Object = "*\ACustomFrameControl.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin Custom_Frame.CustomFrame CustomFrame2 
      Height          =   3255
      Left            =   1440
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5741
      FrameBackType   =   0
   End
   Begin Custom_Frame.CustomFrame CustomFrame1 
      Height          =   1095
      Left            =   480
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1931
      FrameCaption    =   "Now, Transparent..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CustomFrame2_Click()

End Sub

