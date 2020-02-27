VERSION 5.00
Begin VB.Form frmCheaters 
   BackColor       =   &H00000000&
   Caption         =   "Cheaters"
   ClientHeight    =   3120
   ClientLeft      =   360
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "frmCheaters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Click Here"
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "All cheaters will be posted on the site."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmCheaters.frx":014A
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Cheating will not be tolerated. Such as Logging(Disconnect when in a fight). The use of bots. Multiple Accounts."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmCheaters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()
Call OpenLocation("http://swe.dreamersway.com/forums/forumdisplay.php?f=8", SW_SHOWNORMAL)

End Sub
