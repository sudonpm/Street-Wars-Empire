VERSION 5.00
Begin VB.Form frmTravel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airport Checkout"
   ClientHeight    =   4080
   ClientLeft      =   195
   ClientTop       =   360
   ClientWidth     =   3510
   Icon            =   "frmTravel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   234
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewYork 
      BackColor       =   &H00C0C0C0&
      Caption         =   "New York"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1650
   End
   Begin VB.CommandButton cmdLosAngeles 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Los Angeles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1650
   End
   Begin VB.CommandButton cmdHouston 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Houston"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   1650
   End
   Begin VB.CommandButton cmdMiami 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Miami"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   1650
   End
   Begin VB.CommandButton cmdChicago 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chicago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   1650
   End
   Begin VB.CommandButton cmdNewJersey 
      BackColor       =   &H00C0C0C0&
      Caption         =   "New Jersey"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   1650
   End
   Begin VB.CommandButton cmdForgetIt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forget It"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   3450
   End
   Begin VB.Line lneMain 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      X1              =   2
      X2              =   227
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Label lblNewYork 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   1650
   End
   Begin VB.Label lblLosAngeles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   1650
   End
   Begin VB.Label lblHouston 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label lblMiami 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label lblChicago 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1650
   End
   Begin VB.Label lblNewJersey 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   1650
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   870
      Index           =   1
      Left            =   0
      Picture         =   "frmTravel.frx":0442
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmTravel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'  Streetwars Online 2 Version 1.00
'  Copyright 2000 - B.Smith aka (Wuzzbent)
'  All Rights Reserved
'  wuzzbent@swbell.net
'
'  By using this source code, you agree to the following
'  terms and conditions.
'
'  You may use this source code for your own personal
'  pleasure and use.  You may freely distribute it along with
'  any modification(s) made to it.  You may NOT remove, modify,
'  or adjust this copyright information.  You may NOT attempt
'  to charge for the use of this software under any conditions.
'
'  Support Free Software....
'
'******************************************************
'   Street Wars Empire is a modified version of 
'   Streetwars Online 2 Version 1.00 by Wuzzbent
'   Coded by sudonpm

Option Explicit

Private Sub cmdChicago_Click()

'Fly to Chicago
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "chicago" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdForgetIt_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdHouston_Click()

'Fly to Houston
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "houston" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdLosAngeles_Click()

'Fly to Los Angeles
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "los angeles" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdMiami_Click()

'Fly to Miami
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "miami" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdNewJersey_Click()

'Fly to New Jersey
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "new jersey" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdNewYork_Click()

'Fly to New York
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "new york" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub
Private Sub Form_Load()

Dim hMenu As Long
Dim menuItemCount As Long
'Obtain the handle to the form's system menu
hMenu = GetSystemMenu(Me.hWnd, 0)
If hMenu Then
'Obtain the number of items in the menu
menuItemCount = GetMenuItemCount(hMenu)
'Remove the system menu Close menu item.
'The menu item is 0-based, so the last
'item on the menu is menuItemCount - 1
Call RemoveMenu(hMenu, menuItemCount - 1, _
MF_REMOVE Or MF_BYPOSITION)
'Remove the system menu separator line
Call RemoveMenu(hMenu, menuItemCount - 2, _
MF_REMOVE Or MF_BYPOSITION)
'Force a redraw of the menu. This
'refreshes the titlebar, dimming the X
Call DrawMenuBar(Me.hWnd)
End If

End Sub


