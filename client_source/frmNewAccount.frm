VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Account"
   ClientHeight    =   4890
   ClientLeft      =   345
   ClientTop       =   390
   ClientWidth     =   4080
   ControlBox      =   0   'False
   Icon            =   "frmNewAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1140
      Width           =   1695
   End
   Begin VB.TextBox txtPassOne 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1500
      Width           =   1695
   End
   Begin VB.TextBox txtPassTwo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1860
      Width           =   1695
   End
   Begin VB.ComboBox cboCity 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblPasswordConfirm 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblCity1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "HomeTown:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewAccount.frx":08CA
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label lblIPAddress 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP Address:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image imgMain 
      BorderStyle     =   1  'Fixed Single
      Height          =   870
      Index           =   1
      Left            =   120
      Picture         =   "frmNewAccount.frx":0956
      Top             =   120
      Width           =   3870
   End
End
Attribute VB_Name = "frmNewAccount"
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

Private Sub cmdCreate_Click()

'Make sure no info fields are blank
If txtName.Text = "" Or _
   txtPassOne.Text = "" Or _
   txtPassTwo.Text = "" Or _
   txtEmail.Text = "" Or _
   cboCity.Text = "" Then
     lblMessage.Caption = "You must complete all fields before you can continue."
     Exit Sub
End If
   
'Make sure name and password are at least four
'characters in lenght
If Len(Trim$(txtName.Text)) < 4 Or _
   Len(Trim$(txtPassOne.Text)) < 4 Or _
   Len(Trim$(txtPassTwo.Text)) < 4 Then
     lblMessage.Caption = "Your name and password must be four or more characters in length to continue."
     Exit Sub
End If

'Check passwords for a match
If txtPassOne.Text <> txtPassTwo.Text Then
   lblMessage.Caption = "Your passwords do not match."
   Exit Sub
End If

txtName.Enabled = False
txtPassOne.Enabled = False
txtPassTwo.Enabled = False
txtEmail.Enabled = False
cboCity.Enabled = False
cmdCreate.Enabled = False

frmMain.wsk.SendData Trim$(txtName.Text) & Chr$(1) & Trim$(txtPassOne) & Chr$(1) & Trim$(cboCity.Text) & Chr$(1) & Trim$(txtEmail.Text) & Chr$(1) & Chr$(0)
DoEvents

End Sub
Private Sub cmdExit_Click()

'unload the new account form and enable the main form
'and enable the disabled menus
frmMain.wsk.Close
Call ShowText("Your connection to the server has been reset." & vbCrLf & vbCrLf)
frmMain.mnuFileConnect.Enabled = True
frmMain.mnuFileExit.Enabled = True
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub Form_Load()

'Display users IP address
lblIP.Caption = frmMain.wsk.LocalIP

'Add citys to the combo box
cboCity.AddItem ("New York"), 0
cboCity.AddItem ("Los Angeles"), 1
cboCity.AddItem ("Chicago"), 2
cboCity.AddItem ("Houston"), 3
cboCity.AddItem ("Miami"), 4
cboCity.AddItem ("New Jersey"), 5

End Sub
Private Sub Image1_Click()

End Sub


