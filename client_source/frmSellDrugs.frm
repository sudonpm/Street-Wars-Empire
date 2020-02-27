VERSION 5.00
Begin VB.Form frmSellDrugs 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell Drugs"
   ClientHeight    =   5925
   ClientLeft      =   345
   ClientTop       =   390
   ClientWidth     =   5745
   Icon            =   "frmSellDrugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstInventory 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   3240
      TabIndex        =   2
      Top             =   2880
      Width           =   2205
   End
   Begin VB.CommandButton cmdSell 
      BackColor       =   &H00C0C0C0&
      Caption         =   "*** Sell ***"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   5460
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forget It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Label lblDrugDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drug"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblPriceDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
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
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblDrug 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      BorderWidth     =   2
      X1              =   16
      X2              =   208
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Label lblCashDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Alright,  let me see what you got."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Image imgMain 
      BorderStyle     =   1  'Fixed Single
      Height          =   2805
      Left            =   240
      Picture         =   "frmSellDrugs.frx":0442
      Top             =   0
      Width           =   5235
   End
End
Attribute VB_Name = "frmSellDrugs"
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

Private Sub cmdExit_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub


Private Sub cmdSell_Click()

If lstInventory.ListIndex < 0 Or _
   lstInventory.ListIndex > 19 Then
   Exit Sub
End If

frmMain.lblDrugsST.Caption = frmMain.lblDrugsST.Caption + 1
cmdSell.Enabled = False
frmMain.wsk.SendData Chr$(253) & Chr$(4) & lstInventory.ListIndex & Chr$(0)
DoEvents
cmdExit.SetFocus

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


Private Sub lstInventory_Click()

cmdSell.Enabled = False

End Sub


Private Sub lstInventory_DblClick()

If lstInventory.ListIndex < 0 Or _
   lstInventory.ListIndex > 19 Then
   Exit Sub
End If

If lstInventory.Text = "<Empty>" Then
   cmdSell.Enabled = False
ElseIf lstInventory.Text <> "<Empty>" Then
   cmdSell.Enabled = True
   frmMain.wsk.SendData Chr$(253) & Chr$(3) & lstInventory.ListIndex & Chr$(0)
   DoEvents
End If

End Sub


