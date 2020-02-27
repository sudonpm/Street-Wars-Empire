VERSION 5.00
Begin VB.Form frmPawnShop 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Pawn Shop"
   ClientHeight    =   4830
   ClientLeft      =   195
   ClientTop       =   360
   ClientWidth     =   7755
   Icon            =   "frmPawnShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstShop 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   1740
      Width           =   2055
   End
   Begin VB.ListBox lstInv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   5400
      TabIndex        =   3
      Top             =   1740
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>> Buy >>>"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSell 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<< Sell <<<"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
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
      Left            =   2280
      TabIndex        =   0
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1320
      Left            =   0
      Picture         =   "frmPawnShop.frx":0442
      Top             =   0
      Width           =   7620
   End
   Begin VB.Label lblShop 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shop Inventory"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Inventory"
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
      Left            =   5400
      TabIndex        =   16
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome to the Pawn Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblPriceDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Item"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblCanBuyDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Can Buy:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblCanBuy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Item"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   152
      X2              =   352
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Label lblCashDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash:"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblRankDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rank:"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblRank 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblItemDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Item"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "frmPawnShop"
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

Private Sub cmdBuy_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdExit.SetFocus
frmMain.wsk.SendData Chr$(254) & Chr$(5) & lstShop.ListIndex & Chr$(0)
DoEvents

End Sub
Private Sub cmdExit_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub


Private Sub cmdSell_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False
cmdExit.SetFocus
frmMain.wsk.SendData Chr$(254) & Chr$(6) & lstInv.ListIndex & Chr$(0)
DoEvents

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


Private Sub lstInv_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False

End Sub

Private Sub lstInv_DblClick()
   
If lstInv.Text <> "<Empty>" Then
   cmdSell.Enabled = True
   cmdBuy.Enabled = False
   frmMain.wsk.SendData Chr$(254) & Chr$(3) & lstInv.ListIndex & Chr$(0)
   DoEvents
ElseIf lstInv.Text = "<Empty>" Then
   cmdSell.Enabled = False
   cmdBuy.Enabled = False
End If

End Sub

Private Sub lstShop_Click()

cmdBuy.Enabled = False
cmdSell.Enabled = False

End Sub

Private Sub lstShop_DblClick()
   
   cmdBuy.Enabled = True
   cmdSell.Enabled = False
   frmMain.wsk.SendData Chr$(254) & Chr$(4) & lstShop.ListIndex & Chr$(0)
   DoEvents

End Sub


