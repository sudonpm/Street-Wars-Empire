VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Street Wars: Empire"
   ClientHeight    =   7950
   ClientLeft      =   3405
   ClientTop       =   1275
   ClientWidth     =   11025
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPlaying 
      Left            =   120
      Top             =   7560
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      Picture         =   "frmMain.frx":014A
      ScaleHeight     =   915
      ScaleWidth      =   8475
      TabIndex        =   15
      Top             =   0
      Width           =   8535
   End
   Begin VB.CommandButton cmdSkills 
      Caption         =   "Skills"
      Height          =   315
      Left            =   1320
      TabIndex        =   14
      Top             =   7560
      Width           =   1140
   End
   Begin VB.TextBox txtChat 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      HideSelection   =   0   'False
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3120
      Width           =   8295
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Left            =   7800
      Top             =   7560
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   600
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstInventory 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   8640
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdTravel 
      Caption         =   "Travel"
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   7560
      Width           =   1140
   End
   Begin VB.CommandButton cmdPawnShop 
      Caption         =   "Pawn Shop"
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   7560
      Width           =   1140
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Map"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   7560
      Width           =   1140
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   240
      MaxLength       =   200
      TabIndex        =   0
      Top             =   7200
      Width           =   8310
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      HideSelection   =   0   'False
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4320
      Width           =   8310
   End
   Begin VB.TextBox txtNews 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      HideSelection   =   0   'False
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2566
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmMain.frx":9734
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Todays Stats"
      TabPicture(1)   =   "frmMain.frx":9750
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line4"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -75000
         TabIndex        =   39
         Top             =   360
         Width           =   8295
         Begin VB.Label lblDrugsS 
            BackStyle       =   0  'Transparent
            Caption         =   "Drugs Sold:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lblDrugsST 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1320
            TabIndex        =   50
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblDrugsB 
            BackStyle       =   0  'Transparent
            Caption         =   "Drugs Bought:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblDrugsBT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblDrugsU 
            BackStyle       =   0  'Transparent
            Caption         =   "Drugs Used:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblDrugsUT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1320
            TabIndex        =   46
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblPlaying 
            BackStyle       =   0  'Transparent
            Caption         =   "Playing Time:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   45
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblPlayingT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3960
            TabIndex        =   44
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblKilled 
            BackStyle       =   0  'Transparent
            Caption         =   "Times Killed:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblKilledT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3960
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblKillzT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3960
            TabIndex        =   41
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblKillz 
            BackStyle       =   0  'Transparent
            Caption         =   "Kills:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   40
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   8295
         Begin VB.Label lblAmmo 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   6720
            TabIndex        =   38
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblArmor 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   6720
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblAmmoDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ammo:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5640
            TabIndex        =   36
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblArmorDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Armor:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5640
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblWeapon 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   6720
            TabIndex        =   34
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblWeaponDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Weapon:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5640
            TabIndex        =   33
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblKills 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3720
            TabIndex        =   32
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblRank 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3720
            TabIndex        =   31
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblLocation 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblHomeTown 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   3720
            TabIndex        =   29
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblKillsDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Kills:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   28
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblRankDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Rank:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblLocationDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Location:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   26
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblHomeTownDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Town:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblBank 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblCash 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1200
            TabIndex        =   23
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblHealth 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1200
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblName 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "----"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblBankDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblCashDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cash:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblHealthDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Health:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblNameDisplay 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   7
         X1              =   -75000
         X2              =   -66720
         Y1              =   340
         Y2              =   340
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   7
         X1              =   0
         X2              =   8280
         Y1              =   345
         Y2              =   345
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inventory"
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
      Left            =   8640
      TabIndex        =   55
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Players"
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
      Left            =   8640
      TabIndex        =   54
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lbllastkilled 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Killed Location"
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
      Left            =   8760
      TabIndex        =   53
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lbllastkilledT 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   52
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   416
      X2              =   568
      Y1              =   512
      Y2              =   512
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   168
      X2              =   16
      Y1              =   512
      Y2              =   512
   End
   Begin VB.Label lblLastSell 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   11
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label lblLastSellDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Sell Location"
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
      Left            =   8760
      TabIndex        =   10
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label lblLastBuy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lblLastBuyDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Buy Location"
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
      Left            =   8760
      TabIndex        =   8
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Shape shpNavigation 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   2340
      Left            =   8640
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Menu mnuFileConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuFileDisconnect 
      Caption         =   "&Disconnect"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGuide 
         Caption         =   "Street Wars: Empire Help Guide"
      End
      Begin VB.Menu mnuHelpVisitSite 
         Caption         =   "Street Wars: Empire Website"
      End
      Begin VB.Menu mnuHelpForums 
         Caption         =   "Street Wars: Empire Forums"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutCheating 
         Caption         =   "Cheating"
      End
      Begin VB.Menu mnuAboutGame 
         Caption         =   "Game"
      End
   End
   Begin VB.Menu mnuFileExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuInventoryEquip 
         Caption         =   "Equip"
      End
      Begin VB.Menu mnuInventoryUnequip 
         Caption         =   "Un-Equip"
      End
      Begin VB.Menu mnuInventoryExamine 
         Caption         =   "Examine"
      End
      Begin VB.Menu mnuInventoryUse 
         Caption         =   "Use"
      End
      Begin VB.Menu mnuInventoryDrop 
         Caption         =   "Drop"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Users"
      Visible         =   0   'False
      Begin VB.Menu mnuUsersMsg 
         Caption         =   "Msg"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdMap_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(253) & Chr$(5) & Chr$(0)
DoEvents

End Sub
Private Sub cmdSkills_Click()
On Error Resume Next

frmMain.wsk.SendData "skills" & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdPawnShop_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(254) & Chr$(2) & Chr$(0)
DoEvents

End Sub
Private Sub cmdTravel_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(255) & Chr$(6) & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If MoveDelay = True Then
   Exit Sub
End If

If KeyUsed = False Then
If KeyCode = vbKeyUp Then
   frmMain.wsk.SendData "n" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyRight Then
   frmMain.wsk.SendData "e" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyDown Then
   frmMain.wsk.SendData "s" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyLeft Then
   frmMain.wsk.SendData "w" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF1 Then
   frmMain.wsk.SendData "punch" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF2 Then
   frmMain.wsk.SendData "strike" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF3 Then
   frmMain.wsk.SendData "fire" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF4 Then
   frmMain.wsk.SendData "look" & Chr$(0)
   DoEvents
   KeyUsed = True
End If
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
   KeyUsed = False
ElseIf KeyCode = vbKeyRight Then
   KeyUsed = False
ElseIf KeyCode = vbKeyDown Then
   KeyUsed = False
ElseIf KeyCode = vbKeyLeft Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF1 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF2 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF3 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF4 Then
   KeyUsed = False
End If

End Sub
Private Sub Form_Load()
Dim a As Integer 'Counter

'Setup initial inventory slots
For a = 0 To 19
  lstInventory.AddItem "<Empty>"
Next a

txtNews.BackColor = vbBlack
txtNews.ForeColor = vbWhite

End Sub
Private Sub imgEast_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "e" & Chr$(0)
   DoEvents
End If

End Sub

Private Sub imgNorth_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "n" & Chr$(0)
   DoEvents
End If


End Sub

Private Sub imgSouth_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "s" & Chr$(0)
   DoEvents
End If

End Sub

Private Sub imgWest_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "w" & Chr$(0)
   DoEvents
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

wsk.Close

End Sub







Private Sub lstInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuInventory
End If

End Sub


Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuUsers
End If

End Sub

Private Sub mnuAboutCheating_Click()
frmCheaters.Show
End Sub
Private Sub mnuAboutGame_Click()
frmGame.Show
End Sub

Private Sub mnuFileConnect_Click()
On Error Resume Next
Dim iServ As String


iServ = "empireonline.afraid.org"

'Disable menus
frmMain.mnuFileConnect.Enabled = False
frmMain.mnuFileExit.Enabled = False
frmMain.mnuFileDisconnect.Visible = True
frmMain.mnuFileDisconnect.Enabled = True
frmMain.mnuFileConnect.Visible = False

'Track anti hack
PuseT = 0

'Connect to the server
With wsk
  .Close
  .Protocol = sckTCPProtocol
  .RemotePort = ServerPort
  .RemoteHost = iServ
  .Connect
End With

        frmMain.tmrPlaying.Interval = 60000
        frmMain.lblPlayingT.Caption = 0
        frmMain.tmrPlaying.Enabled = True

Call ShowText("Connecting to the Street Wars: Empire server, please stand by..." & vbCrLf & vbCrLf)

End Sub
Private Sub mnuFileDisconnect_Click()
'Disconnect and enable menus

wsk.Close
frmMain.mnuFileDisconnect.Visible = False
frmMain.mnuFileDisconnect.Enabled = False
frmMain.mnuFileConnect.Visible = True
frmMain.mnuFileConnect.Enabled = True
frmMain.mnuFileExit.Enabled = True

frmMain.tmrPlaying.Enabled = False

End Sub

Private Sub mnuFileExit_Click()

   'Close winsock and shut down the game
   wsk.Close
   Unload Me
   End

End Sub



Private Sub mnuHelpForums_Click()

Call OpenLocation("http://swe.dreamersway.com/forums", SW_SHOWNORMAL)

End Sub

Private Sub mnuHelpGuide_Click()

Call OpenLocation("http://swe.dreamersway.com/helpguide.html", SW_SHOWNORMAL)

End Sub

Private Sub mnuHelpVisitSite_Click()

Call OpenLocation("http://swe.dreamersway.com", SW_SHOWNORMAL)

End Sub
Private Sub mnuInventoryDrop_Click()

frmMain.wsk.SendData Chr$(7) & lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub
Private Sub mnuUsersTrack_Click()

frmMain.wsk.SendData "track " & Trim$(Mid$(lstUsers.Text, 7)) & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub
Private Sub mnuUsersmsg_Click()

frmMain.txtInput.Text = "'" & Trim$(Mid$(lstUsers.Text, 7)) & Chr$(0)
DoEvents

End Sub
Private Sub mnuInventoryEquip_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(3) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryExamine_Click()

'Examine the item
frmMain.wsk.SendData Chr$(255) & Chr$(2) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryUnequip_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(4) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryUse_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(5) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub



Private Sub tmrPlaying_Timer()
frmMain.lblPlayingT.Caption = frmMain.lblPlayingT.Caption + 1
DoEvents
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
On Error GoTo Failed 'Error Handler

'Track anti hack
With frmMain.txtInput
  Dim Search, Where
    Search = "track"
    ' Find string in text.
    Where = InStr(frmMain.txtInput.Text, Search)
    If Where Then
        PuseT = 1
    End If
End With

'Send textbox text to the server
If (KeyAscii = 13) And (txtInput.Text <> "") Then
  KeyAscii = 0
  If wsk.State <> sckClosed Then
     If InputDelay = True Then
        Exit Sub
     End If
    wsk.SendData Trim$(txtInput.Text) & Chr$(0)
    DoEvents
    txtInput.Text = ""
  End If
End If
Exit Sub

'If an error occurs,  close the socket and reset
'everything
Failed:
wsk.Close
With txtOutput
  .Text = .Text & "An error has occured while sending data to the server, your connection has been reset." & vbCrLf & vbCrLf
  .SelStart = Len(.Text)
End With
txtInput.Text = ""
tmrMain.Enabled = False
mnuFileConnect.Enabled = True
mnuFileExit.Enabled = True

End Sub
Private Sub txtNews_GotFocus()
  'Don't allow textbox to have focus
  txtInput.SetFocus
End Sub


Private Sub txtOutput_GotFocus()
  'Don't allow textbox to get focus
  txtInput.SetFocus
End Sub
Private Sub txtChat_GotFocus()
  'Don't allow textbox to get focus
  txtInput.SetFocus
End Sub


Private Sub wsk_Connect()

frmMain.wsk.SendData ClientVer & Chr$(0)
DoEvents

End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
Dim a As Integer 'Counter
Dim Msg As String 'String to hold data off the wire
Dim SplitMsg() As String 'String array to parse data

'Pull data off the wire
wsk.GetData Msg, vbString

'Split the string array
SplitMsg = Split(Msg, Chr$(0))

'Loop through data and process accordingly
For a = 0 To UBound(SplitMsg) - 1
   
   Select Case Left$(SplitMsg(a), 2)
      Case Chr$(255) & Chr$(2)
         Call TravelMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(3)
         Call PawnShopMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(4)
         Call UpdateCashRank(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(5)
         Call PawnShopItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(6)
         Call PawnShopPlayerInventoryUpdate(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(7)
         Call UpdateGeneralInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(2)
         Call UpdateGearInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(3)
         Call UpdatePlayerList(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(4)
         Call BuyDrugMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(5)
         Call CloseDrugDealMenu
      Case Chr$(254) & Chr$(6)
         Call DrugDealItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(7)
         Call DrugDealMessage(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(2)
         Call UpdateDealerInventory(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(3)
         Call SellDrugMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(4)
         Call CloseDruggieMenu
      Case Chr$(253) & Chr$(5)
         Call DruggieMenuMessage(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(6)
         Call DruggieMenuItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(7)
         Call ReUpdateDruggieInventory(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(2)
         Call ShowMap(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(3)
         Call UpdateNews(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(4)
         frmMain.lblLastBuy.Caption = Mid$(SplitMsg(a), 3)
      Case Chr$(252) & Chr$(5)
         frmMain.lblLastSell.Caption = Mid$(SplitMsg(a), 3)
      Case Chr$(252) & Chr$(231)
         frmMain.lbllastkilledT.Caption = Mid$(SplitMsg(a), 3)
   End Select
   
   Select Case Left$(SplitMsg(a), 1)
      Case Chr$(2)
         Call ShowText(Mid$(SplitMsg(a), 2))
      Case Chr$(123)
         Call ShowChat(Mid$(SplitMsg(a), 2))
      Case Chr$(3)
         Call NewAccount
      Case Chr$(4)
         Call DupeName
      Case Chr$(5)
         Call AccountCreated
      Case Chr$(6)
         Call UpdateFullInventory(Mid$(SplitMsg(a), 2))
      Case Chr$(7)
         Call UpdateSingleItem(Mid$(SplitMsg(a), 2))
   End Select
Next a

End Sub
