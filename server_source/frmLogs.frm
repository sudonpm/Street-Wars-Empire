VERSION 5.00
Begin VB.Form frmLogs 
   BackColor       =   &H00000000&
   Caption         =   "Logs"
   ClientHeight    =   5430
   ClientLeft      =   660
   ClientTop       =   420
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   751
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ListBox list1 
      Height          =   4545
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   8295
   End
   Begin VB.ListBox lstIndex 
      Height          =   4545
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuMainShow 
         Caption         =   "Show"
      End
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'you just need to implemente a compare to your textbox the code works fine

'draw a drivelistbox named Drive1 , dirlistbox named Dir1 , listbox named list1 and button named Command1

'To use the FileSystemObject, you'll need to add MicroSoft Scripting Runtime
'to your Project References. Go through Project > References and check the
'Microsoft Scripting Runtime check box.


Private fso As New FileSystemObject
Private m_SearchRunning As Boolean
Private Sub SearchFolder(srcFol As String)

Dim fld As Folder, tFld As Folder, fil As File

Set fld = fso.GetFolder(srcFol)
If fld.Files.Count + fld.SubFolders.Count > 0 Then
For Each fil In fld.Files
list1.AddItem fso.BuildPath(fld.Path, fil.name)
Next
For Each tFld In fld.SubFolders
If tFld.Files.Count + tFld.SubFolders.Count > 0 Then
SearchFolder tFld.Path
End If
DoEvents
If m_SearchRunning = False Then
Exit Sub
End If
Next
End If

End Sub

Private Sub Command1_Click()
If Command1.Caption = "Stop" Then
Command1.Caption = "Search"
m_SearchRunning = False
Exit Sub
End If
m_SearchRunning = True
Command1.Caption = "Stop"
Label1.Caption = "Files in " & Dir1.Path
list1.Clear
SearchFolder (Dir1)
If list1.ListCount = 0 Then
list1.AddItem "None Found"
Else
Label1.Caption = "All files in " & Dir1.Path & " - Total: " & list1.ListCount
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Command1.Caption = "Search"

'Disable X (Close Button) on main form
Dim hMenu As Long
Dim menuItemCount As Long
hMenu = GetSystemMenu(Me.hwnd, 0)
If hMenu Then
menuItemCount = GetMenuItemCount(hMenu)
Call RemoveMenu(hMenu, menuItemCount - 1, _
MF_REMOVE Or MF_BYPOSITION)
Call RemoveMenu(hMenu, menuItemCount - 2, _
MF_REMOVE Or MF_BYPOSITION)
Call DrawMenuBar(Me.hwnd)
End If
' I try this code and it works fine, hope it helps
End Sub


Private Sub cmdExit_Click()

Unload Me
frmMain.txtInput.SetFocus

End Sub
Private Sub lstindex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuMain
End If

End Sub
Private Sub mnumainshow_Click()

frmLogs.list1.Clear


End Sub

