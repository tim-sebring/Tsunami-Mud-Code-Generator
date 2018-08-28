VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CodeGen"
   ClientHeight    =   4065
   ClientLeft      =   8835
   ClientTop       =   6210
   ClientWidth     =   2040
   Icon            =   "CodeGen1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   2040
   Begin VB.CommandButton Command5 
      Caption         =   "Create Monster"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create Room"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Weapon"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Armour"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Object"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu File_Menu 
      Caption         =   "&File"
      Begin VB.Menu File_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

frmCreateObject.Show 1

End Sub

Private Sub Command2_Click()

frmCreateArmour.Show 1

End Sub

Private Sub Command3_Click()

frmCreateWeapon.Show 1


End Sub

Private Sub Command4_Click()
' Create room

frmCreateRoom.Show 1

End Sub

Private Sub Command5_Click()
' Create monster
frmCreateMonster.Show 1


End Sub

Private Sub File_Exit_Click()

End

End Sub

Private Sub Form_Load()
On Error GoTo Errorhandler

If Right(App.Path, 1) = "\" Then
    MkDir (App.Path & "Code")
Else
    MkDir (App.Path & "\Code")
End If
Exit Sub
 
Errorhandler:
DoEvents


End Sub
