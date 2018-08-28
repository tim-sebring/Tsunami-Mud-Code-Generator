VERSION 5.00
Begin VB.Form frmCreateWeapon 
   Caption         =   "Create Weapon"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "frmCreateWeapon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Text            =   "GENERAL 0"
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Metal Item"
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Lockerable Item"
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Magical Item"
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   5040
      TabIndex        =   20
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   4560
      Width           =   6255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   4200
      Width           =   6255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   3840
      Width           =   6255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   3480
      Width           =   6255
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCreateWeapon.frx":0442
      Left            =   1800
      List            =   "frmCreateWeapon.frx":0444
      TabIndex        =   2
      Text            =   "sword"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Damage Type"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Set Long Line 4:"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Set Long Line 3:"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Set Long Line 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Set Long Line 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Weapon's Set_Short:"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Weapon's Weight:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Weapon's Value:"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Weapon's ID (opt)"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Weapon Aliases:"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Weapon's Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Weapon Type:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Weapon WC:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmCreateWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Me.Hide
Unload Me

End Sub

Private Sub Command2_Click()

' NOTE: In the future, add a part where the wielder will
'       either receive some kind of bonus, or
'       an extra damage amount.


' Create weapon
Dim Temp As String
Dim DamType As Integer
MousePointer = 11
Temp = (App.Path & "\Code\" & Text1.Text)

Open Temp For Output Access Write As #1

Print #1, "inherit  STD_WEAPON; "
Print #1,
Print #1, "reset (arg)"
Print #1, "{"
Print #1, "    ::reset( arg );"
Print #1, "    if( arg ) return;"
Print #1, "    set_name(""" & Text2.Text & """);"
If Len(Text6.Text) > 0 Then
    Print #1, "    set_id(""" & Text6.Text & """);"
End If
Print #1, "    set_kind(""" & Combo1.Text & """); "
If (Len(Text3.Text) > 0) Then
    Print #1, "    set_alias(""" & Text3.Text & """); "
End If
If (Len(Text4.Text) > 0) Then
    Print #1, "    set_alias(""" & Text4.Text & """); "
End If
If (Len(Text5.Text) > 0) Then
    Print #1, "    set_alias(""" & Text5.Text & """); "
End If

Print #1, "    set_short(""" & Text9.Text & """);"
Print #1, "    set_long("

If (Len(Text10.Text) > 0) Then
    Print #1, "    """ & Text10.Text & " "" "
End If
If (Len(Text11.Text) > 0) Then
    Print #1, "    """ & Text11.Text & " "" "
End If
If (Len(Text12.Text) > 0) Then
    Print #1, "    """ & Text12.Text & " "" "
End If
If (Len(Text13.Text) > 0) Then
    Print #1, "    """ & Text13.Text & " "" "
End If

Print #1, "    );"
Print #1, "    set_class(" & Text14.Text & "); "

Print #1, "    set_value(" & Text7.Text & ");"
Print #1, "    set_weight(" & Text8.Text & ");"
If (Check3.Value = 1) Then
    Print #1, "    set_metal(1);"
Else
    Print #1, "    set_metal(0);"
End If
' Also need to add a set_type function which
' determines which kind of damage messages the
' weapon will give. (int)
Select Case Combo2.Text
Case "GENERAL 0"
    DamType = 0
Case "SLASH 1"
    DamType = 1
Case "PIERCE 2"
    DamType = 2
Case "BLUNT 3"
    DamType = 3
Case "UNARMED 4"
    DamType = 4
Case "MAGIC 5"
    DamType = 5
Case "FIRE 6"
    DamType = 6
Case "SHOCK 7"
    DamType = 7
Case "ACID 8"
    DamType = 8
Case "COLD 9"
    DamType = 9
Case "POISON 10"
    DamType = 10
Case "VORPAL 11"
    DamType = 11
Case "CLAW 12"
    DamType = 12
Case "BITE 13"
    DamType = 13
Case "PSYCHIC 14"
    DamType = 14
Case "STINGER 15"
    DamType = 15
Case "DRAIN 16"
    DamType = 16
Case "HOLY 17"
    DamType = 17
Case "SOUND 18"
    DamType = 18
Case "PETRIFY 21"
    DamType = 21
Case "BRAWLING 22"
    DamType = 22
Case "CHAOS 23"
    DamType = 23
Case "WARP 24"
    DamType = 24
Case "CLEAVE 25"
    DamType = 25
Case "BOW 26"
    DamType = 26
Case Else
    MsgBox ("Invalid Damage Type. Please reselect.")
    Exit Sub
End Select
Print #1, "    set_type(" & DamType & "); "


Print #1,
Print #1, "}"

If (Check1.Value = 1) Then
    Print #1, "query_magic() { return 1; }"
Else
    Print #1, "query_magic() { return 0; }"
End If

If (Check2.Value = 1) Then
    Print #1, "query_no_lock() { return 0;}"
Else
    Print #1, "query_no_lock() { return 1;}"
End If


Close (1)
MousePointer = 0
MsgBox ("Weapon created.")
Me.Hide
Unload Me


End Sub

Private Sub Form_Load()
' Add the weapon types here

Combo1.AddItem ("sword")
Combo1.AddItem ("2-hsword")
Combo1.AddItem ("polearm")
Combo1.AddItem ("flail")
Combo1.AddItem ("mace")
Combo1.AddItem ("staff")
Combo1.AddItem ("axe")
Combo1.AddItem ("dagger")
Combo1.AddItem ("bow")

Combo2.AddItem ("GENERAL 0")
Combo2.AddItem ("SLASH 1")
Combo2.AddItem ("PIERCE 2")
Combo2.AddItem ("BLUNT 3")
Combo2.AddItem ("UNARMED 4")
Combo2.AddItem ("MAGIC 5")
Combo2.AddItem ("FIRE 6")
Combo2.AddItem ("SHOCK 7")
Combo2.AddItem ("ACID 8")
Combo2.AddItem ("COLD 9")
Combo2.AddItem ("POISON 10")
Combo2.AddItem ("VORPAL 11")
Combo2.AddItem ("CLAW 12")
Combo2.AddItem ("BITE 13")
Combo2.AddItem ("PSYCHIC 14")
Combo2.AddItem ("STINGER 15")
Combo2.AddItem ("DRAIN 16")
Combo2.AddItem ("HOLY 17")
Combo2.AddItem ("SOUND 18")
Combo2.AddItem ("PETRIFY 21")
Combo2.AddItem ("BRAWLING 22")
Combo2.AddItem ("CHAOS 23")
Combo2.AddItem ("WARP 24")
Combo2.AddItem ("CLEAVE 25")
Combo2.AddItem ("BOW 26")

End Sub
