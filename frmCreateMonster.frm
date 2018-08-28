VERSION 5.00
Begin VB.Form frmCreateMonster 
   Caption         =   "Create Monster"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "frmCreateMonster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text21 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   23
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox Text20 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Top             =   6600
      Width           =   3255
   End
   Begin VB.TextBox Text19 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   21
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   20
      Top             =   5880
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Set Spell Message?"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   6120
      Width           =   2055
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Add a 'monster_died' Fn?"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   5160
      Width           =   6255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   4800
      Width           =   6255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   4440
      Width           =   6255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   4080
      Width           =   6255
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCreateMonster.frx":0442
      Left            =   1680
      List            =   "frmCreateMonster.frx":0444
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton butSaveExit 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   4320
      TabIndex        =   25
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "Spell Msg 2 (what target sees)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   49
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label23 
      Caption         =   "Spell Msg 1 (what others see)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   48
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label22 
      Caption         =   "Spell Damage:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   47
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "Spell Chance (%):"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   46
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "(Raw text, add_obj, etc)"
      Height          =   255
      Left            =   6480
      TabIndex        =   45
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   "Plain Text to Insert:"
      Height          =   255
      Left            =   360
      TabIndex        =   44
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "0=No, 1=Yes"
      Height          =   255
      Left            =   3840
      TabIndex        =   43
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Aggressive:"
      Height          =   255
      Left            =   360
      TabIndex        =   42
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0=Neuter, 1=Male, 2=Female"
      Height          =   255
      Left            =   3840
      TabIndex        =   41
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label15 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "(eg: ""/std/monster/aeromancer.c"")"
      Height          =   255
      Left            =   3720
      TabIndex        =   39
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   "Set Long Line 4:"
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Set Long Line 3:"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Set Long Line 2:"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Set Long Line 1:"
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Monster's Set_Short:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Random Pick:"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Monster's Level:"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Alignment:"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Monster Aliases:"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Monster's Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Stanard Inherit:"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Add_Money:"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmCreateMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butCancel_Click()

Me.Hide
Unload Me

End Sub

Private Sub butSaveExit_Click()
' Create Monster

Dim Temp As String
MousePointer = 11
Temp = App.Path & "\Code\" & Text1.Text

Open Temp For Output Access Write As #1
If Len(Combo1.Text) > 0 Then
    Print #1, " inherit ""/std/monster/" & Combo1.Text & """;"
Else
    Print #1, "inherit STD_MONSTER;"
End If

Print #1,
Print #1, "reset (arg)"
Print #1, "{"
Print #1, "    ::reset( arg );"
Print #1, "    if( arg ) return;"
Print #1, "    set_name(""" & Text2.Text & """);"

If Len(Text3.Text) > 0 Then
    Print #1, "    add_alias(""" & Text3.Text & """);"
End If
If Len(Text4.Text) > 0 Then
    Print #1, "    add_alias(""" & Text4.Text & """);"
End If
If Len(Text5.Text) > 0 Then
    Print #1, "    add_alias(""" & Text5.Text & """);"
End If
If Check1.Value = 1 Then
    Print #1, "    set_dead_ob(this_object());"
End If

Print #1, "    add_money(" & Text14.Text & ");"
Print #1, "    set_alignment(" & Text6.Text & ");"
Print #1, "    set_random_pick(" & Text8.Text & ");"

Print #1, "    set_gender(" & Text15.Text & ");"
Print #1, "    set_aggressive(" & Text16.Text & ");"

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

Print #1, "    set_level(" & Text7.Text & "); "
If Len(Text17.Text) > 0 Then
    Print #1, Text17.Text
End If

If Check3.Value = 1 Then
    Print #1, "    set_chance(" & Text18.Text & ");"
    Print #1, "    set_spell_dam(" & Text19.Text & ");"
    Print #1, "    set_spell_mess1(""" & Text20.Text & """);"
    Print #1, "    set_spell_mess2(""" & Text21.Text & """);"
End If

Print #1,
Print #1, "}"


Print #1,
If Check1.Value = 1 Then
    Print #1, "monster_died() {"
    Print #1, "object ob;"
    Print #1, " tell_object(this_player(),""The creature slumps over.\n"");"
    Print #1, "}"
End If

Close (1)
MousePointer = 0
MsgBox ("Monster created.")
Me.Hide
Unload Me

End Sub

Private Sub Check3_Click()

If Check3.Value = 1 Then
    Label21.Enabled = True
    Label22.Enabled = True
    Label23.Enabled = True
    Label24.Enabled = True
    Text18.Enabled = True
    Text19.Enabled = True
    Text20.Enabled = True
    Text21.Enabled = True
Else
    Label21.Enabled = False
    Label22.Enabled = False
    Label23.Enabled = False
    Label24.Enabled = False
    Text18.Enabled = False
    Text19.Enabled = False
    Text20.Enabled = False
    Text21.Enabled = False
End If

End Sub

Private Sub Form_Load()

' Add the combo boxes
Combo1.AddItem ("aeromancer.c")
Combo1.AddItem ("amazon.c")
Combo1.AddItem ("animal.c")
Combo1.AddItem ("bandit.c")
Combo1.AddItem ("banshee.c")
Combo1.AddItem ("barb2.c")
Combo1.AddItem ("barbarian.c")
Combo1.AddItem ("bard.c")
Combo1.AddItem ("basic.c")
Combo1.AddItem ("cavalier.c")
Combo1.AddItem ("changeling.c")
Combo1.AddItem ("chaosmage.c")
Combo1.AddItem ("chronomancer.c")
Combo1.AddItem ("cleric.c")
Combo1.AddItem ("cleric2.c")
Combo1.AddItem ("deathknight.c")
Combo1.AddItem ("demon.c")
Combo1.AddItem ("dracolich.c")
Combo1.AddItem ("dragon.c")
Combo1.AddItem ("druid.c")
Combo1.AddItem ("enchanter.c")
Combo1.AddItem ("evoker.c")
Combo1.AddItem ("faeriedragon.c")
Combo1.AddItem ("fighter.c")
Combo1.AddItem ("figsub.c")
Combo1.AddItem ("geomancer.c")
Combo1.AddItem ("giant.c")
Combo1.AddItem ("gremlin.c")
Combo1.AddItem ("hunter.c")
Combo1.AddItem ("hydra.c")
Combo1.AddItem ("hydromancer.c")
Combo1.AddItem ("illusionist.c")
Combo1.AddItem ("knight.c")
Combo1.AddItem ("knight2.c")
Combo1.AddItem ("mage.c")
Combo1.AddItem ("mindflayer.c")
Combo1.AddItem ("monk.c")
Combo1.AddItem ("necromancer.c")
Combo1.AddItem ("ninja.c")
Combo1.AddItem ("ooze.c")
Combo1.AddItem ("ooze2.c")
Combo1.AddItem ("paladin.c")
Combo1.AddItem ("priest.c")
Combo1.AddItem ("pyromancer.c")
Combo1.AddItem ("samurai.c")
Combo1.AddItem ("shade.c")
Combo1.AddItem ("shadow.c")
Combo1.AddItem ("spider.c")
Combo1.AddItem ("swashbuckler.c")
Combo1.AddItem ("templar.c")
Combo1.AddItem ("thief.c")
Combo1.AddItem ("thug.c")
Combo1.AddItem ("troll.c")
Combo1.AddItem ("undead.c")
Combo1.AddItem ("unicorn.c")
Combo1.AddItem ("unicorn2.c")
Combo1.AddItem ("vamp2.c")
Combo1.AddItem ("vampire.c")
Combo1.AddItem ("vampyre.c")
Combo1.AddItem ("worm.c")

End Sub


