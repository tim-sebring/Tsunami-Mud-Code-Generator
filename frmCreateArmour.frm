VERSION 5.00
Begin VB.Form frmCreateArmour 
   Caption         =   "Create Armor"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmCreateArmour.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   1680
      TabIndex        =   34
      Top             =   4800
      Width           =   6255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Metal Item"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCreateArmour.frx":0442
      Left            =   1440
      List            =   "frmCreateArmour.frx":0444
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5760
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
      Width           =   6255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   3720
      Width           =   6255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   4080
      Width           =   6255
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   4440
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Magical Item"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Lockerable Item"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Armour Info:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Armour AC:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Armour Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Armour's Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Armour Aliases:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Armour's ID (opt)"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Armour's Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Armour's Weight:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Armour's Set_Short:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Set Long Line 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Set Long Line 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Set Long Line 3:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Set Long Line 4:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1455
   End
End
Attribute VB_Name = "frmCreateArmour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'ignore if they type something in here.
Exit Sub


End Sub

Private Sub Command1_Click()

frmCreateArmour.Hide
Unload frmCreateArmour

End Sub

Private Sub Command2_Click()

' Create armour
Dim Temp As String, AType As String
MousePointer = 11

Temp = (App.Path & "\Code\" & Text1.Text)

Open Temp For Output Access Write As #1

Print #1, "inherit  STD_ARMOUR; "
Print #1,
Print #1, "reset (arg)"
Print #1, "{"
Print #1, "    ::reset( arg );"
Print #1, "    if( arg ) return;"
Print #1, "    set_name(""" & Text2.Text & """);"
If Len(Text6.Text) > 0 Then
    Print #1, "    set_id(""" & Text6.Text & """);"
End If
Select Case Combo1.Text
Case "Body Armor"
    AType = "armour"
Case "Hat/Helmet"
    AType = "helmet"
Case "Boots"
    AType = "boots"
Case "Gloves"
    AType = "gloves"
Case "Ring"
    AType = "ring"
Case "Amulet"
    AType = "amulet"
Case "Cloak"
    AType = "cloak"
Case "Eyewear"
    AType = "eyewear"
Case "Other"
    AType = "other"
Case Else
    MsgBox ("Armor Type is not a valid choice. Please reselect.")
    Close (1)
    Exit Sub
End Select
Print #1, "    set_type(""" & AType & """); "
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
Print #1, "    set_ac(" & Text14.Text & "); "

Print #1, "    set_value(" & Text7.Text & ");"
Print #1, "    set_weight(" & Text8.Text & ");"
If (Check3.Value = 1) Then
    Print #1, "    set_metal(1);"
Else
    Print #1, "    set_metal(0);"
End If


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
If Len(Text15.Text) > 0 Then
    Print #1, "query_info() { return """ & Text15.Text & """;"
End If
Close (1)
MousePointer = 0
MsgBox ("Armor created.")
Me.Hide
Unload Me

End Sub

Private Sub Form_Load()

' Add the armour types here

Combo1.AddItem ("Body Armor")
Combo1.AddItem ("Hat/Helmet")
Combo1.AddItem ("Boots")
Combo1.AddItem ("Gloves")
Combo1.AddItem ("Ring")
Combo1.AddItem ("Amulet")
Combo1.AddItem ("Cloak")
Combo1.AddItem ("Eyewear")
Combo1.AddItem ("Other")


End Sub
