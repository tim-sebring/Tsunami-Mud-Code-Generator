VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCreateRoom 
   Caption         =   "Create Room"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "frmCreateRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   240
      TabIndex        =   69
      Top             =   4080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Add Exits"
      TabPicture(0)   =   "frmCreateRoom.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text5"
      Tab(0).Control(1)=   "Text4"
      Tab(0).Control(2)=   "Text3"
      Tab(0).Control(3)=   "Text15"
      Tab(0).Control(4)=   "Text16"
      Tab(0).Control(5)=   "Text17"
      Tab(0).Control(6)=   "Text18"
      Tab(0).Control(7)=   "Text19"
      Tab(0).Control(8)=   "Text20"
      Tab(0).Control(9)=   "Text21"
      Tab(0).Control(10)=   "Text22"
      Tab(0).Control(11)=   "Text23"
      Tab(0).Control(12)=   "Label16"
      Tab(0).Control(13)=   "Label17"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Add Items"
      TabPicture(1)   =   "frmCreateRoom.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label21"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text31"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text30"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text29"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text28"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text27"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text26"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text25"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text6"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text7"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text8"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Add Objects"
      TabPicture(2)   =   "frmCreateRoom.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text44"
      Tab(2).Control(1)=   "Text43"
      Tab(2).Control(2)=   "Check6"
      Tab(2).Control(3)=   "Text42"
      Tab(2).Control(4)=   "Text41"
      Tab(2).Control(5)=   "Check2"
      Tab(2).Control(6)=   "Text40"
      Tab(2).Control(7)=   "Text39"
      Tab(2).Control(8)=   "Check3"
      Tab(2).Control(9)=   "Text38"
      Tab(2).Control(10)=   "Text37"
      Tab(2).Control(11)=   "Check4"
      Tab(2).Control(12)=   "Text36"
      Tab(2).Control(13)=   "Text35"
      Tab(2).Control(14)=   "Check5"
      Tab(2).Control(15)=   "Label3"
      Tab(2).Control(16)=   "Label26"
      Tab(2).Control(17)=   "Label25"
      Tab(2).Control(18)=   "Label24"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Listen/Smell"
      TabPicture(3)   =   "frmCreateRoom.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text32"
      Tab(3).Control(1)=   "Text33"
      Tab(3).Control(2)=   "Label5"
      Tab(3).Control(3)=   "Label6"
      Tab(3).ControlCount=   4
      Begin VB.TextBox Text32 
         Height          =   285
         Left            =   -72240
         TabIndex        =   48
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox Text33 
         Height          =   285
         Left            =   -72240
         TabIndex        =   49
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox Text44 
         Height          =   285
         Left            =   -74880
         TabIndex        =   33
         Top             =   1260
         Width           =   4215
      End
      Begin VB.TextBox Text43 
         Height          =   285
         Left            =   -70320
         TabIndex        =   34
         Top             =   1260
         Width           =   1335
      End
      Begin VB.CheckBox Check6 
         Height          =   255
         Left            =   -68520
         TabIndex        =   35
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox Text42 
         Height          =   285
         Left            =   -74880
         TabIndex        =   36
         Top             =   1620
         Width           =   4215
      End
      Begin VB.TextBox Text41 
         Height          =   285
         Left            =   -70320
         TabIndex        =   37
         Top             =   1620
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   -68520
         TabIndex        =   38
         Top             =   1620
         Width           =   375
      End
      Begin VB.TextBox Text40 
         Height          =   285
         Left            =   -74880
         TabIndex        =   39
         Top             =   1980
         Width           =   4215
      End
      Begin VB.TextBox Text39 
         Height          =   285
         Left            =   -70320
         TabIndex        =   40
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   -68520
         TabIndex        =   41
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox Text38 
         Height          =   285
         Left            =   -74880
         TabIndex        =   42
         Top             =   2340
         Width           =   4215
      End
      Begin VB.TextBox Text37 
         Height          =   285
         Left            =   -70320
         TabIndex        =   43
         Top             =   2340
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   -68520
         TabIndex        =   44
         Top             =   2340
         Width           =   375
      End
      Begin VB.TextBox Text36 
         Height          =   285
         Left            =   -74880
         TabIndex        =   45
         Top             =   2700
         Width           =   4215
      End
      Begin VB.TextBox Text35 
         Height          =   285
         Left            =   -70320
         TabIndex        =   46
         Top             =   2700
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   -68520
         TabIndex        =   47
         Top             =   2700
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   600
         TabIndex        =   25
         Top             =   1740
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   1380
         Width           =   5415
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   600
         TabIndex        =   31
         Top             =   2820
         Width           =   1215
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   600
         TabIndex        =   29
         Top             =   2460
         Width           =   1215
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   600
         TabIndex        =   27
         Top             =   2100
         Width           =   1215
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   2820
         Width           =   5415
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   2460
         Width           =   5415
      End
      Begin VB.TextBox Text30 
         Height          =   285
         Left            =   2280
         TabIndex        =   28
         Top             =   2100
         Width           =   5415
      End
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   1740
         Width           =   5415
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -74280
         TabIndex        =   15
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -74280
         TabIndex        =   13
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   11
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   -74280
         TabIndex        =   17
         Top             =   2580
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   -74280
         TabIndex        =   19
         Top             =   2940
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   -74280
         TabIndex        =   21
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   -72840
         TabIndex        =   12
         Top             =   1500
         Width           =   5535
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   -72840
         TabIndex        =   20
         Top             =   2940
         Width           =   5535
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   -72840
         TabIndex        =   18
         Top             =   2580
         Width           =   5535
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   -72840
         TabIndex        =   16
         Top             =   2220
         Width           =   5535
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   -72840
         TabIndex        =   14
         Top             =   1860
         Width           =   5535
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   -72840
         TabIndex        =   22
         Top             =   3300
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "(NOTE: Use quotes if it's an absolute path)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Listen_Room String:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   78
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Smell_Room String:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   77
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Path and Filename of Objects:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   76
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label25 
         Caption         =   "Number of Objects:"
         Height          =   255
         Left            =   -70320
         TabIndex        =   75
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Unique?"
         Height          =   255
         Left            =   -68640
         TabIndex        =   74
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Item Description:"
         Height          =   255
         Left            =   2280
         TabIndex        =   73
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Item Name:"
         Height          =   255
         Left            =   600
         TabIndex        =   72
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Direction:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   71
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Path: (NOTE: Use quotes if it's an absolute path)"
         Height          =   255
         Left            =   -72840
         TabIndex        =   70
         Top             =   1140
         Width           =   4095
      End
   End
   Begin VB.TextBox Text34 
      Height          =   285
      Left            =   3720
      TabIndex        =   51
      Top             =   7920
      Width           =   4455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "NG (No gate/summon)"
      Height          =   255
      Left            =   360
      TabIndex        =   50
      Top             =   7920
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "0 - Temperate"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCreateRoom.frx":04B2
      Left            =   1800
      List            =   "frmCreateRoom.frx":04B4
      TabIndex        =   4
      Text            =   "forest"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   6255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   3360
      Width           =   6255
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   3720
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   52
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   53
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   "Location:"
      Height          =   255
      Left            =   2640
      TabIndex        =   68
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Room's Climate:"
      Height          =   255
      Left            =   240
      TabIndex        =   67
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "(e.g. path.h)"
      Height          =   255
      Left            =   4080
      TabIndex        =   66
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "File to Include:"
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "(0 for Darkness at Night)"
      Height          =   255
      Left            =   3960
      TabIndex        =   64
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "(1-4)"
      Height          =   255
      Left            =   3960
      TabIndex        =   63
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Room Light:"
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Room's Terrain:"
      Height          =   255
      Left            =   240
      TabIndex        =   61
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Room's Outdoors:"
      Height          =   255
      Left            =   240
      TabIndex        =   59
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Room's Set_Short:"
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Set Long Line 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Set Long Line 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Set Long Line 3:"
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Set Long Line 4:"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmCreateRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' Cancel

Me.Hide
Unload Me

End Sub

Private Sub Command2_Click()
' Create Room
Dim Temp As String, AType As String

MousePointer = 11
Temp = (App.Path & "\Code\" & Text1.Text)

Open Temp For Output Access Write As #1

Print #1, "inherit STD_ROOM; "
If Len(Text24.Text) > 0 Then
    Print #1, "#include """ & Text24.Text & """"
End If
Print #1,
Print #1,
Print #1, "reset(arg) {"
Print #1, "    ::reset(arg);"
Print #1, "    if(arg) return;"
Print #1, "    set_outdoors(" & Text2.Text & "); "
Print #1, "    set_terrain(""" & Trim(Combo1.Text) & """); "
Print #1, "    set_light(" & Text14.Text & "); "
Print #1, "    set_climate(" & Trim(Left(Combo2.Text, 1)) & ");   // " & Mid(Combo2.Text, 5)
Print #1, "    set_short(""" & Trim(Text9.Text) & """);"
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

' Add the add_items
If (Len(Text6.Text) > 0) Then
    Print #1, "    add_item(""" & Text6.Text & """,""" & Text7.Text & """); "
End If

If (Len(Text8.Text) > 0) Then
    Print #1, "    add_item(""" & Text8.Text & """,""" & Text31.Text & """); "
End If

If (Len(Text27.Text) > 0) Then
    Print #1, "    add_item(""" & Text27.Text & """,""" & Text30.Text & """); "
End If

If (Len(Text26.Text) > 0) Then
    Print #1, "    add_item(""" & Text26.Text & """,""" & Text29.Text & """); "
End If

If (Len(Text25.Text) > 0) Then
    Print #1, "    add_item(""" & Text25.Text & """,""" & Text28.Text & """); "
End If
'***********************************************************************
' Add the add_exits
' This code assumes that a relative path will be used, for example
' add_exit(ROOM+"blah.c","north");

If (Len(Text18.Text) > 0) Then
    Print #1, "    add_exit(" & Text18.Text & ",""" & Text3.Text & """); "
End If

If (Len(Text22.Text) > 0) Then
    Print #1, "    add_exit(" & Text22.Text & ",""" & Text4.Text & """); "
End If

If (Len(Text21.Text) > 0) Then
    Print #1, "    add_exit(" & Text21.Text & ",""" & Text5.Text & """); "
End If

If (Len(Text20.Text) > 0) Then
    Print #1, "    add_exit(" & Text20.Text & ",""" & Text15.Text & """); "
End If

If (Len(Text19.Text) > 0) Then
    Print #1, "    add_exit(" & Text19.Text & ",""" & Text16.Text & """); "
End If

If (Len(Text23.Text) > 0) Then
    Print #1, "    add_exit(" & Text23.Text & ",""" & Text17.Text & """); "
End If
'***********************************************************************
' Now we have to add the add_objects to the code

If (Len(Text44.Text) > 0) Then
    If Check6.Value = 1 Then
        Print #1, "    add_obj(" & Text44.Text & "," & Text43.Text & ",""exist"");"
    Else
        Print #1, "    add_obj(" & Text44.Text & "," & Text43.Text & ",""present"");"
    End If
End If

If (Len(Text42.Text) > 0) Then
    If Check2.Value = 1 Then
        Print #1, "    add_obj(" & Text42.Text & "," & Text41.Text & ",""exist"");"
    Else
        Print #1, "    add_obj(" & Text42.Text & "," & Text41.Text & ",""present"");"
    End If
End If

If (Len(Text40.Text) > 0) Then
    If Check3.Value = 1 Then
        Print #1, "    add_obj(" & Text40.Text & "," & Text39.Text & ",""exist"");"
    Else
        Print #1, "    add_obj(" & Text40.Text & "," & Text39.Text & ",""present"");"
    End If
End If

If (Len(Text38.Text) > 0) Then
    If Check4.Value = 1 Then
        Print #1, "    add_obj(" & Text38.Text & "," & Text37.Text & ",""exist"");"
    Else
        Print #1, "    add_obj(" & Text38.Text & "," & Text37.Text & ",""present"");"
    End If
End If

If (Len(Text36.Text) > 0) Then
    If Check5.Value = 1 Then
        Print #1, "    add_obj(" & Text36.Text & "," & Text35.Text & ",""exist"");"
    Else
        Print #1, "    add_obj(" & Text36.Text & "," & Text35.Text & ",""present"");"
    End If
End If
'***********************************************************************
' END OF RESET FUNCTION
Print #1, "}"
'***********************************************************************
' And here can go the listen
If Len(Text32.Text) > 0 Then
    Print #1, "int listen_room(){"
    Print #1, "     tell_object(this_player(), """ & Text32.Text & "\n"");"
    Print #1, "      return 1;"
    Print #1, "  }"
    Print #1, ""
End If
'***********************************************************************
' And here's the smell
If Len(Text33.Text) > 0 Then
    Print #1, "int smell_room(){"
    Print #1, "     tell_object(this_player(), """ & Text33.Text & "\n"");"
    Print #1, "      return 1;"
    Print #1, "  }"
    Print #1, ""
End If
'***********************************************************************

Print #1, ""
Print #1, "/*"
Print #1, "// Uncomment this if you want an add_action here"
Print #1, "init() {"
Print #1, "    ::init();"

Print #1, "    add_action(""_blah"",""blah"");"

Print #1, "}"

Print #1, "_blah(str) {"
Print #1, "some stuff here"
Print #1, " }"
Print #1, "*/"

Print #1, "query_location() { return """ & Text34.Text & """; }"
If Check1.Value = 1 Then
    Print #1, "realm() { return ""NG""; }"
End If
Print #1, "query_realmowner() { return ""Kane""; }"

Close (1)
MousePointer = 0
MsgBox ("Room created.")
Me.Hide
Unload Me


End Sub



Private Sub Form_Load()

Combo1.AddItem ("forest")
Combo1.AddItem ("hills")
Combo1.AddItem ("mountains ")
Combo1.AddItem ("plains ")
Combo1.AddItem ("swamp ")
Combo1.AddItem ("jungle")
Combo1.AddItem ("desert ")
Combo1.AddItem ("tundra ")
Combo1.AddItem ("underground ")
Combo1.AddItem ("shore ")
Combo1.AddItem ("underwater ")
Combo1.AddItem ("water ")
Combo1.AddItem ("building ")
Combo1.AddItem ("town ")
Combo1.AddItem ("farm ")
Combo1.AddItem ("cemetery ")
Combo1.AddItem ("road ")
Combo1.AddItem ("bridge ")
Combo1.AddItem ("air ")
Combo1.AddItem ("dimension ")
Combo1.AddItem ("ship ")
Combo1.AddItem ("garden ")
Combo1.AddItem ("field ")

Combo2.AddItem ("0 - Temperate")
Combo2.AddItem ("1 - Subterranean")
Combo2.AddItem ("2 - Arctic")
Combo2.AddItem ("3 - Sub-arctic")
Combo2.AddItem ("4 - Temperate")
Combo2.AddItem ("5 - Tropical")
Combo2.AddItem ("6 - Arid")



End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Caption
Case "Add Exits"
    Text3.SetFocus
Case "Add Items"
    Text6.SetFocus
Case "Add Objects"
    Text44.SetFocus
Case "Listen/Smell"
    Text32.SetFocus
Case Else
    DoEvents
End Select



End Sub


Private Sub Text10_KeyPress(KeyAscii As Integer)
' If its a carriage return, set focus to next line
If KeyAscii = 13 Then
    Text11.SetFocus
End If

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
' If its a carriage return, set focus to next line
If KeyAscii = 13 Then
    Text12.SetFocus
End If

End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
' If its a carriage return, set focus to next line
If KeyAscii = 13 Then
    Text13.SetFocus
End If

End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
' If its a carriage return, set focus to next line
If KeyAscii = 13 Then
    Text10.SetFocus
End If

End Sub
