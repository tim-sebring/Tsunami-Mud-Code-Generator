VERSION 5.00
Begin VB.Form frmCreateObject 
   Caption         =   "Create Object"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   Icon            =   "frmCreateObject.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Lockerable Item"
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Magical Item"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save and Exit"
      Height          =   495
      Left            =   4920
      TabIndex        =   27
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   26
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1680
      TabIndex        =   23
      Top             =   3720
      Width           =   6255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1680
      TabIndex        =   22
      Top             =   3360
      Width           =   6255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   2640
      Width           =   6255
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Set Long Line 4:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Set Long Line 3:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Set Long Line 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Set Long Line 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Object's Set_Short:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Object's Weight:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Object's Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Object's ID (opt)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Object Aliases:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Object's Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmCreateObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

' Cancel the operation
frmCreateObject.Hide
Unload frmCreateObject


End Sub

Private Sub Command2_Click()
' Save and exit on Create Object Form
Dim Temp As String


If (Len(Text1.Text) = 0) Then
' No filename
    MsgBox ("No filename. Click Cancel to end session without saving.")
    Exit Sub
End If
MousePointer = 11


Temp = App.Path & "\Code\" & Text1.Text

Open Temp For Output Access Write As #1

Print #1, "inherit  STD_TREASURE; "
Print #1,
Print #1, "reset (arg)"
Print #1, "{"
Print #1, "    ::reset( arg );"
Print #1, "    if( arg ) return;"
Print #1, "    set_name(""" & Text2.Text & """);"

If (Len(Text3.Text) > 0) Then
    Print #1, "    set_alias(""" & Text3.Text & """); "
End If

If (Len(Text4.Text) > 0) Then
    Print #1, "    set_alias(""" & Text4.Text & """); "
End If

If (Len(Text5.Text) > 0) Then
    Print #1, "    set_alias(""" & Text5.Text & """); "
End If

If (Len(Text6.Text) > 0) Then
    Print #1, "    set_id(""" & Text6.Text & """); "
End If

Print #1, "    set_value(" & Text7.Text & ");"
Print #1, "    set_weight(" & Text8.Text & ");"
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
Print #1,
Print #1, "}"
Print #1, " /* This is for add_action, remove comments to use. */"
Print #1, " /*"
Print #1, "init()  {"
Print #1,
Print #1, "    ::init();"
Print #1, "    add_action(""_functionname"",""actionname"");"
Print #1,
Print #1, "}"
Print #1,
Print #1, "_functionname(str) {"
Print #1, "    if(!str) { notify_fail(""Do What?\n""); return 0; }"
Print #1, "    if( str == ""objectname"") {"
Print #1, "        write(""You need to change this.\n"");"
Print #1, "        say(this_player()->query_name()+""does something.\n"");"
Print #1, "        return 1;"
Print #1, "    } "
Print #1, " }"
Print #1, " */"
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
MsgBox ("Object created.")

Me.Hide
Unload Me

End Sub

