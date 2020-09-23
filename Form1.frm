VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Searching Altavista"
   ClientHeight    =   8595
   ClientLeft      =   900
   ClientTop       =   180
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10440
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6720
      ScaleHeight     =   495
      ScaleWidth      =   2655
      TabIndex        =   5
      Top             =   360
      Width           =   2655
      Begin VB.Label Label1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   13
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   12
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   10
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   9
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   135
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7095
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12515
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "no."
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "URL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   13123
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "home"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Double click the number to open the website."
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   9600
      TabIndex        =   17
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Searched:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   16
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   240
      X2              =   11400
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pageno As String
Dim number As Integer
'--searches...
Private Sub Command1_Click()
Call Search(Text2.Text, 0)
Label1(1).ForeColor = vbBlack
Label3.Caption = Text2.Text
End Sub
'---resets textbox
Private Sub Command2_Click()
Text2.Text = ""
End Sub
'---adds to listView-----
Private Sub sOutput(intNo, strTitle As String, strURL As String, strDescription As String)
    Dim itm As ListItem
    Set itm = ListView1.ListItems.Add(1, , intNo)
    itm.SubItems(1) = strTitle
    itm.SubItems(2) = strURL
    itm.SubItems(3) = strDescription
    Set itm = Nothing
End Sub
'--this is the main code---
Private Function Search(words As String, pageno As Integer)
Dim i
Dim Number2 As String
Dim strWebTxt
Dim Title As String
Dim URL As String
Dim Desc As String
Dim nCategory
Dim n1
Dim n2
Dim a1 As String
Dim b1 As String
Dim b2 As String
Dim b3
Dim NewURL As String


If pageno = 0 Then
    Number2 = ""
Else
End If
'clear listview
    ListView1.ListItems.Clear
'search------
NewURL = Replace(words, " ", "+")

Label4.Caption = "Searching........"
strWebTxt = Inet1.OpenURL("http://www.altavista.com/cgi-bin/query?q=" & words & "&kl=XX&pg=q&Translate=on&stq=" & pageno)
Label4.Caption = "Completed"
Me.Caption = "Searching Altavista - " & words
'-END--------
For i = 1 To 10

nCategory = InStr(1, strWebTxt, i & ".<")

If nCategory > 0 Then
Else
    a1 = i - 1
    GoTo 1
End If
Next i
a1 = 10
1

If a1 = "" Then
    MsgBox "No Results found for your query.", vbCritical, "Search Error"
    Exit Function
End If
    For i = 1 To a1
    '------FIND 1st link------'
    nCategory = InStr(1, strWebTxt, Number2 & i & ".<")
    n1 = InStr(nCategory, strWebTxt, "href=") + 11
    n2 = InStr(n1, strWebTxt, "onM") - 2
    URL = Mid(strWebTxt, n1, (n2 - n1))
    '---FIND TITLE-------
    n1 = InStr(n2, strWebTxt, "true;") + 7
    n2 = InStr(n1, strWebTxt, "</a>")
    Title = Mid(strWebTxt, n1, (n2 - n1))

    '---find description---
    n1 = InStr(n2, strWebTxt, "<dd>") + 5
    n2 = InStr(n1, strWebTxt, "<br>")
    
    Desc = Mid(strWebTxt, n1, (n2 - n1))
    b1 = Mid(Desc, 1, 73)

'puts a second row of description if too long for one row
If Len(Desc) >= 65 Then
    b3 = 73
    b2 = Mid(Desc, b3, 73)
    sOutput " ", " ", " ", b2
    sOutput i, Title, URL, b1
Else
    sOutput i, Title, URL, Desc
End If
sOutput "", "", "", ""
Next i
End Function




'below here is to edit the colour of the numbers if u wnat to go to the next page'
'and also search a next page
Private Sub Label1_Click(Index As Integer)
Search Label3.Caption, Index
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
GetNumber
If number = Index Then
    Exit Sub
Else
    Label1(Index).ForeColor = vbRed
End If
End Sub
'this opens the site in the default browser
Private Sub ListView1_DblClick()
Shell ("start " & ListView1.SelectedItem.ListSubItems(2).Text)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i
For i = 1 To 10
GetNumber

If i = number Then
    Label1(i).ForeColor = vbBlack
    GoTo 1
Else
    Label1(i).ForeColor = vbBlue
End If
1
Next i
End Sub
