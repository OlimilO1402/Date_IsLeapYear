VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Is it a Leap Year?"
   ClientHeight    =   855
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":1782
      Left            =   3120
      List            =   "Form1.frx":1784
      TabIndex        =   1
      Top             =   150
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":1786
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldFormat As String

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    With Combo1
        .AddItem "Long Date"
        .AddItem "dddd, mmm yyyy"
        
'        .AddItem "d  Kurzes Datumsmuster."
'        .AddItem "D  Langes Datumsmuster."
'        .AddItem "f  Vollständiges Datums-/Zeitmuster (kurze Zeit)."
'        .AddItem "F  Vollständiges Datums-/Zeitmuster (lange Zeit)."
'        .AddItem "g  Allgemeines Datums-/Zeitmuster (kurze Zeit)."
'        .AddItem "G  Allgemeines Datums-/Zeitmuster (lange Zeit)."
'        .AddItem "M, m  Monatstagmuster."
'        .AddItem "O, o  Datums-/Uhrzeitmuster für Roundtrip."
'        .AddItem "R, r  RFC1123-Muster."
'        .AddItem "s  Sortierbares Datums-/Zeitmuster."
'        .AddItem "t  Kurzes Zeitmuster."
'        .AddItem "T  Langes Zeitmuster."
'        .AddItem "u  Universelles, sortierbares Datums-/Zeitmuster."
'        .AddItem "U  Universelles Datums-/Zeitmuster (Koordinierte Weltzeit)."
'        .AddItem "Y, y  Jahr-Monat-Muster."
        '.ListIndex = 1
    End With
    OldFormat = "Long Date"
    
    Text1.Text = Format(Now, Combo1.Text) 'Left(Combo1.Text, 1)) ' '"14. feb. 1995"
    'Text1_LostFocus
End Sub

Private Sub Command1_Click()
'1 220 1 1 472
'MADE IN USA 2011 Worn Brown
'
'Y DDD Y B RRR
'
'1 290 1 1 379
'MADE IN USA 2011 Worn Cherry
'
'001 - 499: Kalamazoo
'500 - 999: Nashville
    
    
    MsgBox Date_ParseFromDayNumber(2011, 220)
    MsgBox Date_ParseFromDayNumber(2011, 290)
    
End Sub

Private Sub Combo1_Click()
    Dim d As Date:
    If Date_TryParse(Text1.Text, d) Then
        Text1.Text = Format(d, Combo1.Text)
    End If
    OldFormat = Combo1.Text
End Sub

Private Sub Label1_Click()
    Text1_LostFocus
End Sub

Private Sub Text1_LostFocus()
    Dim s As String: s = Text1.Text
    Dim d As Date
    If Not Date_TryParse(s, d) Then Exit Sub
    s = s & "The date " & d & " was the " & CStr(DayOfYear(d)) & ". day of" & vbCrLf
    Dim y As Long: y = year(d)
    s = s & "the year " & CStr(y) & " that was " & IIf(IsLeapYear(y), "", "not ") & "a leap year."
    Text2.Text = s
End Sub

