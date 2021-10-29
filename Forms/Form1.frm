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

Private Sub Command1_Click()
    
    MsgBox Date_ParseFromDayNumber(2011, 220)
    MsgBox Date_ParseFromDayNumber(2011, 290)
    
End Sub

Private Sub Form_Load()
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

Private Sub Combo1_Click()
    Dim d As Date:
    If Date_TryParse(Text1.Text, d) Then
        Text1.Text = Format(d, Combo1.Text)
    End If
    OldFormat = Combo1.Text
End Sub

Function Date_ParseFromDayNumber(ByVal y As Integer, ByVal DayNr As Integer) As Date
    Dim mds As Integer
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 1, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 28 - CInt(IsLeapYear(y))
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 2, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 3, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 30
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 4, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 5, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 30
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 6, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 7, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 8, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 30
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 9, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 10, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 30
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 11, DayNr): Exit Function
    DayNr = DayNr - mds
    
    mds = 31
    If DayNr <= mds Then Date_ParseFromDayNumber = DateSerial(y, 12, DayNr): Exit Function
End Function


Private Sub Label1_Click()
    Text1_LostFocus
End Sub

Private Sub Text1_LostFocus()
    Dim s       As String
    Dim strDate As String
    Dim d       As Date
    Dim y       As Long
    
    If Not Date_TryParse(Text1.Text, d) Then Exit Sub
    'strDate = Text1.Text
    'd = CDate(strDate)
    s = s & "The date " & d & " was the " & CStr(DayOfYear(d)) & ". day of" & vbCrLf
    y = year(d)
    s = s & "the year " & CStr(y) & " that was " & IIf(IsLeapYear(y), "", "not ") & "a leap year."
    Text2.Text = s

End Sub

Function Date_TryParse(ByVal s As String, ByRef out_date As Date) As Boolean
Try: On Error GoTo Catch
    If LCase(s) = "now" Or LCase(s) = "jetzt" Then s = Now
    out_date = CDate(s)
    Date_TryParse = True
    Exit Function
Catch:
    MsgBox Err.Number & " " & Err.Description
End Function

Public Function DayOfYear(d As Date) As Long
    Dim y As Long
    Dim i As Long
    y = year(d)
    For i = 1 To month(d) - 1
        DayOfYear = DayOfYear + DaysInMonth(y, i)
    Next
    DayOfYear = DayOfYear + Day(d) 'Day(d)=DayOfMonth
End Function

Public Function DaysInMonth(ByVal year As Long, ByVal month As Long) As Long
    Select Case month
    Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
    Case 2: If IsLeapYear(year) Then DaysInMonth = 29 Else DaysInMonth = 28
    Case 4, 6, 9, 11: DaysInMonth = 30
    End Select
End Function

Public Function IsLeapYear(ByVal y As Long) As Boolean
'Schaltjahr (LeapYear)
'a leap year is a year which is
'either (i.)
'    evenly divisible
'        by 4
'    and not
'        by 100
'or (ii.)
'    evenly divisible
'        by 400
    IsLeapYear = (((y Mod 4) = 0) And Not ((y Mod 100) = 0)) Or ((y Mod 400) = 0)
End Function


'https://docs.microsoft.com/de-de/dotnet/standard/base-types/standard-date-and-time-format-strings
'
'Formatbezeichner    Beschreibung    Beispiele
'"d"     Kurzes Datumsmuster.
'
'Weitere Informationen finden Sie unter Der Formatbezeichner für das kurze Datum („d“).  2009-06-15T13:45:30 -> 6/15/2009 (en-US)
'
'2009-06-15T13:45:30 -> 15/06/2009 (fr-FR)
'
'2009-06-15T13:45:30 -> 2009/06/15 (ja-JP)
'"D"     Langes Datumsmuster.
'
'Weitere Informationen finden Sie unter Der Formatbezeichner für das lange Datum („D“).  2009-06-15T13:45:30 -> Monday, June 15, 2009 (en-US)
'
'2009-06-15T13:45:30 -> 15 ???? 2009 ?. (ru-RU)
'
'2009-06-15T13:45:30 -> Montag, 15. Juni 2009 (de-DE)
'"f"     Vollständiges Datums-/Zeitmuster (kurze Zeit).
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für vollständiges Datum und kurze Zeit („f“).  2009-06-15T13:45:30 -> Monday, June 15, 2009 1:45 PM (en-US)
'
'2009-06-15T13:45:30 -> den 15 juni 2009 13:45 (sv-SE)
'
'2009-06-15T13:45:30 -> ?e?t??a, 15 ??????? 2009 1:45 µµ (el-GR)
'"F"     Vollständiges Datums-/Zeitmuster (lange Zeit).
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für vollständiges Datum und lange Zeit („F“).  2009-06-15T13:45:30 -> Monday, June 15, 2009 1:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> den 15 juni 2009 13:45:30 (sv-SE)
'
'2009-06-15T13:45:30 -> ?e?t??a, 15 ??????? 2009 1:45:30 µµ (el-GR)
'"g"     Allgemeines Datums-/Zeitmuster (kurze Zeit).
'
'Weitere Informationen finden Sie unter: Der allgemeine Formatbezeichner für Datum und kurze Zeit („g“).     2009-06-15T13:45:30 -> 6/15/2009 1:45 PM (en-US)
'
'2009-06-15T13:45:30 -> 15/06/2009 13:45 (es-ES)
'
'2009-06-15T13:45:30 -> 2009/6/15 13:45 (zh-CN)
'"G"     Allgemeines Datums-/Zeitmuster (lange Zeit).
'
'Weitere Informationen finden Sie unter: Der allgemeine Formatbezeichner für Datum und lange Zeit („G“).     2009-06-15T13:45:30 -> 6/15/2009 1:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> 15/06/2009 13:45:30 (es-ES)
'
'2009-06-15T13:45:30 -> 2009/6/15 13:45:30 (zh-CN)
'"M", "m"    Monatstagmuster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für den Monat („M“, „m“).  2009-06-15T13:45:30 -> June 15 (en-US)
'
'2009-06-15T13:45:30 -> 15. juni (da-DK)
'
'2009-06-15T13:45:30 -> 15 Juni (id-ID)
'"O", "o"    Datums-/Uhrzeitmuster für Roundtrip.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für Roundtrips („O“, „o“).     DateTime-Werte sind:
'
'2009-06-15T13:45:30 (DateTimeKind.Local) --> 2009-06-15T13:45:30.0000000-07:00
'
'2009-06-15T13:45:30 (DateTimeKind.Utc) --> 2009-06-15T13:45:30.0000000Z
'
'2009-06-15T13:45:30 (DateTimeKind.Unspecified) --> 2009-06-15T13:45:30.0000000
'
'DateTimeOffset:
'
'2009-06-15T13:45:30-07:00 --> 2009-06-15T13:45:30.0000000-07:00
'"R", "r"    RFC1123-Muster.
'
'Weitere Informationen finden Sie unter: Der RFC1123-Formatbezeichner („R“, „r“).    2009-06-15T13:45:30 -> Mon, 15 Jun 2009 20:45:30 GMT
'"s"     Sortierbares Datums-/Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der sortierbare Formatbezeichner („s“).     2009-06-15T13:45:30 (DateTimeKind.Local) -> 2009-06-15T13:45:30
'
'2009-06-15T13:45:30 (DateTimeKind.Utc) -> 2009-06-15T13:45:30
'"t"     Kurzes Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für kurze Zeit („t“).  2009-06-15T13:45:30 -> 1:45 PM (en-US)
'
'2009-06-15T13:45:30 -> 13:45 (hr-HR)
'
'2009-06-15T13:45:30 -> 01:45 ? (ar-EG)
'"T"     Langes Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für lange Zeit („T“).  2009-06-15T13:45:30 -> 1:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> 13:45:30 (hr-HR)
'
'2009-06-15T13:45:30 -> 01:45:30 ? (ar-EG)
'"u"     Universelles, sortierbares Datums-/Zeitmuster.
'
'Weitere Informationen finden Sie unter: Der universelle sortierbare Formatbezeichner („u“).     Mit einem DateTime-Wert: 2009-06-15T13:45:30 -> 2009-06-15 13:45:30Z
'
'Mit einem DateTimeOffset-Wert: 2009-06-15T13:45:30 -> 2009-06-15 20:45:30Z
'"U"     Universelles Datums-/Zeitmuster (Koordinierte Weltzeit).
'
'Weitere Informationen finden Sie unter: Der universelle vollständige Formatbezeichner („U“).    2009-06-15T13:45:30 -> Monday, June 15, 2009 8:45:30 PM (en-US)
'
'2009-06-15T13:45:30 -> den 15 juni 2009 20:45:30 (sv-SE)
'
'2009-06-15T13:45:30 -> ?e?t??a, 15 ??????? 2009 8:45:30 µµ (el-GR)
'"Y", "y"    Jahr-Monat-Muster.
'
'Weitere Informationen finden Sie unter: Der Formatbezeichner für Jahr-Monat („Y“).  2009-06-15T13:45:30 -> Juni 2009 (en-US)
'
'2009-06-15T13:45:30 -> juni 2009 (da-DK)
'
'2009-06-15T13:45:30 -> Juni 2009 (id-ID)
'Jedes andere einzelne Zeichen   Unbekannter Bezeichner.     Löst eine FormatException zur Laufzeit aus.
