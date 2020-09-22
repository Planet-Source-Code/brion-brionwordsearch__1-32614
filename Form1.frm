VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Brion's Word Search"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "Mystery"
      Height          =   600
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton Command13 
      Caption         =   "3   Vowels     In A Row"
      Height          =   600
      Left            =   1560
      TabIndex        =   14
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Calc Value"
      Height          =   255
      Left            =   1560
      TabIndex        =   38
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Clear List"
      Height          =   600
      Left            =   1560
      TabIndex        =   23
      Top             =   6360
      Width           =   1200
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Q without QU"
      Height          =   600
      Left            =   1560
      TabIndex        =   12
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Length by Letter(s)"
      Height          =   600
      Left            =   1560
      TabIndex        =   16
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Suffix Search"
      Height          =   600
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Prefix Search"
      Height          =   600
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Short to Long"
      Height          =   500
      Left            =   9240
      TabIndex        =   27
      Top             =   6480
      Width           =   950
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Z to A"
      Height          =   500
      Left            =   9240
      TabIndex        =   25
      Top             =   5280
      Width           =   950
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Long to Short"
      Height          =   500
      Left            =   9240
      TabIndex        =   26
      Top             =   5880
      Width           =   950
   End
   Begin VB.CommandButton Command18 
      Caption         =   "A to Z"
      Height          =   500
      Left            =   9240
      TabIndex        =   24
      Tag             =   "1"
      Top             =   4680
      Width           =   950
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Exit"
      Height          =   600
      Left            =   120
      TabIndex        =   22
      Top             =   6360
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2280
      List            =   "Form1.frx":0019
      TabIndex        =   21
      Text            =   "2"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Bingo!"
      Height          =   600
      Left            =   1560
      TabIndex        =   18
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Unscrambler"
      Height          =   600
      Left            =   1560
      TabIndex        =   20
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Wildcard Search"
      Height          =   600
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Word Search"
      Height          =   600
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton Command10 
      Caption         =   "10+"
      Height          =   300
      Left            =   9240
      TabIndex        =   36
      Top             =   3360
      Width           =   950
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   35
      Top             =   3000
      Width           =   950
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   34
      Top             =   2640
      Width           =   950
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   33
      Top             =   2280
      Width           =   950
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   32
      Top             =   1920
      Width           =   950
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   31
      Top             =   1560
      Width           =   950
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   30
      Top             =   1200
      Width           =   950
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   28
      Top             =   480
      Width           =   950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "All"
      Height          =   300
      Left            =   9240
      TabIndex        =   37
      Top             =   3720
      Width           =   950
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3 Letters"
      Height          =   300
      Left            =   9240
      TabIndex        =   29
      Top             =   840
      Width           =   950
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Columns         =   1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6540
      ItemData        =   "Form1.frx":0032
      Left            =   2880
      List            =   "Form1.frx":0034
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   6255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Knowledge is Power"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
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
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Sort:"
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
      Left            =   9240
      TabIndex        =   6
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "End Time:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Start Time:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Minimum Word Length: (For Unscrambler Only) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Searches:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Browse:"
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
      Left            =   9240
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Word List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Word Count:  "
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
      Left            =   3000
      TabIndex        =   5
      Top             =   7080
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Dim SourceFile As String 'File containing word list
Dim X As Long 'Subscript variable for word list, also counter for word list
Dim c As Long 'Global counter for words found
Dim sStringArray() As String ' Word list array
Dim sStringArraySL() As String 'Word list array sorted by letter in word

Dim HashArray() As Long 'Hash search array
Dim wildcard As String 'Currently set to ? to represent a blank
Dim Index() As Long 'Array to hold index numbers for unscrambling

Public Sub unscramble(ByRef sa() As String, Word As String, wl As Long, bingo As Long)
    Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
    Dim k As Long 'sStringArray() length holder
    Dim temp As String 'Temporary variable
    Dim match As Boolean 'Match flag
    Dim g As String 'Temporarily holds a letter value from sStringArray()
        
    For i = 0 To X
        k = Len(sa(i))
        If k <= wl Then
            j = 1
            match = False
            temp = Word
            Do While j <= k
                'check each letter from the word to see if it exists in the rack
                g = Mid$(sa(i), j, 1)
                If InStr(1, temp, g) <> 0 Then
                    match = True
                    'Replace the "used" letter
                    temp = Replace(temp, g, "", , 1)
                'If no letter match, see if there is a wildcard match
                ElseIf InStr(1, temp, wildcard) <> 0 Then
                    match = True
                    temp = Replace(temp, wildcard, "", , 1)
                'If no matches (If first time through, use index to skip ahead)
                Else
                    match = False
                    If j = 1 Then
                        Select Case g
                        Case "a"
                            i = Index(2)
                        Case "b"
                            i = Index(3)
                        Case "c"
                            i = Index(4)
                        Case "d"
                            i = Index(5)
                        Case "e"
                            i = Index(6)
                        Case "f"
                            i = Index(7)
                        Case "g"
                            i = Index(8)
                        Case "h"
                            i = Index(9)
                        Case "i"
                            i = Index(10)
                        Case "j"
                            i = Index(11)
                        Case "k"
                            i = Index(12)
                        Case "l"
                            i = Index(13)
                        Case "m"
                            i = Index(14)
                        Case "n"
                            i = Index(15)
                        Case "o"
                            i = Index(16)
                        Case "p"
                            i = Index(17)
                        Case "q"
                            i = Index(18)
                        Case "r"
                            i = Index(19)
                        Case "s"
                            i = Index(20)
                        Case "t"
                            i = Index(21)
                        Case "u"
                            i = Index(22)
                        Case "v"
                            i = Index(23)
                        Case "w"
                            i = Index(24)
                        Case "x"
                            i = Index(25)
                        Case "y"
                            i = Index(26)
                        Case Else
                        End Select
                    End If
                    Exit Do
                End If
                j = j + 1
            Loop
            If match = True Then
                If Len(sa(i)) >= bingo Then
                    List1.AddItem sa(i)
                    c = c + 1
                End If
            End If
        End If
     Next i
End Sub
Public Function CalcPoints(Word As String) As Long
'Calculates point values for words (no board position checking)
    Dim i As Long 'Loop counter
    Dim wl As Long 'Word length
    Dim value As Long 'Word value
    
    wl = Len(Word)
      
    For i = 1 To wl
        Select Case Mid$(Word, i, 1)
        Case "a", "e", "i", "l", "n", "o", "r", "s", "t", "u"
            value = value + 1
        Case "d", "g"
            value = value + 2
        Case "b", "c", "m", "p"
            value = value + 3
        Case "f", "h", "v", "w", "y"
            value = value + 4
        Case "k"
            value = value + 5
        Case "j", "x"
            value = value + 8
        Case "q", "z"
            value = value + 10
        Case Else
        End Select
    Next i
    CalcPoints = value
    
End Function

Public Function NumPossible(num As Long) As Long
'Calculates the number of possible letter combinations of a set of letters
    Dim i As Long 'Loop counter
    Dim factorial As Long 'Calculation variable
    Dim j As Long 'Loop counter
    Dim result As Long 'Calculation variable
        
    For i = 1 To num 'Loop
        factorial = num 'Reset calculation variable to num
        For j = num - 1 To i Step -1 'Step down loop
            factorial = factorial * j 'Calculate
        Next j
        result = result + factorial 'Save that value
    Next i
    result = result - num 'Subtract out the original number
    NumPossible = result 'Set the function value
End Function

Public Sub RemoveDuplicates(listbox As listbox)
'Removes duplicates from a listbox list
    Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
        
    For i = 0 To listbox.ListCount - 1 'Loop through list
        For j = i + 1 To listbox.ListCount - 1 'Loop through list
        If listbox.List(i) = listbox.List(j) Then 'Check for duplicates
            listbox.RemoveItem j 'Remove it
            j = j - 1 'Decrement j to start at same point in loop again
        End If
        Next j
    Next i
End Sub

Private Sub Command1_Click()
'Lists all words into the listbox from the word list
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 2
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        List1.AddItem sStringArray(i) 'Put them into the listbox
        c = c + 1 'Increment word counter
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command10_Click()
'Lists all words in list greater than 9 letters in length
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
           
    Command27.Enabled = True
    List1.Columns = 2
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) > 9 Then 'Check for word length greater than 9
            List1.AddItem sStringArray(i) 'If so, list in listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub



Private Sub Command11_Click()
'Searches array for match to input
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim Word As String 'Input from user
    Dim d As Long 'Binary search variable
        
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 2
    List1.Clear 'Clear list
    Word = InputBox("Enter word to search:  ") 'Get input
    If Word <> "" Then
        If Right$(Word, 3) = "   " Then
            List1.AddItem Word
        Else
            Label9.Caption = "Search Time:"
            Label9.Refresh
            Label6.Caption = "Start: " & Time
            Label6.Refresh
            Label7.Caption = "End: "
            Label7.Refresh
            Word = LCase(Word) 'Change to lower case
            'd = BinarySearch(sStringArray(), Word) 'Set the binary search variable
            d = HashSearch(sStringArray(), HashArray(), Word) 'Send the search function some data, get back subscript number (Array is passed by reference)
            Label7.Caption = "End: " & Time
            If d = -1 Then 'Check for no match
                Label1.Caption = "Word Count:  " & c 'Display word count
                MsgBox ("Sorry, " & Word & " is not a valid word.")
            Else
                List1.AddItem sStringArray(d) 'If there is a match, add to listbox
                c = c + 1 'Increment word count
                Label1.Caption = "Word Count:  " & c 'Display word count
            End If
        End If
    End If
    Label1.Caption = "Word Count:  " & List1.ListCount 'Display word count
    
End Sub

Private Sub Command12_Click()
'Wildcard search - finds all matches substuting any letter where there is a ? in input
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim Word As String 'Input from user
    Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
    Dim match As Long 'Word match count
    
    
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 3
    List1.Clear 'Clear Listbox
    Word = InputBox("Enter word form to search - Example:  a??h?b?t") 'Get input
    If Word <> "" Then
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    Word = LCase(Word) 'Change to lowercase
    For i = 0 To X 'Loop through all words in list
        match = 0 'Set match = to 0
        If Len(sStringArray(i)) = Len(Word) Then 'Check matching word length
            For j = 1 To Len(Word) 'Loop through all characters in input
                If Mid$(sStringArray(i), j, 1) = Mid$(Word, j, 1) Then 'Check each charater in input vs. current word in master list
                    match = match + 1 'If matches, increment match
                ElseIf Mid$(Word, j, 1) = wildcard Then 'If the current charater matches the wildcard
                    match = match + 1 'increment match
                End If
            Next j
            If match = Len(Word) Then 'Check if match matches the length of user input - if not, then match was not incremented above, meaning no character match
                List1.AddItem sStringArray(i) 'Add to listbox
                c = c + 1 'Increment word counter
            End If
        End If
    Next i
    Label7.Caption = "End: " & Time
    If c = 0 Then
        Label1.Caption = "Word Count:  " & c 'Display word count
        MsgBox ("Sorry!  No match found.") 'Check for no matches
    End If
    End If
    Label1.Caption = "Word Count:  " & c 'Display word count
End Sub

Private Sub Command13_Click()
    'Lists words with 3 vowels in a row
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim wl As Long
    Dim match As String
    Dim temp As String
    Dim hold As String
    Dim temp2() As String
    
    c = 0 'Word counter
        
    List1.Clear 'Clear list
    Command27.Enabled = True
    List1.Columns = 4
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    
    For i = 0 To X
        wl = Len(sStringArray(i))
        If wl > 2 And wl < 8 Then
            match = 0
            For j = 1 To wl
                temp = Mid$(sStringArray(i), j, 3)
                For k = 1 To Len(temp)
                    Select Case Mid$(temp, k, 1)
                    Case "a", "e", "i", "o", "u"
                        match = match + 1
                    Case Else
                    End Select
                Next k
                If match = 3 Then
                    'List1.AddItem sStringArray(i)
                    ReDim Preserve temp2(c)
                    temp2(c) = sStringArray(i)
                    c = c + 1
                    Exit For
                End If
                match = 0
            Next j
        End If
    Next i
    'bubble sort the results
    For i = 0 To c - 1
        For j = i + 1 To c - 1
            If Len(temp2(i)) > Len(temp2(j)) Then
                hold = temp2(i)
                temp2(i) = temp2(j)
                temp2(j) = hold
            End If
        Next j
    Next i
    List1.Clear
    For i = 0 To c - 1
        List1.AddItem temp2(i)
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
    
End Sub

Private Sub Command14_Click()
Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
    Dim k As Long 'sStringArray() length holder
    Dim Word As String 'Input from user
    Dim wl As Long 'Word length
    Dim temp As String 'Temporary variable
    Dim cNum As Long
    Dim match As Boolean 'Match flag
    'Dim starttime As Single
    'Dim endtime As Single
    'Dim totaltime As Single
    Dim g As String 'Temporarily holds a letter value from sStringArray()
    Dim h As String
    
    c = 0 'Word counter
    Command27.Enabled = True
    List1.Columns = 3
    If Combo1.Text > "9" Or Combo1.Text < "2" Then
        Combo1.Text = "2"
    End If
    cNum = Combo1.Text
    List1.Clear 'Clear listbox
    Word = InputBox("Enter your letters (use ? to represent a blank):  ") 'Get input
    If Word <> "" Then
    
    'starttime = Timer
    
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    Word = LCase(Word) 'Change it to lower case
    wl = Len(Word) 'Set word length variable
    ReDim w(1 To wl)
        If wl = 0 Then
            'MsgBox ("Enter at least 2 letters.")
        ElseIf wl = 1 Then
            MsgBox ("Enter more than one letter.")
        ElseIf wl > 26 Then
            MsgBox ("Too many letters!")
        Else
        Call unscramble(sStringArray(), Word, wl, cNum)
        End If
    End If
    Label7.Caption = "End: " & Time
    
    'endtime = Timer
    'totaltime = endtime - starttime
    'Label9.Caption = "Search Time: " & Format(totaltime, "##.####") + " sec"
    
    Label1.Caption = "Word Count:  " & c 'Display word count
    If wl > 26 Then
    ElseIf c < 1 And wl > 1 Then
        MsgBox ("You can make NO words!!") 'If no possible words can be made
    End If
    Combo1.Text = "2"
    
End Sub

Private Sub Command15_Click()
'Mystery Button
   
End Sub

Private Sub Command16_Click()
'Checks for a Scrabble Bingo - use all seven letter
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    Dim Word As String 'Input from user
    
    Dim wl As Long 'Word length
    
    Dim w() As String
    
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 3
    List1.Clear 'Clear listbox
    Word = InputBox("Enter 7, 8 or 9 letters:  ") 'Get input
    If Word <> "" Then
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    Word = LCase(Word) 'Change to lowercase
    wl = Len(Word) 'Set word length variable
    ReDim w(1 To wl)
    For i = 1 To wl
        w(i) = Mid$(Word, i, 1)
    Next i
        If wl = 0 Then
            'Do nothing
        ElseIf wl < 7 Then
            MsgBox ("Enter 7, 8 or 9 letters.")
        ElseIf wl > 9 Then
            MsgBox ("Enter 7, 8 or 9 letters.")
        Else
            Call unscramble(sStringArray(), Word, wl, 7)
        End If
        Label7.Caption = "End: " & Time
        Label1.Caption = "Word Count:  " & List1.ListCount 'Display word count
        If c < 1 Then MsgBox ("No Bingo possible using " & Word & ".") 'If no match
    End If
    
    Label1.Caption = "Word Count:  " & List1.ListCount 'Display word count
    
End Sub

Private Sub Command17_Click()
    'Exits the program
    Unload Me
    End
End Sub

Private Sub Command18_Click()
'Ascending sorting code
Dim i As Long
Dim temp() As String

If List1.ListCount <> 0 Then
    If c < 32000 Then
        ReDim temp(c - 1)
        For i = 0 To c - 1
            temp(i) = List1.List(i)
        Next i

        Call TriQuickSortString(temp(), SortAscending)
        List1.Clear
        For i = 0 To c - 1
            List1.AddItem temp(i)
        Next i
    Else
        MsgBox ("Too many words to sort!  Narrow your search.")
    End If
End If
End Sub

Private Sub Command19_Click()
'sorting by number code - Big to Little
Dim i As Long
Dim j As Long
Dim temp() As String
Dim hold As String
Dim Y As Long
Dim z As Long

If List1.ListCount <> 0 Then
    If c < 4000 Then
        ReDim temp(c - 1)
        For i = 0 To c - 1
            temp(i) = List1.List(i)
        Next i
        For i = 0 To c - 1
            For j = i + 1 To c - 1
                If Len(temp(i)) < Len(temp(j)) Then
                    hold = temp(i)
                    temp(i) = temp(j)
                    temp(j) = hold
                End If
            Next j
        Next i
        List1.Clear
        For i = 0 To c - 1
            List1.AddItem temp(i)
        Next i
    Else
        MsgBox ("Too many words!  Narrow your search.")
    End If
End If
End Sub



Private Sub Command2_Click()
'Lists all two letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 6
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 2 Then 'Check for word length of 2
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command20_Click()
'Descending sorting code
Dim i As Long
Dim temp() As String

If List1.ListCount <> 0 Then
    If c < 32000 Then
        ReDim temp(c - 1)
        For i = 0 To c - 1
            temp(i) = List1.List(i)
        Next i

        Call TriQuickSortString(temp(), SortDescending)
        List1.Clear
        For i = 0 To c - 1
            List1.AddItem temp(i)
        Next i
    Else
        MsgBox ("Too many words to sort!  Narrow your search.")
    End If
End If
End Sub

Private Sub Command21_Click()
'sorting by number code - Little to Big
Dim i As Long
Dim j As Long
Dim temp() As String
Dim hold As String
Dim Y As Long
Dim z As Long

If List1.ListCount <> 0 Then
    If c < 4000 Then
        ReDim temp(c - 1)
        For i = 0 To c - 1
            temp(i) = List1.List(i)
        Next i
        For i = 0 To c - 1
            For j = i + 1 To c - 1
                If Len(temp(i)) > Len(temp(j)) Then
                    hold = temp(i)
                    temp(i) = temp(j)
                    temp(j) = hold
                End If
            Next j
        Next i
        List1.Clear
        For i = 0 To c - 1
            List1.AddItem temp(i)
        Next i
    Else
        MsgBox ("Too many words!  Narrow your search.")
    End If
End If
End Sub

Private Sub Command22_Click()
'Wildcard beginning search - Ex - All words that begin with ??y or l?g
    Dim Word As String 'Input from user
    Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
    Dim match As Long 'Word match count
    
    
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 2
    List1.Clear 'Clear Listbox
    Word = InputBox("Enter word beginning to search (Use ? to represent blank):") 'Get input
    If Word <> "" Then
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    Word = LCase(Word) 'Change to lowercase
    For i = 0 To X 'Loop through all words in list
        match = 0 'Set match = to 0
        'If Len(sStringArray(i)) = Len(Word) Then 'Check matching word length
            For j = 1 To Len(Word) 'Loop through all characters in input
                If Mid$(sStringArray(i), j, 1) = Mid$(Word, j, 1) Then 'Check each charater in input vs. current word in master list
                    match = match + 1 'If matches, increment match
                ElseIf Mid$(Word, j, 1) = wildcard Then 'If the current charater matches the wildcard
                    match = match + 1 'increment match
                End If
            Next j
            If match = Len(Word) Then 'Check if match matches the length of user input - if not, then match was not incremented above, meaning no character match
                List1.AddItem sStringArray(i) 'Add to listbox
                c = c + 1 'Increment word counter
            End If
        'End If
    Next i
    Label7.Caption = "End: " & Time
    If c = 0 Then
        Label1.Caption = "Word Count:  " & c 'Display word count
        MsgBox ("Sorry!  No match found.") 'Check for no matches
    End If
    End If
    Label1.Caption = "Word Count:  " & c 'Display word count
End Sub

Private Sub Command23_Click()
'Wildcard ending search - Ex - All words that end with k?? or ?r?r
    Dim Word As String 'Input from user
    Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
    Dim match As Long 'Word match count
    Dim ending As String 'Word ending
    
    Dim wl As Long
  
    
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 2
    List1.Clear 'Clear Listbox
    Word = InputBox("Enter word beginning to search - (Use ? to represent blank):") 'Get input
    If Word <> "" Then
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    Word = LCase(Word) 'Change to lowercase
    wl = Len(Word)
    For i = 0 To X 'Loop through all words in list
        match = 0 'Set match = to 0
        ending = Right$(sStringArray(i), wl)
        'If Len(sStringArray(i)) = Len(Word) Then 'Check matching word length
            For j = 1 To wl 'Loop through all characters in input
                If Mid$(ending, j, 1) = Mid$(Word, j, 1) Then 'Check each charater in input vs. current word in master list
                    match = match + 1 'If matches, increment match
                ElseIf Mid$(Word, j, 1) = wildcard Then 'If the current charater matches the wildcard
                    match = match + 1 'increment match
                End If
            Next j
            If match = wl Then 'Check if match matches the length of user input - if not, then match was not incremented above, meaning no character match
                List1.AddItem sStringArray(i) 'Add to listbox
                c = c + 1 'Increment word counter
            End If
        'End If
    Next i
    Label7.Caption = "End: " & Time
    If c = 0 Then
        Label1.Caption = "Word Count:  " & c 'Display word count
        MsgBox ("Sorry!  No match found.") 'Check for no matches
    End If
    End If
    Label1.Caption = "Word Count:  " & c 'Display word count
End Sub

Private Sub Command24_Click()
'List all words of a length that have a certain letter
'Ex - All 4 letter words that have a z in any place
    Dim Word As String 'Input from user
    Dim swl As String 'Input from user
    Dim wl As Long
    Dim i As Long 'Loop counter
    Dim j As Long 'Loop counter
    Dim lw As Long
            
    c = 0 'Word counter
    
    Command27.Enabled = True
    List1.Columns = 2
    List1.Clear 'Clear Listbox
    Word = InputBox("Enter letter or letters search: ") 'Get input
    If Word <> "" Then
    swl = InputBox("Enter word length to check (Enter 9 for all words): ", , 9)
    If swl <> "" And swl > "1" And swl <= "9" Then
    wl = swl
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    Word = LCase(Word) 'Change to lowercase
    lw = Len(Word)
    For i = 0 To X 'Loop through all words in list
        If wl <> 9 Then
            If Len(sStringArray(i)) = wl Then 'Check matching word length
                For j = 1 To wl 'Loop through all characters
                    If Mid$(sStringArray(i), j, lw) = Word Then 'Check each charater in input vs. current word in master list
                        List1.AddItem sStringArray(i) 'Add to listbox
                        c = c + 1 'Increment word counter
                    End If
                Next j
            End If
        Else
            For j = 1 To wl 'Loop through all characters
                If Mid$(sStringArray(i), j, lw) = Word Then 'Check each charater in input vs. current word in master list
                    List1.AddItem sStringArray(i) 'Add to listbox
                    c = c + 1 'Increment word counter
                End If
            Next j
        End If
    Next i
    Label7.Caption = "End: " & Time
    If c = 0 Then
        Label1.Caption = "Word Count:  " & c 'Display word count
        MsgBox ("Sorry!  No match found.") 'Check for no matches
    End If
    End If
    End If
    Label1.Caption = "Word Count:  " & c 'Display word count
End Sub

Private Sub Command25_Click()
'List all words that use Q without U
    Dim i As Long 'Loop counter
    Dim match As Long
    Dim j As Long
    Dim k As Long
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 3
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X
        match = 0
        For j = 1 To Len(sStringArray(i))
            If Mid$(sStringArray(i), j, 1) = "q" Then
                For k = 1 To Len(sStringArray(i))
                    If Mid$(sStringArray(i), k, 2) <> "qu" Then
                        match = match + 1
                    End If
                Next k
                If match = Len(sStringArray(i)) Then
                    List1.AddItem sStringArray(i)
                    c = c + 1
                End If
            End If
        Next j
        
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command26_Click()
    List1.Clear
    c = 0
    Label1.Caption = "Word Count:  " & c 'Display word count
End Sub

Private Sub Command27_Click()
'Calculates word values for what is in List1
    Dim i As Long 'Loop counter
    Dim vTemp() As Long
    Dim wtemp() As String
    Dim lc As Long
    
    If c < 31000 Then
        lc = List1.ListCount
        ReDim vTemp(lc)
        ReDim wtemp(lc)
        For i = 0 To lc - 1
            wtemp(i) = List1.List(i)
        Next i
        For i = 0 To lc - 1
            vTemp(i) = CalcPoints(wtemp(i))
        Next i
        List1.Clear
        If List1.Columns > 1 Then
            List1.Columns = List1.Columns - 1
        Else
            List1.Columns = 1
        End If
        For i = 0 To lc - 1
            List1.AddItem wtemp(i) & " - " & vTemp(i)
        Next i
    Else
        MsgBox ("Too many words! Narrow your search.")
    End If
    Command27.Enabled = False
End Sub

Private Sub Command3_Click()
'Lists all three letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 6
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 3 Then 'Check for word length of 3
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command4_Click()
'Lists all four letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 5
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 4 Then 'Check for word length of 4
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command5_Click()
'Lists all five letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 5
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 5 Then 'Check for word length of 5
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command6_Click()
'Lists all six letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 4
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 6 Then 'Check for word length of 6
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command7_Click()
'Lists all seven letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 4
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 7 Then 'Check for word length of 7
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command8_Click()
'Lists all eight letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
        
    Command27.Enabled = True
    List1.Columns = 3
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 8 Then 'Check for word length of 8
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Command9_Click()
'Lists all nine letter words in word array
'Note - had to use a counter variable, Listbox.ListCount must use integer not long
    Dim i As Long 'Loop counter
    
    c = 0 'Word counter
       
    Command27.Enabled = True
    List1.Columns = 3
    Label9.Caption = "Search Time:"
    Label9.Refresh
    Label6.Caption = "Start: " & Time
    Label6.Refresh
    Label7.Caption = "End: "
    Label7.Refresh
    List1.Clear 'Clear list
    For i = 0 To X 'Loop through all words
        If Len(sStringArray(i)) = 9 Then 'Check for word length of 9
            List1.AddItem sStringArray(i) 'If so, add word to listbox
            c = c + 1 'Increment word counter
        End If
    Next i
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
End Sub

Private Sub Form_Load()
'Loads the form, opens file, reads words into an array
'Assumes sorted file
    Dim temp As String
    Dim temp2 As String
    Dim V As String
    
    wildcard = "?"
    Load Form2
    Form2.Show
    DoEvents
    Label9.Caption = "Load Time:"
    Label6.Caption = "Start: " & Time
    SourceFile = App.Path & "\enable3.txt" 'Source file
    X = 0 'Set subscript/word counter to 0
    Open SourceFile For Input As #1 'Open source file
        Do Until EOF(1) 'Loop until end of file
            Input #1, temp 'Read from file into array
            If Len(temp) > 1 Then 'And Len(temp) < 15 Then (Use this part if limiting words for Scrabble)
                ReDim Preserve sStringArray(X) 'Redimension word array so can use dynamic list
                sStringArray(X) = temp 'Assign word into array
                X = X + 1 'Icrement subscript
            End If
        Loop
    Close #1 'Close file
    X = X - 1
    Call TriQuickSortString(sStringArray(), SortAscending)
    Call BuildHashTable(sStringArray(), HashArray())
    Call CreateIndex(sStringArray())
    'Call SortWord(sStringArray())
    Label1.Caption = "Word Count:  " & c 'Display word count
    Label7.Caption = "End: " & Time
    Unload Form2
    Me.Show
    Command11.SetFocus
    Command15.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
'Unloads form
    
    Unload Me
    End
End Sub


Private Sub SortWord(ByRef sa() As String)
'Sorts each word by letter, currently not used in this program
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim wl As Long
    Dim temp() As String
    Dim hold As String
    
    ReDim sStringArraySL(X)
    For i = 0 To X
        wl = Len(sStringArray(i))
        ReDim temp(1 To wl)
        For j = 1 To wl
            temp(j) = Mid$(sStringArray(i), j, 1)
        Next j
        For j = 1 To wl
            For k = j + 1 To wl
            If temp(j) > temp(k) Then
                    hold = temp(j)
                    temp(j) = temp(k)
                    temp(k) = hold
                End If
            Next k
        Next j
        For j = 1 To wl
            sStringArraySL(i) = sStringArraySL(i) & temp(j)
        Next j
    Next i
End Sub

Private Sub CreateIndex(ByRef sa() As String)
'Creates index for skipping ahead when unscrambling
    Dim i As Long
    Dim j As Long
    Dim V As String
    Dim t As String
            
    ReDim Index(1 To 26)
    Index(1) = 0
    t = "a"
    j = 2
    For i = 0 To X
        V = Left$(sa(i), 1)
        If t <> V Then
            Index(j) = i
            j = j + 1
            t = V
        End If
    Next i
End Sub

