Attribute VB_Name = "Public"
Option Explicit


Public X As Long 'Subscript variable for word list, also counter for word list
Public c As Long 'Global counter for words found
Public sStringArray() As String ' Word list array
Public sStringArraySL() As String
Public def As String
Public HashArray() As Long 'Hash search array
Public wildcard As String
Public Index() As Long


            
Public Function wordsearch(tempword() As String) As Boolean
    Dim i As Long
    Dim d As Long
    Dim al As Long
    Dim check() As Boolean
    Dim temp As String
    Dim temp2 As String
    
    wordsearch = True
    temp = ""
    al = UBound(tempword())
    ReDim check(al)
    For i = 1 To al
        tempword(i) = LCase(tempword(i)) 'Change to lower case
    Next i
    For i = 1 To al
        d = HashSearch(sStringArray(), HashArray(), tempword(i)) 'Send the search function some data, get back subscript number (Array is passed by reference)
        If d = -1 Then 'Check for no match
            check(i) = False
            wordsearch = False
        Else
            check(i) = True
        End If
    Next i
    For i = 1 To al
        If check(i) = False Then
            temp = temp & tempword(i) & " "
        Else
            temp2 = temp2 & tempword(i) & " "
        End If
    Next i
    'If temp <> "" Then
        MsgBox ("Non-Valid Word(s): " & temp & vbNewLine & "Valid Word(s): " & temp2)
    'End If
End Function
            
