Attribute VB_Name = "AIMPass"
Public Function GetCase(CurrentPos As Integer, TempLetter As String) As Integer
    'This algorithm determines if the given
    'charactor in the given position is one
    'of the following:
    '1. upper case A to O
    '2. upper case P to Z
    '3. lower case A to O
    '4. lower case P to Z
    '5. number
    'The translation is based on current
    'position in the encrypted string and
    'the 'Code' charactor.  These two
    'combined output 1 of the 5 values
    'listed above.
    Dim TheCase As Integer
    Select Case CurrentPos
        Case 1, 6
            Select Case TempLetter
                Case "A"
                    TheCase = 1
                Case "B"
                    TheCase = 2
                Case "C"
                    TheCase = 3
                Case "D"
                    TheCase = 4
                Case "H"
                    TheCase = 5
            End Select
        Case 2, 7
            Select Case TempLetter
                Case "M"
                    TheCase = 1
                Case "N"
                    TheCase = 2
                Case "O"
                    TheCase = 3
                Case "P"
                    TheCase = 4
                Case "L"
                    TheCase = 5
            End Select
        Case 3, 8
            Select Case TempLetter
                Case "E"
                    TheCase = 1
                Case "F"
                    TheCase = 2
                Case "G"
                    TheCase = 3
                Case "H"
                    TheCase = 4
                Case "D"
                    TheCase = 5
            End Select
        Case 4, 9
            Select Case TempLetter
                Case "F"
                    TheCase = 1
                Case "E"
                    TheCase = 2
                Case "H"
                    TheCase = 3
                Case "G"
                    TheCase = 4
                Case "C"
                    TheCase = 5
            End Select
        Case 5, 10
            Select Case TempLetter
                Case "G"
                    TheCase = 1
                Case "H"
                    TheCase = 2
                Case "E"
                    TheCase = 3
                Case "F"
                    TheCase = 4
                Case "B"
                    TheCase = 5
            End Select
        Case 11
            Select Case TempLetter
                Case "D"
                    TheCase = 1
                Case "A"
                    TheCase = 2
                Case "D"
                    TheCase = 3
                Case "C"
                    TheCase = 4
                Case "G"
                    TheCase = 5
            End Select
        Case 12
            Select Case TempLetter
                Case "P"
                    TheCase = 1
                Case "O"
                    TheCase = 2
                Case "N"
                    TheCase = 3
                Case "M"
                    TheCase = 4
                Case "I"
                    TheCase = 5
            End Select
        Case 13
            Select Case TempLetter
                Case "D"
                    TheCase = 1
                Case "C"
                    TheCase = 2
                Case "B"
                    TheCase = 3
                Case "A"
                    TheCase = 4
                Case "E"
                    TheCase = 5
            End Select
        Case 14
            Select Case TempLetter
                Case "K"
                    TheCase = 1
                Case "L"
                    TheCase = 2
                Case "I"
                    TheCase = 3
                Case "J"
                    TheCase = 4
                Case "N"
                    TheCase = 5
            End Select
        Case 15
            Select Case TempLetter
                Case "J"
                    TheCase = 1
                Case "I"
                    TheCase = 2
                Case "L"
                    TheCase = 3
                Case "K"
                    TheCase = 4
                Case "O"
                    TheCase = 5
            End Select
        Case 16
            Select Case TempLetter
                Case "O"
                    TheCase = 1
                Case "P"
                    TheCase = 2
                Case "M"
                    TheCase = 3
                Case "N"
                    TheCase = 4
                Case "J"
                    TheCase = 5
            End Select
    End Select
    GetCase = TheCase
End Function

Public Function TranslateLetter(CurrentPos As Integer, CurChar As String, TheCase As Integer)
    'This is the big algorithm.  It's so big because
    'it is a triple nested Select Case statement.
    'First, we look at where the current letter is
    'in the string.  The algorithm allows for up to
    '16 letter screennames.  Then we need to do the
    '5 cases.  We only need 3 statements here since
    'we won't bother to get the case of the letter
    'until afterwards.  It's a good thing or this
    'Function would be almost twice as long.  And
    'finally, we match up the encrypted letter
    'with the real thing.  So, for each of the 5
    'cases (shortened to 3) there is 36 possible
    'characters.  The total amount of possible
    'values is 16 x 36.  576.  Yes, it took me a
    'long time to write this.
    Select Case CurrentPos
        Case 1
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "D"
                            TranslateLetter = "a"
                        Case "A"
                            TranslateLetter = "b"
                        Case "B"
                            TranslateLetter = "c"
                        Case "G"
                            TranslateLetter = "d"
                        Case "H"
                            TranslateLetter = "e"
                        Case "E"
                            TranslateLetter = "f"
                        Case "F"
                            TranslateLetter = "g"
                        Case "K"
                            TranslateLetter = "h"
                        Case "L"
                            TranslateLetter = "i"
                        Case "I"
                            TranslateLetter = "j"
                        Case "J"
                            TranslateLetter = "k"
                        Case "O"
                            TranslateLetter = "l"
                        Case "P"
                            TranslateLetter = "m"
                        Case "M"
                            TranslateLetter = "n"
                        Case "N"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "C"
                            TranslateLetter = "p"
                        Case "D"
                            TranslateLetter = "q"
                        Case "A"
                            TranslateLetter = "r"
                        Case "B"
                            TranslateLetter = "s"
                        Case "G"
                            TranslateLetter = "t"
                        Case "H"
                            TranslateLetter = "u"
                        Case "E"
                            TranslateLetter = "v"
                        Case "F"
                            TranslateLetter = "w"
                        Case "K"
                            TranslateLetter = "x"
                        Case "L"
                            TranslateLetter = "y"
                        Case "I"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "C"
                            TranslateLetter = "0"
                        Case "D"
                            TranslateLetter = "1"
                        Case "A"
                            TranslateLetter = "2"
                        Case "B"
                            TranslateLetter = "3"
                        Case "G"
                            TranslateLetter = "4"
                        Case "H"
                            TranslateLetter = "5"
                        Case "E"
                            TranslateLetter = "6"
                        Case "F"
                            TranslateLetter = "7"
                        Case "K"
                            TranslateLetter = "8"
                        Case "L"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 2
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "F"
                            TranslateLetter = "a"
                        Case "G"
                            TranslateLetter = "b"
                        Case "H"
                            TranslateLetter = "c"
                        Case "A"
                            TranslateLetter = "d"
                        Case "B"
                            TranslateLetter = "e"
                        Case "C"
                            TranslateLetter = "f"
                        Case "D"
                            TranslateLetter = "g"
                        Case "M"
                            TranslateLetter = "h"
                        Case "N"
                            TranslateLetter = "i"
                        Case "O"
                            TranslateLetter = "j"
                        Case "P"
                            TranslateLetter = "k"
                        Case "I"
                            TranslateLetter = "l"
                        Case "J"
                            TranslateLetter = "m"
                        Case "K"
                            TranslateLetter = "n"
                        Case "L"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "E"
                            TranslateLetter = "p"
                        Case "F"
                            TranslateLetter = "q"
                        Case "G"
                            TranslateLetter = "r"
                        Case "H"
                            TranslateLetter = "s"
                        Case "A"
                            TranslateLetter = "t"
                        Case "B"
                            TranslateLetter = "u"
                        Case "C"
                            TranslateLetter = "v"
                        Case "D"
                            TranslateLetter = "w"
                        Case "M"
                            TranslateLetter = "x"
                        Case "N"
                            TranslateLetter = "y"
                        Case "O"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "E"
                            TranslateLetter = "0"
                        Case "F"
                            TranslateLetter = "1"
                        Case "G"
                            TranslateLetter = "2"
                        Case "H"
                            TranslateLetter = "3"
                        Case "A"
                            TranslateLetter = "4"
                        Case "B"
                            TranslateLetter = "5"
                        Case "C"
                            TranslateLetter = "6"
                        Case "D"
                            TranslateLetter = "7"
                        Case "M"
                            TranslateLetter = "8"
                        Case "N"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 3
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "J"
                            TranslateLetter = "a"
                        Case "K"
                            TranslateLetter = "b"
                        Case "L"
                            TranslateLetter = "c"
                        Case "M"
                            TranslateLetter = "d"
                        Case "N"
                            TranslateLetter = "e"
                        Case "O"
                            TranslateLetter = "f"
                        Case "P"
                            TranslateLetter = "g"
                        Case "A"
                            TranslateLetter = "h"
                        Case "B"
                            TranslateLetter = "i"
                        Case "C"
                            TranslateLetter = "j"
                        Case "D"
                            TranslateLetter = "k"
                        Case "E"
                            TranslateLetter = "l"
                        Case "F"
                            TranslateLetter = "m"
                        Case "G"
                            TranslateLetter = "n"
                        Case "H"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "I"
                            TranslateLetter = "p"
                        Case "J"
                            TranslateLetter = "q"
                        Case "K"
                            TranslateLetter = "r"
                        Case "L"
                            TranslateLetter = "s"
                        Case "M"
                            TranslateLetter = "t"
                        Case "N"
                            TranslateLetter = "u"
                        Case "O"
                            TranslateLetter = "v"
                        Case "P"
                            TranslateLetter = "w"
                        Case "A"
                            TranslateLetter = "x"
                        Case "B"
                            TranslateLetter = "y"
                        Case "C"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "I"
                            TranslateLetter = "0"
                        Case "J"
                            TranslateLetter = "1"
                        Case "K"
                            TranslateLetter = "2"
                        Case "L"
                            TranslateLetter = "3"
                        Case "M"
                            TranslateLetter = "4"
                        Case "N"
                            TranslateLetter = "5"
                        Case "O"
                            TranslateLetter = "6"
                        Case "P"
                            TranslateLetter = "7"
                        Case "A"
                            TranslateLetter = "8"
                        Case "B"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 4
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "B"
                            TranslateLetter = "a"
                        Case "C"
                            TranslateLetter = "b"
                        Case "D"
                            TranslateLetter = "c"
                        Case "E"
                            TranslateLetter = "d"
                        Case "F"
                            TranslateLetter = "e"
                        Case "G"
                            TranslateLetter = "f"
                        Case "H"
                            TranslateLetter = "g"
                        Case "I"
                            TranslateLetter = "h"
                        Case "J"
                            TranslateLetter = "i"
                        Case "K"
                            TranslateLetter = "j"
                        Case "L"
                            TranslateLetter = "k"
                        Case "M"
                            TranslateLetter = "l"
                        Case "N"
                            TranslateLetter = "m"
                        Case "O"
                            TranslateLetter = "n"
                        Case "P"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "A"
                            TranslateLetter = "p"
                        Case "B"
                            TranslateLetter = "q"
                        Case "C"
                            TranslateLetter = "r"
                        Case "D"
                            TranslateLetter = "s"
                        Case "E"
                            TranslateLetter = "t"
                        Case "F"
                            TranslateLetter = "u"
                        Case "G"
                            TranslateLetter = "v"
                        Case "H"
                            TranslateLetter = "w"
                        Case "I"
                            TranslateLetter = "x"
                        Case "J"
                            TranslateLetter = "y"
                        Case "K"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "A"
                            TranslateLetter = "0"
                        Case "B"
                            TranslateLetter = "1"
                        Case "C"
                            TranslateLetter = "2"
                        Case "D"
                            TranslateLetter = "3"
                        Case "E"
                            TranslateLetter = "4"
                        Case "F"
                            TranslateLetter = "5"
                        Case "G"
                            TranslateLetter = "6"
                        Case "H"
                            TranslateLetter = "7"
                        Case "I"
                            TranslateLetter = "8"
                        Case "J"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 5
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "A"
                            TranslateLetter = "a"
                        Case "D"
                            TranslateLetter = "b"
                        Case "C"
                            TranslateLetter = "c"
                        Case "F"
                            TranslateLetter = "d"
                        Case "E"
                            TranslateLetter = "e"
                        Case "H"
                            TranslateLetter = "f"
                        Case "G"
                            TranslateLetter = "g"
                        Case "J"
                            TranslateLetter = "h"
                        Case "I"
                            TranslateLetter = "i"
                        Case "L"
                            TranslateLetter = "j"
                        Case "K"
                            TranslateLetter = "k"
                        Case "N"
                            TranslateLetter = "l"
                        Case "M"
                            TranslateLetter = "m"
                        Case "P"
                            TranslateLetter = "n"
                        Case "O"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "B"
                            TranslateLetter = "p"
                        Case "A"
                            TranslateLetter = "q"
                        Case "D"
                            TranslateLetter = "r"
                        Case "C"
                            TranslateLetter = "s"
                        Case "F"
                            TranslateLetter = "t"
                        Case "E"
                            TranslateLetter = "u"
                        Case "H"
                            TranslateLetter = "v"
                        Case "G"
                            TranslateLetter = "w"
                        Case "J"
                            TranslateLetter = "x"
                        Case "I"
                            TranslateLetter = "y"
                        Case "L"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "B"
                            TranslateLetter = "0"
                        Case "A"
                            TranslateLetter = "1"
                        Case "D"
                            TranslateLetter = "2"
                        Case "C"
                            TranslateLetter = "3"
                        Case "F"
                            TranslateLetter = "4"
                        Case "E"
                            TranslateLetter = "5"
                        Case "H"
                            TranslateLetter = "6"
                        Case "G"
                            TranslateLetter = "7"
                        Case "J"
                            TranslateLetter = "8"
                        Case "I"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 6
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "D"
                            TranslateLetter = "a"
                        Case "A"
                            TranslateLetter = "b"
                        Case "B"
                            TranslateLetter = "c"
                        Case "G"
                            TranslateLetter = "d"
                        Case "H"
                            TranslateLetter = "e"
                        Case "E"
                            TranslateLetter = "f"
                        Case "F"
                            TranslateLetter = "g"
                        Case "K"
                            TranslateLetter = "h"
                        Case "L"
                            TranslateLetter = "i"
                        Case "I"
                            TranslateLetter = "j"
                        Case "J"
                            TranslateLetter = "k"
                        Case "O"
                            TranslateLetter = "l"
                        Case "P"
                            TranslateLetter = "m"
                        Case "M"
                            TranslateLetter = "n"
                        Case "N"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "C"
                            TranslateLetter = "p"
                        Case "D"
                            TranslateLetter = "q"
                        Case "A"
                            TranslateLetter = "r"
                        Case "B"
                            TranslateLetter = "s"
                        Case "G"
                            TranslateLetter = "t"
                        Case "H"
                            TranslateLetter = "u"
                        Case "E"
                            TranslateLetter = "v"
                        Case "F"
                            TranslateLetter = "w"
                        Case "K"
                            TranslateLetter = "x"
                        Case "L"
                            TranslateLetter = "y"
                        Case "I"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "C"
                            TranslateLetter = "0"
                        Case "D"
                            TranslateLetter = "1"
                        Case "A"
                            TranslateLetter = "2"
                        Case "B"
                            TranslateLetter = "3"
                        Case "G"
                            TranslateLetter = "4"
                        Case "H"
                            TranslateLetter = "5"
                        Case "E"
                            TranslateLetter = "6"
                        Case "F"
                            TranslateLetter = "7"
                        Case "K"
                            TranslateLetter = "8"
                        Case "L"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 7
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "E"
                            TranslateLetter = "a"
                        Case "H"
                            TranslateLetter = "b"
                        Case "G"
                            TranslateLetter = "c"
                        Case "B"
                            TranslateLetter = "d"
                        Case "A"
                            TranslateLetter = "e"
                        Case "D"
                            TranslateLetter = "f"
                        Case "C"
                            TranslateLetter = "g"
                        Case "N"
                            TranslateLetter = "h"
                        Case "M"
                            TranslateLetter = "i"
                        Case "P"
                            TranslateLetter = "j"
                        Case "O"
                            TranslateLetter = "k"
                        Case "J"
                            TranslateLetter = "l"
                        Case "I"
                            TranslateLetter = "m"
                        Case "L"
                            TranslateLetter = "n"
                        Case "K"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "F"
                            TranslateLetter = "p"
                        Case "E"
                            TranslateLetter = "q"
                        Case "H"
                            TranslateLetter = "r"
                        Case "G"
                            TranslateLetter = "s"
                        Case "B"
                            TranslateLetter = "t"
                        Case "A"
                            TranslateLetter = "u"
                        Case "D"
                            TranslateLetter = "v"
                        Case "C"
                            TranslateLetter = "w"
                        Case "N"
                            TranslateLetter = "x"
                        Case "M"
                            TranslateLetter = "y"
                        Case "P"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "F"
                            TranslateLetter = "0"
                        Case "E"
                            TranslateLetter = "1"
                        Case "H"
                            TranslateLetter = "2"
                        Case "G"
                            TranslateLetter = "3"
                        Case "B"
                            TranslateLetter = "4"
                        Case "A"
                            TranslateLetter = "5"
                        Case "D"
                            TranslateLetter = "6"
                        Case "C"
                            TranslateLetter = "7"
                        Case "M"
                            TranslateLetter = "8"
                        Case "N"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 7
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "K"
                            TranslateLetter = "a"
                        Case "J"
                            TranslateLetter = "b"
                        Case "I"
                            TranslateLetter = "c"
                        Case "P"
                            TranslateLetter = "d"
                        Case "O"
                            TranslateLetter = "e"
                        Case "N"
                            TranslateLetter = "f"
                        Case "M"
                            TranslateLetter = "g"
                        Case "D"
                            TranslateLetter = "h"
                        Case "C"
                            TranslateLetter = "i"
                        Case "B"
                            TranslateLetter = "j"
                        Case "A"
                            TranslateLetter = "k"
                        Case "H"
                            TranslateLetter = "l"
                        Case "G"
                            TranslateLetter = "m"
                        Case "F"
                            TranslateLetter = "n"
                        Case "E"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "L"
                            TranslateLetter = "p"
                        Case "K"
                            TranslateLetter = "q"
                        Case "J"
                            TranslateLetter = "r"
                        Case "I"
                            TranslateLetter = "s"
                        Case "P"
                            TranslateLetter = "t"
                        Case "O"
                            TranslateLetter = "u"
                        Case "N"
                            TranslateLetter = "v"
                        Case "M"
                            TranslateLetter = "w"
                        Case "D"
                            TranslateLetter = "x"
                        Case "C"
                            TranslateLetter = "y"
                        Case "B"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "L"
                            TranslateLetter = "0"
                        Case "K"
                            TranslateLetter = "1"
                        Case "J"
                            TranslateLetter = "2"
                        Case "I"
                            TranslateLetter = "3"
                        Case "P"
                            TranslateLetter = "4"
                        Case "O"
                            TranslateLetter = "5"
                        Case "N"
                            TranslateLetter = "6"
                        Case "M"
                            TranslateLetter = "7"
                        Case "D"
                            TranslateLetter = "8"
                        Case "C"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 8
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "K"
                            TranslateLetter = "a"
                        Case "J"
                            TranslateLetter = "b"
                        Case "I"
                            TranslateLetter = "c"
                        Case "P"
                            TranslateLetter = "d"
                        Case "O"
                            TranslateLetter = "e"
                        Case "N"
                            TranslateLetter = "f"
                        Case "M"
                            TranslateLetter = "g"
                        Case "D"
                            TranslateLetter = "h"
                        Case "C"
                            TranslateLetter = "i"
                        Case "B"
                            TranslateLetter = "j"
                        Case "A"
                            TranslateLetter = "k"
                        Case "H"
                            TranslateLetter = "l"
                        Case "G"
                            TranslateLetter = "m"
                        Case "F"
                            TranslateLetter = "n"
                        Case "E"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "L"
                            TranslateLetter = "p"
                        Case "K"
                            TranslateLetter = "q"
                        Case "J"
                            TranslateLetter = "r"
                        Case "I"
                            TranslateLetter = "s"
                        Case "P"
                            TranslateLetter = "t"
                        Case "O"
                            TranslateLetter = "u"
                        Case "N"
                            TranslateLetter = "v"
                        Case "M"
                            TranslateLetter = "w"
                        Case "D"
                            TranslateLetter = "x"
                        Case "C"
                            TranslateLetter = "y"
                        Case "B"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "L"
                            TranslateLetter = "0"
                        Case "K"
                            TranslateLetter = "1"
                        Case "J"
                            TranslateLetter = "2"
                        Case "I"
                            TranslateLetter = "3"
                        Case "P"
                            TranslateLetter = "4"
                        Case "O"
                            TranslateLetter = "5"
                        Case "N"
                            TranslateLetter = "6"
                        Case "M"
                            TranslateLetter = "7"
                        Case "D"
                            TranslateLetter = "8"
                        Case "C"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 9
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "G"
                            TranslateLetter = "a"
                        Case "F"
                            TranslateLetter = "b"
                        Case "E"
                            TranslateLetter = "c"
                        Case "D"
                            TranslateLetter = "d"
                        Case "C"
                            TranslateLetter = "e"
                        Case "B"
                            TranslateLetter = "f"
                        Case "A"
                            TranslateLetter = "g"
                        Case "P"
                            TranslateLetter = "h"
                        Case "O"
                            TranslateLetter = "i"
                        Case "N"
                            TranslateLetter = "j"
                        Case "M"
                            TranslateLetter = "k"
                        Case "L"
                            TranslateLetter = "l"
                        Case "K"
                            TranslateLetter = "m"
                        Case "J"
                            TranslateLetter = "n"
                        Case "I"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "H"
                            TranslateLetter = "p"
                        Case "G"
                            TranslateLetter = "q"
                        Case "F"
                            TranslateLetter = "r"
                        Case "E"
                            TranslateLetter = "s"
                        Case "D"
                            TranslateLetter = "t"
                        Case "C"
                            TranslateLetter = "u"
                        Case "B"
                            TranslateLetter = "v"
                        Case "A"
                            TranslateLetter = "w"
                        Case "P"
                            TranslateLetter = "x"
                        Case "O"
                            TranslateLetter = "y"
                        Case "N"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "H"
                            TranslateLetter = "0"
                        Case "G"
                            TranslateLetter = "1"
                        Case "F"
                            TranslateLetter = "2"
                        Case "E"
                            TranslateLetter = "3"
                        Case "D"
                            TranslateLetter = "4"
                        Case "C"
                            TranslateLetter = "5"
                        Case "B"
                            TranslateLetter = "6"
                        Case "A"
                            TranslateLetter = "7"
                        Case "P"
                            TranslateLetter = "8"
                        Case "O"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 10
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "P"
                            TranslateLetter = "a"
                        Case "M"
                            TranslateLetter = "b"
                        Case "N"
                            TranslateLetter = "c"
                        Case "K"
                            TranslateLetter = "d"
                        Case "L"
                            TranslateLetter = "e"
                        Case "I"
                            TranslateLetter = "f"
                        Case "J"
                            TranslateLetter = "g"
                        Case "G"
                            TranslateLetter = "h"
                        Case "H"
                            TranslateLetter = "i"
                        Case "E"
                            TranslateLetter = "j"
                        Case "F"
                            TranslateLetter = "k"
                        Case "C"
                            TranslateLetter = "l"
                        Case "D"
                            TranslateLetter = "m"
                        Case "A"
                            TranslateLetter = "n"
                        Case "B"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "O"
                            TranslateLetter = "p"
                        Case "P"
                            TranslateLetter = "q"
                        Case "M"
                            TranslateLetter = "r"
                        Case "N"
                            TranslateLetter = "s"
                        Case "K"
                            TranslateLetter = "t"
                        Case "L"
                            TranslateLetter = "u"
                        Case "I"
                            TranslateLetter = "v"
                        Case "J"
                            TranslateLetter = "w"
                        Case "G"
                            TranslateLetter = "x"
                        Case "H"
                            TranslateLetter = "y"
                        Case "E"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "O"
                            TranslateLetter = "0"
                        Case "P"
                            TranslateLetter = "1"
                        Case "M"
                            TranslateLetter = "2"
                        Case "N"
                            TranslateLetter = "3"
                        Case "K"
                            TranslateLetter = "4"
                        Case "L"
                            TranslateLetter = "5"
                        Case "I"
                            TranslateLetter = "6"
                        Case "J"
                            TranslateLetter = "7"
                        Case "G"
                            TranslateLetter = "8"
                        Case "H"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 11
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "M"
                            TranslateLetter = "a"
                        Case "P"
                            TranslateLetter = "b"
                        Case "O"
                            TranslateLetter = "c"
                        Case "J"
                            TranslateLetter = "d"
                        Case "I"
                            TranslateLetter = "e"
                        Case "L"
                            TranslateLetter = "f"
                        Case "K"
                            TranslateLetter = "g"
                        Case "F"
                            TranslateLetter = "h"
                        Case "E"
                            TranslateLetter = "i"
                        Case "H"
                            TranslateLetter = "j"
                        Case "G"
                            TranslateLetter = "k"
                        Case "B"
                            TranslateLetter = "l"
                        Case "A"
                            TranslateLetter = "m"
                        Case "D"
                            TranslateLetter = "n"
                        Case "C"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "N"
                            TranslateLetter = "p"
                        Case "M"
                            TranslateLetter = "q"
                        Case "P"
                            TranslateLetter = "r"
                        Case "O"
                            TranslateLetter = "s"
                        Case "J"
                            TranslateLetter = "t"
                        Case "I"
                            TranslateLetter = "u"
                        Case "L"
                            TranslateLetter = "v"
                        Case "K"
                            TranslateLetter = "w"
                        Case "F"
                            TranslateLetter = "x"
                        Case "E"
                            TranslateLetter = "y"
                        Case "H"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "N"
                            TranslateLetter = "0"
                        Case "M"
                            TranslateLetter = "1"
                        Case "P"
                            TranslateLetter = "2"
                        Case "O"
                            TranslateLetter = "3"
                        Case "J"
                            TranslateLetter = "4"
                        Case "I"
                            TranslateLetter = "5"
                        Case "L"
                            TranslateLetter = "6"
                        Case "K"
                            TranslateLetter = "7"
                        Case "F"
                            TranslateLetter = "8"
                        Case "E"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 12
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "L"
                            TranslateLetter = "a"
                        Case "I"
                            TranslateLetter = "b"
                        Case "J"
                            TranslateLetter = "c"
                        Case "O"
                            TranslateLetter = "d"
                        Case "P"
                            TranslateLetter = "e"
                        Case "M"
                            TranslateLetter = "f"
                        Case "N"
                            TranslateLetter = "g"
                        Case "C"
                            TranslateLetter = "h"
                        Case "D"
                            TranslateLetter = "i"
                        Case "A"
                            TranslateLetter = "j"
                        Case "B"
                            TranslateLetter = "k"
                        Case "G"
                            TranslateLetter = "l"
                        Case "H"
                            TranslateLetter = "m"
                        Case "E"
                            TranslateLetter = "n"
                        Case "F"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "K"
                            TranslateLetter = "p"
                        Case "L"
                            TranslateLetter = "q"
                        Case "I"
                            TranslateLetter = "r"
                        Case "J"
                            TranslateLetter = "s"
                        Case "O"
                            TranslateLetter = "t"
                        Case "P"
                            TranslateLetter = "u"
                        Case "M"
                            TranslateLetter = "v"
                        Case "N"
                            TranslateLetter = "w"
                        Case "C"
                            TranslateLetter = "x"
                        Case "D"
                            TranslateLetter = "y"
                        Case "A"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "K"
                            TranslateLetter = "0"
                        Case "L"
                            TranslateLetter = "1"
                        Case "I"
                            TranslateLetter = "2"
                        Case "J"
                            TranslateLetter = "3"
                        Case "O"
                            TranslateLetter = "4"
                        Case "P"
                            TranslateLetter = "5"
                        Case "M"
                            TranslateLetter = "6"
                        Case "N"
                            TranslateLetter = "7"
                        Case "C"
                            TranslateLetter = "8"
                        Case "D"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 13
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "F"
                            TranslateLetter = "a"
                        Case "G"
                            TranslateLetter = "b"
                        Case "H"
                            TranslateLetter = "c"
                        Case "A"
                            TranslateLetter = "d"
                        Case "B"
                            TranslateLetter = "e"
                        Case "C"
                            TranslateLetter = "f"
                        Case "D"
                            TranslateLetter = "g"
                        Case "M"
                            TranslateLetter = "h"
                        Case "N"
                            TranslateLetter = "i"
                        Case "O"
                            TranslateLetter = "j"
                        Case "P"
                            TranslateLetter = "k"
                        Case "I"
                            TranslateLetter = "l"
                        Case "J"
                            TranslateLetter = "m"
                        Case "K"
                            TranslateLetter = "n"
                        Case "L"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "E"
                            TranslateLetter = "p"
                        Case "F"
                            TranslateLetter = "q"
                        Case "G"
                            TranslateLetter = "r"
                        Case "H"
                            TranslateLetter = "s"
                        Case "A"
                            TranslateLetter = "t"
                        Case "B"
                            TranslateLetter = "u"
                        Case "C"
                            TranslateLetter = "v"
                        Case "D"
                            TranslateLetter = "w"
                        Case "M"
                            TranslateLetter = "x"
                        Case "N"
                            TranslateLetter = "y"
                        Case "O"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "E"
                            TranslateLetter = "0"
                        Case "F"
                            TranslateLetter = "1"
                        Case "G"
                            TranslateLetter = "2"
                        Case "H"
                            TranslateLetter = "3"
                        Case "A"
                            TranslateLetter = "4"
                        Case "B"
                            TranslateLetter = "5"
                        Case "C"
                            TranslateLetter = "6"
                        Case "D"
                            TranslateLetter = "7"
                        Case "M"
                            TranslateLetter = "8"
                        Case "N"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 14
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "J"
                            TranslateLetter = "a"
                        Case "K"
                            TranslateLetter = "b"
                        Case "L"
                            TranslateLetter = "c"
                        Case "M"
                            TranslateLetter = "d"
                        Case "N"
                            TranslateLetter = "e"
                        Case "O"
                            TranslateLetter = "f"
                        Case "P"
                            TranslateLetter = "g"
                        Case "A"
                            TranslateLetter = "h"
                        Case "B"
                            TranslateLetter = "i"
                        Case "C"
                            TranslateLetter = "j"
                        Case "D"
                            TranslateLetter = "k"
                        Case "E"
                            TranslateLetter = "l"
                        Case "F"
                            TranslateLetter = "m"
                        Case "G"
                            TranslateLetter = "n"
                        Case "H"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "I"
                            TranslateLetter = "p"
                        Case "J"
                            TranslateLetter = "q"
                        Case "K"
                            TranslateLetter = "r"
                        Case "L"
                            TranslateLetter = "s"
                        Case "M"
                            TranslateLetter = "t"
                        Case "N"
                            TranslateLetter = "u"
                        Case "O"
                            TranslateLetter = "v"
                        Case "P"
                            TranslateLetter = "w"
                        Case "A"
                            TranslateLetter = "x"
                        Case "B"
                            TranslateLetter = "y"
                        Case "C"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "I"
                            TranslateLetter = "0"
                        Case "J"
                            TranslateLetter = "1"
                        Case "K"
                            TranslateLetter = "2"
                        Case "L"
                            TranslateLetter = "3"
                        Case "M"
                            TranslateLetter = "4"
                        Case "N"
                            TranslateLetter = "5"
                        Case "O"
                            TranslateLetter = "6"
                        Case "P"
                            TranslateLetter = "7"
                        Case "A"
                            TranslateLetter = "8"
                        Case "B"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 15
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "B"
                            TranslateLetter = "a"
                        Case "C"
                            TranslateLetter = "b"
                        Case "D"
                            TranslateLetter = "c"
                        Case "E"
                            TranslateLetter = "d"
                        Case "F"
                            TranslateLetter = "e"
                        Case "G"
                            TranslateLetter = "f"
                        Case "H"
                            TranslateLetter = "g"
                        Case "I"
                            TranslateLetter = "h"
                        Case "J"
                            TranslateLetter = "i"
                        Case "K"
                            TranslateLetter = "j"
                        Case "L"
                            TranslateLetter = "k"
                        Case "M"
                            TranslateLetter = "l"
                        Case "N"
                            TranslateLetter = "m"
                        Case "O"
                            TranslateLetter = "n"
                        Case "P"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "A"
                            TranslateLetter = "p"
                        Case "B"
                            TranslateLetter = "q"
                        Case "C"
                            TranslateLetter = "r"
                        Case "D"
                            TranslateLetter = "s"
                        Case "E"
                            TranslateLetter = "t"
                        Case "F"
                            TranslateLetter = "u"
                        Case "G"
                            TranslateLetter = "v"
                        Case "H"
                            TranslateLetter = "w"
                        Case "I"
                            TranslateLetter = "x"
                        Case "J"
                            TranslateLetter = "y"
                        Case "K"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "A"
                            TranslateLetter = "0"
                        Case "B"
                            TranslateLetter = "1"
                        Case "C"
                            TranslateLetter = "2"
                        Case "D"
                            TranslateLetter = "3"
                        Case "E"
                            TranslateLetter = "4"
                        Case "F"
                            TranslateLetter = "5"
                        Case "G"
                            TranslateLetter = "6"
                        Case "H"
                            TranslateLetter = "7"
                        Case "I"
                            TranslateLetter = "8"
                        Case "J"
                            TranslateLetter = "9"
                    End Select
            End Select
        Case 16
            Select Case TheCase
                Case 1, 3
                    Select Case CurChar
                        Case "A"
                            TranslateLetter = "a"
                        Case "D"
                            TranslateLetter = "b"
                        Case "C"
                            TranslateLetter = "c"
                        Case "F"
                            TranslateLetter = "d"
                        Case "E"
                            TranslateLetter = "e"
                        Case "H"
                            TranslateLetter = "f"
                        Case "G"
                            TranslateLetter = "g"
                        Case "J"
                            TranslateLetter = "h"
                        Case "I"
                            TranslateLetter = "i"
                        Case "L"
                            TranslateLetter = "j"
                        Case "K"
                            TranslateLetter = "k"
                        Case "N"
                            TranslateLetter = "l"
                        Case "M"
                            TranslateLetter = "m"
                        Case "P"
                            TranslateLetter = "n"
                        Case "O"
                            TranslateLetter = "o"
                    End Select
                Case 2, 4
                    Select Case CurChar
                        Case "B"
                            TranslateLetter = "p"
                        Case "A"
                            TranslateLetter = "q"
                        Case "D"
                            TranslateLetter = "r"
                        Case "C"
                            TranslateLetter = "s"
                        Case "F"
                            TranslateLetter = "t"
                        Case "E"
                            TranslateLetter = "u"
                        Case "H"
                            TranslateLetter = "v"
                        Case "G"
                            TranslateLetter = "w"
                        Case "J"
                            TranslateLetter = "x"
                        Case "I"
                            TranslateLetter = "y"
                        Case "L"
                            TranslateLetter = "z"
                    End Select
                Case 5
                    Select Case CurChar
                        Case "B"
                            TranslateLetter = "0"
                        Case "A"
                            TranslateLetter = "1"
                        Case "D"
                            TranslateLetter = "2"
                        Case "C"
                            TranslateLetter = "3"
                        Case "F"
                            TranslateLetter = "4"
                        Case "E"
                            TranslateLetter = "5"
                        Case "H"
                            TranslateLetter = "6"
                        Case "G"
                            TranslateLetter = "7"
                        Case "J"
                            TranslateLetter = "8"
                        Case "I"
                            TranslateLetter = "9"
                    End Select
            End Select
    End Select
End Function

Public Function DecryptAIMPassword(ThePassword As String) As String
    Dim Temp As String, i As Integer, FirstPart As String
    Dim SecondPart As String, FirstOrSecond As Boolean
    Dim TempLetter As String, TempOutputLetter As String
    Dim TheCase As Integer
    'Only every other chracter in the encrypted
    'password actually will be translated into
    'another letter.  Every other character just
    'tells you a general idea of the chracter,
    'these being listed in the GetCase function.
    'We seperate them into two strings which will
    'be handled by seperate algoerithms.
    For i = 3 To Len(ThePassword)
        FirstOrSecond = Not FirstOrSecond
        If FirstOrSecond = True Then
            FirstPart = FirstPart & Mid(ThePassword, i, 1)
        Else
            SecondPart = SecondPart & Mid(ThePassword, i, 1)
        End If
    Next
    For i = 1 To Len(FirstPart)
        TempLetter = Mid(FirstPart, i, 1)
        'We get which of the 5 kinds of letters
        'the current chracter is.
        TheCase = GetCase(i, TempLetter)
        TempLetter = Mid(SecondPart, i, 1)
        'We then use the information we learned
        'in the first algorithm to arrive at a
        'final letter.
        TempLetter = TranslateLetter(i, TempLetter, TheCase)
        'if the GetCase algorithm returns 1 or 2
        'then that means the letter should be
        'capital.  This is not neccessary since
        'passwords are not case sensative.
        If TheCase = 1 Or TheCase = 2 Then TempLetter = UCase(TempLetter)
        'Add it to the full password and move on.
        Temp = Temp & TempLetter
    Next
    DecryptAIMPassword = Temp
End Function

Public Sub GetAIMs(Screenname As ListBox, Password As ListBox)
    Dim SubDirs As Variant, Reg As New CReadWriteEasyReg, i As Integer
    Dim TempPassword As String, TempScreenname As String, SubDirs2 As Variant
    Dim ScreennameCol As New Collection, PasswordCol As New Collection, j As Integer
    Dim AlreadyFound As Boolean
    'This part is where I call on the class
    'module I found at PlanetSourceCode.  It
    'is to extract screennames and encrypted
    'passwords from the registry, then decrypt
    'the passwords and add to listboxes.
    If Not Reg.OpenRegistry(HKEY_USERS, ".DEFAULT\Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users") Then Exit Sub
    SubDirs = Reg.GetAllSubDirectories
    'We get first set of screennames and passwords here
    For i = LBound(SubDirs) To UBound(SubDirs)
        If Not Reg.OpenRegistry(HKEY_USERS, ".DEFAULT\Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users") Then Exit Sub
        TempScreenname = Reg.GetValue(SubDirs(i))
        ScreennameCol.Add TempScreenname
        If Reg.OpenRegistry(HKEY_USERS, ".DEFAULT\Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users\" & SubDirs(i) & "\Login") Then
            TempPassword = Reg.GetValue("Password")
            If TempPassword = "" Then
                PasswordCol.Add "<No Password>"
            Else
                PasswordCol.Add DecryptAIMPassword(TempPassword)
            End If
        Else
            PasswordCol.Add "<No Password>"
        End If
    Next
    'Now we check HKEY_CURRENT_USER
    If Not Reg.OpenRegistry(HKEY_CURRENT_USER, "Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users") Then Exit Sub
    SubDirs2 = Reg.GetAllSubDirectories
    'We get second set of screennames and passwords here
    For i = LBound(SubDirs2) To UBound(SubDirs2)
        If Not Reg.OpenRegistry(HKEY_CURRENT_USER, "Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users") Then Exit Sub
        TempScreenname = Reg.GetValue(SubDirs2(i))
        If Reg.OpenRegistry(HKEY_CURRENT_USER, "Software\America Online\AOL Instant Messenger (TM)\CurrentVersion\Users\" & SubDirs2(i) & "\Login") Then
            TempPassword = Reg.GetValue("Password")
            If TempPassword = "" Then
                TempPassword = "<No Password>"
            Else
                TempPassword = DecryptAIMPassword(TempPassword)
            End If
        Else
            TempPassword = "<No Password>"
        End If
        AlreadyFound = False
        For j = 1 To ScreennameCol.Count
            If TempScreenname = ScreennameCol.Item(j) Then
                If Not TempPassword = PasswordCol.Item(j) Then
                    ScreennameCol.Add TempScreenname
                    PasswordCol.Add TempPassword
                End If
                AlreadyFound = True
                Exit For
            End If
        Next
        If AlreadyFound = False Then
            ScreennameCol.Add TempScreenname
            PasswordCol.Add TempPassword
        End If
    Next
    Reg.CloseRegistry
    For i = 1 To ScreennameCol.Count
        Screenname.AddItem ScreennameCol.Item(i)
        Password.AddItem PasswordCol.Item(i)
    Next
End Sub
