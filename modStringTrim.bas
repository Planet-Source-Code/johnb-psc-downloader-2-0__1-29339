Attribute VB_Name = "modStringTrim"
Option Explicit

'//////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'// Function ExtTrim                                                                        \\
'// Trims a string - removing/replacing signs and spaces - Used to trim a filename...       \\
'//                                                                                         \\
'// s                           String      The string to trim...                           \\
'// NoSigns                     True/False  Sets wether to allow certain signs              \\
'// AllowSpaces                 True/False  Set wheter to delete spaces or not              \\
'// SpaceReplacementCharacter   Character   Sets a character to replace space. Ex.: "_"     \\
'// ReplacementCharacter        Character   Sets a character to replace signs. Ex.: "_"     \\
'//                                                                                         \\
'//////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


Public Function ExtTrim(ByVal s As String, Optional NoSigns As Boolean = True, Optional AllowSpaces As Boolean = True, Optional SpaceReplacementCharacter As String = "", Optional ReplacementCharacter As String = "") As String
    Dim rS As String    'ResultString
    Dim sLen As Integer 'Length of string
    Dim i As Integer    'counter
    Dim Rpl As String
    Dim SpcRpl As String
        
    If Len(SpaceReplacementCharacter) = 0 Then
        SpcRpl = ""
    Else
        SpcRpl = SpaceReplacementCharacter
    End If
    If Len(ReplacementCharacter) = 0 Then
        Rpl = ""
    Else
        Rpl = ReplacementCharacter
    End If
    sLen = Len(s)
    
    For i = 1 To sLen   'Global counter
        rS = rS & SortChr(Mid(s, i, 1), NoSigns, AllowSpaces, SpcRpl, Rpl)
    Next i
        ExtTrim = rS
End Function

Private Function SortChr(ByVal c As String, NoSigns As Boolean, allowSpace As Boolean, Optional SpcRepl As String, Optional Repl As String) As String
    On Error Resume Next
    Dim cVal As Integer
    cVal = Asc(c)
    Select Case cVal
    Case 0 To 31
        'garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case 32
        'space
        If allowSpace Then
            If Len(SpcRepl) = 0 Then
                SortChr = c
            Else
                SortChr = SpcRepl
            End If
        Else
            SortChr = ""
        End If
    Case 33 To 39
        'Garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case 40 To 41
        'signs
        If NoSigns Then
            SortChr = ""
        Else
            SortChr = c
        End If
    Case 42
        'Garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case 43 To 46
        'Signs
        If NoSigns Then
            SortChr = ""
        Else
            SortChr = c
        End If
    Case 47
        'Garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case 48 To 57
        'numbers
        SortChr = c
    Case 58 To 63
        'Garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case 64
        'Signs
        If NoSigns Then
            SortChr = ""
        Else
            SortChr = c
        End If
    Case 65 To 90
        'Characters
        SortChr = c
    Case 91
        'Signs
        If NoSigns Then
            SortChr = ""
        Else
            SortChr = c
        End If
    Case 92
        'Garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case 93 To 96
        'Signs
        If NoSigns Then
            SortChr = ""
        Else
            SortChr = c
        End If
    Case 97 To 122
        'Characters
        SortChr = c
    Case 123 To 126
        'Signs
        If NoSigns Then
            SortChr = ""
        Else
            SortChr = c
        End If
    Case cVal > 127
        'Garbage
        If NoSigns Then
            SortChr = ""
        Else
            If Len(Repl) = 0 Then
                SortChr = ""
            Else
                SortChr = Repl
            End If
        End If
    Case Else
        SortChr = ""
    End Select
End Function

'Ascii evaluation
'                       In files
' values    Character   Yes No  Replacement?
' ------------------------------------------
' 0-31      Garbage         X
' 32        Space       X   X _
' 33        !               X
' 34        "               X
' 35        #               X
' 36        $               X
' 37        %               X
' 38        &               X
' 39        '               X
' 40        (           X
' 41        )           X
' 42        *               X
' 43        +           X
' 44        ,           X
' 45        -           X
' 46        .           X
' 47        /               X
' 48-57     Numbers     X
' 58        :               X
' 59        ;               X
' 60        <               X
' 61        =               X
' 62        >               X
' 63        ?               X
' 64        @           X
' 65-90     Characters  X
' 91        [           X
' 92        \               X
' 93        ]           X
' 94        ^           X
' 95        _           X
' 96        `           X   X
' 97-122    Characters  X
' 123       {           X   X
' 124       |           X
' 125       }           X   X
' 126       ~           X
' 127-      Garbage         X
'
'Rules...
' Not allowed:  0-31, 33-39, 42, 47, 58-63, 92, >127
' Alowed:       32, 40, 41, 43-46, 48-57 num, 64, 65-90 chr, 91, 93-96, 97-122 chr, 123-126


