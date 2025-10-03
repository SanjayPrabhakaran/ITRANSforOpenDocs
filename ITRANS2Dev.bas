REM  *****  BASIC  *****
REM  ITRANS to Devanagari Converter for LibreOffice Writer
REM  Select text in ITRANS format and run this macro to convert it
	
Sub ConvertITRANSToSanskrit()
    Dim oDoc As Object
    Dim oSelection As Object
    Dim sText As String
    Dim sConverted As String
    
    ' Get the current document and selection
    oDoc = ThisComponent
    oSelection = oDoc.getCurrentSelection()
    
    ' Check if text is selected
    If oSelection.getCount() = 0 Then
        MsgBox "Please select text to convert", 48, "No Selection"
        Exit Sub
    End If
    
    ' Get selected text
    sText = oSelection.getByIndex(0).getString()
    
    If Len(sText) = 0 Then
        MsgBox "No text selected", 48, "Empty Selection"
        Exit Sub
    End If
    
    ' Convert ITRANS to Devanagari
    sConverted = ITRANSToDev(sText)
    
    ' Copy to clipboard using dispatcher
    CopyToClipboard(sConverted)
    
    MsgBox "Converted text copied to clipboard!" & Chr(10) & Chr(10) & _
           "Original: " & sText & Chr(10) & Chr(10) & _
           "Converted: " & sConverted, 64, "Conversion Complete"
End Sub

Sub CopyToClipboard(sText As String)
    ' Use dispatcher method to copy text to clipboard
    Dim oFrame As Object
    Dim oDispatcher As Object
    Dim oController As Object
    Dim oModel As Object
    Dim oText As Object
    Dim oCursor As Object
    
    ' Create a temporary hidden document
    oModel = StarDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, Array())
    oController = oModel.getCurrentController()
    oFrame = oController.getFrame()
    oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    
    ' Get text object and insert our converted text
    oText = oModel.getText()
    oCursor = oText.createTextCursor()
    oText.insertString(oCursor, sText, False)
    
    ' Select all text
    oCursor.gotoStart(False)
    oCursor.gotoEnd(True)
    oController.select(oCursor)
    
    ' Copy to clipboard
    oDispatcher.executeDispatch(oFrame, ".uno:Copy", "", 0, Array())
    
    ' Close temporary document without saving
    oModel.close(True)
End Sub

Function ITRANSToDev(sInput As String) As String
    Dim sResult As String
    Dim i As Integer
    Dim iLen As Integer
    Dim s2 As String, s3 As String
    Dim bAfterConsonant As Boolean
    Dim bMatched As Boolean
    Dim bIsConsonant As Boolean
    Dim sConsonant As String
    Dim iConsumed As Integer
    Dim sVowel As String
    Dim iVowelLen As Integer
    Dim sChar As String
    
    sResult = ""
    i = 1
    iLen = Len(sInput)
    bAfterConsonant = False
    
    While i <= iLen
        bMatched = False
        bIsConsonant = False
        sConsonant = ""
        iConsumed = 0
        
        ' Check 3-character sequences
        If i <= iLen - 2 Then
            s3 = Mid(sInput, i, 3)
            Select Case s3
                ' Independent vowels (only if not after consonant)
                Case "RRi": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ऋ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ृ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "RRI": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ॠ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ॄ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "LLi": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ऌ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ॢ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "LLI": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ॡ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ॣ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "R^i": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ऋ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ृ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "R^I": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ॠ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ॄ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "L^i": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ऌ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ॢ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "L^I": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ॡ": i = i + 3: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ॣ": i = i + 3: bMatched = True: bAfterConsonant = False
                    End If
                Case "AUM": sResult = sResult & "ॐ": i = i + 3: bMatched = True: bAfterConsonant = False
                Case "kSh": sConsonant = "क्ष्": iConsumed = 3: bIsConsonant = True
                Case "j~n": sConsonant = "ज्ञ्": iConsumed = 3: bIsConsonant = True
                Case "dny": sConsonant = "ज्ञ्": iConsumed = 3: bIsConsonant = True
                Case "shr": sConsonant = "श्र्": iConsumed = 3: bIsConsonant = True
                Case "shh": sConsonant = "ष्": iConsumed = 3: bIsConsonant = True
                Case ".Dh": sConsonant = "ढ़्": iConsumed = 3: bIsConsonant = True
            End Select
        End If
        
        If bMatched Then GoTo NextIteration
        
        ' Check 2-character sequences
        If Not bIsConsonant And i <= iLen - 1 Then
            s2 = Mid(sInput, i, 2)
            Select Case s2
                ' Vowels - check if after consonant for matra
                Case "aa", "AA": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "आ": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ा": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                Case "ii", "II": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ई": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ी": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                Case "uu", "UU": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ऊ": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ू": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                Case "ai": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "ऐ": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ै": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                Case "au": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "औ": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ौ": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                Case "aM": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "अं": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ं": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                Case "aH": 
                    If Not bAfterConsonant Then
                        sResult = sResult & "अः": i = i + 2: bMatched = True: bAfterConsonant = False
                    Else
                        sResult = sResult & "ः": i = i + 2: bMatched = True: bAfterConsonant = False
                    End If
                
                ' Consonants (will be rendered with halanta by default)
                Case "kh": sConsonant = "ख्": iConsumed = 2: bIsConsonant = True
                Case "gh": sConsonant = "घ्": iConsumed = 2: bIsConsonant = True
                Case "~N": sConsonant = "ङ्": iConsumed = 2: bIsConsonant = True
                Case "ch": sConsonant = "च्": iConsumed = 2: bIsConsonant = True
                Case "Ch": sConsonant = "छ्": iConsumed = 2: bIsConsonant = True
                Case "jh": sConsonant = "झ्": iConsumed = 2: bIsConsonant = True
                Case "~n": sConsonant = "ञ्": iConsumed = 2: bIsConsonant = True
                Case "Th": sConsonant = "ठ्": iConsumed = 2: bIsConsonant = True
                Case "Dh": sConsonant = "ढ्": iConsumed = 2: bIsConsonant = True
                Case "th": sConsonant = "थ्": iConsumed = 2: bIsConsonant = True
                Case "dh": sConsonant = "ध्": iConsumed = 2: bIsConsonant = True
                Case "ph": sConsonant = "फ्": iConsumed = 2: bIsConsonant = True
                Case "bh": sConsonant = "भ्": iConsumed = 2: bIsConsonant = True
                Case "sh": sConsonant = "श्": iConsumed = 2: bIsConsonant = True
                Case "Sh": sConsonant = "ष्": iConsumed = 2: bIsConsonant = True
                Case "GY": sConsonant = "ज्ञ्": iConsumed = 2: bIsConsonant = True
                Case "ld": sConsonant = "ळ्": iConsumed = 2: bIsConsonant = True
                Case "OM": sResult = sResult & "ॐ": i = i + 2: bMatched = True: bAfterConsonant = False
                
                ' Nukta consonants
                Case ".D": sConsonant = "ड़्": iConsumed = 2: bIsConsonant = True
                
                ' Specials
                Case ".n": sResult = sResult & "ं": i = i + 2: bMatched = True: bAfterConsonant = False
                Case ".m": sResult = sResult & "ं": i = i + 2: bMatched = True: bAfterConsonant = False
                Case ".a": sResult = sResult & "ऽ": i = i + 2: bMatched = True: bAfterConsonant = False
                Case ".c": sResult = sResult & "ँ": i = i + 2: bMatched = True: bAfterConsonant = False
                Case ".N": sResult = sResult & "ँ": i = i + 2: bMatched = True: bAfterConsonant = False
                Case ".h": sResult = sResult & "्": i = i + 2: bMatched = True: bAfterConsonant = False
            End Select
        End If
        
        If bMatched Then GoTo NextIteration
        
        ' Process consonant with following vowel
        If bIsConsonant Then
            sVowel = GetVowelMatra(sInput, i + iConsumed, iVowelLen)
            
            If sVowel = "" Then
                ' No vowel following, keep halanta as-is
                sResult = sResult & sConsonant
                bAfterConsonant = True
            ElseIf sVowel = "a" Then
                ' Inherent 'a', remove halanta and skip the 'a'
                sResult = sResult & Left(sConsonant, Len(sConsonant) - 1)
                i = i + iVowelLen  ' Skip the 'a' character
                bAfterConsonant = False
            Else
                ' Add consonant without halanta and with vowel matra
                sResult = sResult & Left(sConsonant, Len(sConsonant) - 1) & sVowel
                i = i + iVowelLen  ' Skip the vowel characters
                bAfterConsonant = False
            End If
            i = i + iConsumed
            GoTo NextIteration
        End If
        
        ' Single character mappings
        sChar = Mid(sInput, i, 1)
        
        Select Case sChar
            ' Vowels - check if after consonant
            Case "a": 
                If Not bAfterConsonant Then
                    sResult = sResult & "अ": bAfterConsonant = False
                Else
                    ' Inherent a after consonant - already handled
                    bAfterConsonant = False
                End If
            Case "A": 
                If Not bAfterConsonant Then
                    sResult = sResult & "आ": bAfterConsonant = False
                Else
                    sResult = sResult & "ा": bAfterConsonant = False
                End If
            Case "i": 
                If Not bAfterConsonant Then
                    sResult = sResult & "इ": bAfterConsonant = False
                Else
                    sResult = sResult & "ि": bAfterConsonant = False
                End If
            Case "I": 
                If Not bAfterConsonant Then
                    sResult = sResult & "ई": bAfterConsonant = False
                Else
                    sResult = sResult & "ी": bAfterConsonant = False
                End If
            Case "u": 
                If Not bAfterConsonant Then
                    sResult = sResult & "उ": bAfterConsonant = False
                Else
                    sResult = sResult & "ु": bAfterConsonant = False
                End If
            Case "U": 
                If Not bAfterConsonant Then
                    sResult = sResult & "ऊ": bAfterConsonant = False
                Else
                    sResult = sResult & "ू": bAfterConsonant = False
                End If
            Case "e": 
                If Not bAfterConsonant Then
                    sResult = sResult & "ए": bAfterConsonant = False
                Else
                    sResult = sResult & "े": bAfterConsonant = False
                End If
            Case "o": 
                If Not bAfterConsonant Then
                    sResult = sResult & "ओ": bAfterConsonant = False
                Else
                    sResult = sResult & "ो": bAfterConsonant = False
                End If
            
            ' Single consonants
            Case "k", "g", "j", "T", "D", "N", "t", "d", "n", "p", "b", "m", "y", "r", "l", "v", "w", "s", "h", "L", "q", "K", "G", "z", "J", "f"
                sConsonant = GetSingleConsonant(sChar)
                sVowel = GetVowelMatra(sInput, i + 1, iVowelLen)
                If sVowel = "" Then
                    sResult = sResult & sConsonant
                    bAfterConsonant = True
                ElseIf sVowel = "a" Then
                    sResult = sResult & Left(sConsonant, Len(sConsonant) - 1)
                    i = i + iVowelLen
                    bAfterConsonant = False
                Else
                    sResult = sResult & Left(sConsonant, Len(sConsonant) - 1) & sVowel
                    i = i + iVowelLen
                    bAfterConsonant = False
                End If
            
            ' Specials
            Case "M": sResult = sResult & "ं": bAfterConsonant = False
            Case "H": sResult = sResult & "ः": bAfterConsonant = False
            Case "x": 
                sConsonant = "क्ष्": sVowel = GetVowelMatra(sInput, i + 1, iVowelLen)
                If sVowel = "" Then
                    sResult = sResult & sConsonant
                    bAfterConsonant = True
                ElseIf sVowel = "a" Then
                    sResult = sResult & Left(sConsonant, Len(sConsonant) - 1)
                    i = i + iVowelLen
                    bAfterConsonant = False
                Else
                    sResult = sResult & Left(sConsonant, Len(sConsonant) - 1) & sVowel
                    i = i + iVowelLen
                    bAfterConsonant = False
                End If
            Case "R": sResult = sResult & "र्": bAfterConsonant = False
            Case "Y": sResult = sResult & "য": bAfterConsonant = False
            
            ' Keep other characters as-is (spaces, punctuation, etc.)
            Case Else: sResult = sResult & sChar: bAfterConsonant = False
        End Select
        
        i = i + 1
        
        NextIteration:
    Wend
    
    ITRANSToDev = sResult
End Function

Function GetSingleConsonant(sChar As String) As String
    Select Case sChar
        Case "k": GetSingleConsonant = "क्"
        Case "g": GetSingleConsonant = "ग्"
        Case "j": GetSingleConsonant = "ज्"
        Case "T": GetSingleConsonant = "ट्"
        Case "D": GetSingleConsonant = "ड्"
        Case "N": GetSingleConsonant = "ण्"
        Case "t": GetSingleConsonant = "त्"
        Case "d": GetSingleConsonant = "द्"
        Case "n": GetSingleConsonant = "न्"
        Case "p": GetSingleConsonant = "प्"
        Case "b": GetSingleConsonant = "ब्"
        Case "m": GetSingleConsonant = "म्"
        Case "y": GetSingleConsonant = "य्"
        Case "r": GetSingleConsonant = "र्"
        Case "l": GetSingleConsonant = "ल्"
        Case "v": GetSingleConsonant = "व्"
        Case "w": GetSingleConsonant = "व्"
        Case "s": GetSingleConsonant = "स्"
        Case "h": GetSingleConsonant = "ह्"
        Case "L": GetSingleConsonant = "ळ्"
        Case "q": GetSingleConsonant = "क़्"
        Case "K": GetSingleConsonant = "ख़्"
        Case "G": GetSingleConsonant = "ग़्"
        Case "z": GetSingleConsonant = "ज़्"
        Case "J": GetSingleConsonant = "ज़्"
        Case "f": GetSingleConsonant = "फ़्"
        Case Else: GetSingleConsonant = ""
    End Select
End Function

Function GetVowelMatra(sInput As String, iPos As Integer, ByRef iVowelLen As Integer) As String
    ' Check if there's a vowel following the consonant and return the matra
    ' Returns "a" for inherent 'a', empty string if consonant/space follows
    ' iVowelLen returns the number of characters consumed
    
    iVowelLen = 0
    
    If iPos > Len(sInput) Then
        GetVowelMatra = ""
        Exit Function
    End If
    
    Dim sNext As String, sNext2 As String, sNext3 As String
    sNext = Mid(sInput, iPos, 1)
    
    If iPos <= Len(sInput) - 1 Then
        sNext2 = Mid(sInput, iPos, 2)
    Else
        sNext2 = ""
    End If
    
    If iPos <= Len(sInput) - 2 Then
        sNext3 = Mid(sInput, iPos, 3)
    Else
        sNext3 = ""
    End If
    
    ' Check 3-char vowels first
    Select Case sNext3
        Case "RRi": GetVowelMatra = "ृ": iVowelLen = 3: Exit Function
        Case "RRI": GetVowelMatra = "ॄ": iVowelLen = 3: Exit Function
        Case "LLi": GetVowelMatra = "ॢ": iVowelLen = 3: Exit Function
        Case "LLI": GetVowelMatra = "ॣ": iVowelLen = 3: Exit Function
        Case "R^i": GetVowelMatra = "ृ": iVowelLen = 3: Exit Function
        Case "R^I": GetVowelMatra = "ॄ": iVowelLen = 3: Exit Function
        Case "L^i": GetVowelMatra = "ॢ": iVowelLen = 3: Exit Function
        Case "L^I": GetVowelMatra = "ॣ": iVowelLen = 3: Exit Function
    End Select
    
    ' Check 2-char vowels
    Select Case sNext2
        Case "aa", "AA": GetVowelMatra = "ा": iVowelLen = 2
        Case "ii", "II": GetVowelMatra = "ी": iVowelLen = 2
        Case "uu", "UU": GetVowelMatra = "ू": iVowelLen = 2
        Case "ai": GetVowelMatra = "ै": iVowelLen = 2
        Case "au": GetVowelMatra = "ौ": iVowelLen = 2
        Case "aM": GetVowelMatra = "ं": iVowelLen = 2
        Case "aH": GetVowelMatra = "ः": iVowelLen = 2
        Case Else
            ' Check single char vowels
            Select Case sNext
                Case "a": GetVowelMatra = "a": iVowelLen = 1 ' Inherent 'a', return marker
                Case "A": GetVowelMatra = "ा": iVowelLen = 1
                Case "i": GetVowelMatra = "ि": iVowelLen = 1
                Case "I": GetVowelMatra = "ी": iVowelLen = 1
                Case "u": GetVowelMatra = "ु": iVowelLen = 1
                Case "U": GetVowelMatra = "ू": iVowelLen = 1
                Case "e": GetVowelMatra = "े": iVowelLen = 1
                Case "o": GetVowelMatra = "ो": iVowelLen = 1
                Case "M": GetVowelMatra = "ं": iVowelLen = 1
                Case "H": GetVowelMatra = "ः": iVowelLen = 1
                Case Else: GetVowelMatra = "": iVowelLen = 0 ' No vowel, keep halanta
            End Select
    End Select
End Function
