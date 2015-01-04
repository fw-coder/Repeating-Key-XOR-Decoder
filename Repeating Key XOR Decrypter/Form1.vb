﻿Public Class Form1

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ''TextBox5.Text = HammingDistance(TextBox3.Text, TextBox4.Text).ToString ' write hamming distance to result box
    End Sub

    Private Function HammingDistance(inTxt1 As String, inTxt2 As String) As Integer
        Dim txt1 As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(inTxt1) ' Convert inputted strings into hex byte arrays
        Dim txt2 As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(inTxt2)
        Dim btxt1 As String() = New String(txt1.Length - 1) {}  ' Set up empty arrays to hold binary versions of hex byte arrays
        Dim btxt2 As String() = New String(txt2.Length - 1) {}
        Dim txtPosition1 As Long = 0  ' Set up index values to use for conversion from hex to binary
        Dim txtPosition2 As Long = 0

        ' Converting byte arrays to binary arrays
        For txtPosition1 = 0 To txt1.Length - 1
            btxt1(txtPosition1) = Convert.ToString(txt1(txtPosition1), 2) 'conv value to binary
            While btxt1(txtPosition1).Length < 8
                btxt1(txtPosition1) = "0" & btxt1(txtPosition1)  'make sure bin is a 8-bit binary string
            End While
        Next
        For txtPosition2 = 0 To txt2.Length - 1
            btxt2(txtPosition2) = Convert.ToString(txt2(txtPosition2), 2) 'conv value to binary
            While btxt2(txtPosition2).Length < 8
                btxt2(txtPosition2) = "0" & btxt2(txtPosition2)  'make sure bin is a 8-bit binary string
            End While
        Next

        ' Calculating hamming distance between values in input boxes by
        ' performing XOR bit by bit and summing the 1's (different bits) in HammingDistance
        HammingDistance = 0 ' Initializing HammingDistance to zero
        For i = 0 To btxt1.Length - 1                                   ' loop through each byte in the array,
            For j As Integer = 0 To 7                                   ' and loop through each bit in the byte
                Dim currentBit1 As String = btxt1(i).Substring(j, 1)
                Dim currentBit2 As String = btxt2(i).Substring(j, 1)
                If Not currentBit1 = currentBit2 Then HammingDistance = HammingDistance + 1
            Next
        Next

    End Function

    Private Sub OpenFileDialog1_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim strm As System.IO.Stream

        strm = OpenFileDialog1.OpenFile()

        TextBox1.Text = OpenFileDialog1.FileName.ToString()

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If

        ' Import strings from text file
        Dim R As New IO.StreamReader(OpenFileDialog1.FileName)
        Dim str As String = R.ReadToEnd() ' Delimiter is vbLF (LineFeed)
        cipherInput.fromFile = System.Text.ASCIIEncoding.ASCII.GetBytes(str)

        RichTextBox1.Text = str ' Put the strings from the imported file into the list box
        R.Close()

    End Sub
    Public Class cipherInput

        Public Shared fromFile() As Byte

    End Class

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = "C:temp"
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Probably would be better with a multi-dimensional array, but oh well \o/
        Dim bestLengthGuess1 As Integer = 0
        Dim bestLengthGuess2 As Integer = 0
        Dim bestLengthGuess3 As Integer = 0
        Dim shortestNormalizedDistance1 As Integer = 9999
        Dim shortestNormalizedDistance2 As Integer = 9999
        Dim shortestNormalizedDistance3 As Integer = 9999
        Dim cipherString As String = System.Text.ASCIIEncoding.ASCII.GetString(cipherInput.fromFile)

        ' Find the three keysizes with the shortest hamming distances and write them to textBox2
        For lengthGuess As Integer = 2 To 8 'should be to 40
            Dim NormalizedHammingDistance As Integer = HammingDistance(cipherString.Substring(0, lengthGuess), _
                                                                       cipherString.Substring(lengthGuess, lengthGuess)) / lengthGuess
            If NormalizedHammingDistance < shortestNormalizedDistance1 Then
                bestLengthGuess1 = lengthGuess
                shortestNormalizedDistance1 = NormalizedHammingDistance
            ElseIf NormalizedHammingDistance < shortestNormalizedDistance2 Then
                bestLengthGuess2 = lengthGuess
                shortestNormalizedDistance2 = NormalizedHammingDistance
            ElseIf NormalizedHammingDistance < shortestNormalizedDistance3 Then
                bestLengthGuess3 = lengthGuess
                shortestNormalizedDistance3 = NormalizedHammingDistance
            End If
        Next
        TextBox2.Text = _
            "L1=" & bestLengthGuess1 & " (" & shortestNormalizedDistance1 & ")" & Environment.NewLine & _
            "L2=" & bestLengthGuess2 & " (" & shortestNormalizedDistance2 & ")" & Environment.NewLine & _
            "L3=" & bestLengthGuess3 & " (" & shortestNormalizedDistance3 & ")" & Environment.NewLine

    End Sub

   
    
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        ' Grab text to decrypt - use cipherInput.fromFile()
        ' Dim cipher As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(RichTextBox1.Text)

        ' Grab keysize
        Dim keysize As Integer = TextBox7.Text

        ' 2-D array tutorial: http://www.homeandlearn.co.uk/NET/nets6p5.html
        ' 2D array defines a grid of data of size (#rows, #columns)
        'a0,0 = block 0, char 0 byte 1
        'a0,1 = block 0, char 1 byte 2      for string with 15 bytes, keysize of 5
        'a0,2 = block 0, char 2 byte 3      array = a(#bytes/keysize = 15/5 = 3, keysize = 5)
        'a0,3 = block 0, char 3 byte 4
        'a0,4 = block 0, char 4 byte 5 (note that 4, not 5, = index value to set up 2nd dimension of array)
        'a1,0 = block 1, char 0 byte 6
        'a1,1 = block 1, char 1 byte 7
        'a1,2 = block 1, char 2 byte 8
        'a1,3 = block 1, char 3 byte 9
        'a1,4 = block 1, char 4 byte 10
        'a2,0 = block 2, char 0 byte 11
        'a2,1 = block 2, char 1 byte 12
        'a2,2 = block 2, char 2 byte 13
        'a2,3 = block 2, char 3 byte 14
        'a2,4 = block 2, char 4 byte 15
        Dim arrayRow As Integer = (cipherInput.fromFile.Length / keysize) - 1 ' 0 to several hundred; each block of cipher on separate row
        Dim arrayCol As Integer = keysize - 1                   ' 0 to keysize (minus 1); each column of bytes in block of cipher

        Dim cipherBlocks(,) As Byte = New Byte(arrayRow, arrayCol) {}
        Dim counter As Integer = 0
        ' keep a counter going while creating 2d array & if counter >= #items in cipher, then stop or add empty values
        ' Break text into blocks of size = keysize
        For r As Integer = 0 To arrayRow
            For c As Integer = 0 To arrayCol
                cipherBlocks(r, c) = cipherInput.fromFile(counter)
                counter = counter + 1
                If counter >= cipherInput.fromFile.Length Then Exit For
            Next
        Next

        ' Transpose those blocks: block1 = 1st byte of each block, block2 = 2nd byte, etc.
        Dim cipherBlocksTrans(,) As Byte = New Byte(arrayCol, arrayRow) {}
        Dim counter2 As Integer = 0
        For r2 As Integer = 0 To arrayCol
            For c2 As Integer = 0 To arrayRow
                cipherBlocksTrans(r2, c2) = cipherBlocks(c2, r2)
                If counter2 >= cipherInput.fromFile.Length - 1 Then Exit For
            Next
        Next


        ' Solve each transposed block using single-byte XOR using bytes 00-FF and find byte for each block
        '     that gives best result

        Dim result As String = singleByteXor(cipherBlocksTrans) ' not returning anything useful yet

        TextBox8.Text = TextBox8.Text & "Done."


        ' Put single byte solutions together to make the key
        ' Use the key to decrypt original text using repeating-text XOR 

        Dim dummy As Integer = 0 ' just a place to put a break



        ' Getting repeating key and text to encrypt
        'Dim key As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(TextBox1.Text)
        'Dim txt As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(RichTextBox1.Text)
        'Dim counter As Integer = 0

        'For i As Integer = 0 To txt.Length - 1 Step 3
        '    For j As Integer = 0 To key.Length - 1
        '        Dim out As Byte = txt(counter) Xor key(j)
        '        If counter = 0 And out < &H10 Then
        '            TextBox3.Text = TextBox3.Text & "0" & Convert.ToString(out, 16) 'add leading 0 to first byte
        '        Else
        '            TextBox3.Text = TextBox3.Text & Convert.ToString(out, 16)
        '        End If
        '        counter = counter + 1
        '        If counter = txt.Length Then Exit For
        '    Next
        'Next
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        MessageBox.Show("Instructions:" & Environment.NewLine & "1. Select file containing ascii to decode." & _
                        Environment.NewLine & "2. Click button to find likely keysizes" & _
                        Environment.NewLine & "3. Type a keysize in the keysize box to use for decyrption" & _
                        Environment.NewLine & "4. Click the decrypt button" & _
                        Environment.NewLine & "5. If decrypted text looks wrong, pick another keysize and try again")

    End Sub

    Private Function singleByteXor(cipher(,) As Byte) As String
        
        For row As Integer = 0 To cipher.GetLength(0) - 1
            Dim topScore1 As Single = 999999
            Dim topScore2 As Single = 999999
            Dim topScore3 As Single = 999999
            Dim topScoreResultString1 As String = String.Empty
            Dim topScoreResultString2 As String = String.Empty
            Dim topScoreResultString3 As String = String.Empty
            Dim topScoreResultKey1 As Integer = 0
            Dim topScoreResultKey2 As Integer = 0
            Dim topScoreResultKey3 As Integer = 0

            Dim topScore4 As Single = 999999
            Dim topScoreResultString4 As String = String.Empty
            Dim topScoreResultKey4 As Integer = 0
            Dim topScore5 As Single = 999999
            Dim topScoreResultString5 As String = String.Empty
            Dim topScoreResultKey5 As Integer = 0
            Dim topScore6 As Single = 999999
            Dim topScoreResultString6 As String = String.Empty
            Dim topScoreResultKey6 As Integer = 0
            Dim topScore7 As Single = 999999
            Dim topScoreResultString7 As String = String.Empty
            Dim topScoreResultKey7 As Integer = 0
            Dim topScore8 As Single = 999999
            Dim topScoreResultString8 As String = String.Empty
            Dim topScoreResultKey8 As Integer = 0
            Dim topScore9 As Single = 999999
            Dim topScoreResultString9 As String = String.Empty
            Dim topScoreResultKey9 As Integer = 0
            Dim topScore10 As Single = 999999
            Dim topScoreResultString10 As String = String.Empty
            Dim topScoreResultKey10 As Integer = 0

            Dim FreqAnalysisData() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                                                "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", _
                                                "u", "v", "w", "x", "y", "z"}
            Dim FreqAnalysisScore() As Single = {8.2, 1.5, 2.8, 4.3, 12.7, 2.2, 2.0, 6.1, _
                                                 7.0, 0.15, 0.8, 4.0, 2.4, 6.7, 7.5, 1.9, _
                                                 0.1, 6.0, 6.3, 9.1, 2.8, 1.0, 2.4, 0.15, _
                                                 2.0, 0.07}

            For t As Byte = &H0 To &HFF  '******** should be &H0 To &HFF 
                Dim resultString As String = String.Empty ' use to hold string for frequency analysis
                Dim skipFreqAnalysis As Integer = 0 ' set to 1 if string/key combo has bad characters so freq analysis will be skipped
                Dim testByte As Byte = t ' we will xor this with each value in the current row
                For col As Integer = 0 To cipher.GetLength(1) - 1 '********* should be 0 to cipher....
                    '' Convert each pair of characters to a byte
                    'Dim currentByte1 As String = stringData1.Substring(i, 2) ' currentByte is two hex values
                    Dim value1 As Byte = cipher(row, col) 'value is hex byte
                    Dim tempByte As Byte = value1 Xor testByte
                    ' ************************************************************************************
                    ' Experimenting with turbo-mode - exit loop if character is out of standard text range
                    'If (tempByte >= &HC0) AndAlso (tempByte <= &HF6) Then
                    '    skipFreqAnalysis = 1
                    '    ' ******TextBox3.Text = TextBox3.Text & "**GARBAGE**"
                    '    Exit For
                    'ElseIf tempByte >= &HF8 Then
                    '    skipFreqAnalysis = 1
                    '    ' ******TextBox3.Text = TextBox3.Text & "**GARBAGE**"
                    '    Exit For
                    'ElseIf (tempByte >= &H7F) AndAlso (tempByte <= &H8E) Then
                    '    skipFreqAnalysis = 1
                    '    ' ******TextBox3.Text = TextBox3.Text & "**GARBAGE**"
                    '    Exit For
                    'Else
                    Dim com As String = ChrW(CInt(tempByte))
                    ' build a string in resultString to use for freq analysis
                    If tempByte = 0 Then
                        resultString = resultString & " " ' if this is a null, add a space
                    ElseIf tempByte = &H5E Or tempByte = &HD Or tempByte = &HA Then
                        resultString = resultString & "~" ' if this is a ^ or a return, add a ~ so it doesn't mess up Excel delimiting
                    Else
                        resultString = resultString & com ' if it is not a null, add the character
                    End If
                    'End If
                Next
                '
                '**********************************************
                '* FREQ ANALYSIS using RMSE
                '**********************************************
                ' RMSE = sqr(sum((y - ^y)^2)/n)
                ' where 
                ' y = % of occurrences of letter
                ' ^y = predicted % of occurrences of letter
                ' n = # of letters
                '
                ' for each letter a-z
                ' count # of occurrences
                ' calculate # occurrences / string length to get % of occurrences
                ' calculate (% of occurrences - FreqAnalysisScore)^2 to get square of difference between
                '      actual % and typical %
                ' after all of a-z are done, calc sum of differences (actually I'm doing a running total)
                ' calculate normalized sum = sum / n (n=26)
                ' calculate RMSE = sqr(normalized sum)
                '
                ' Update topScore1-3; smaller RMSE is better because it more closely follows the 
                '      typical distributions, so if thisSum<topScore
                ' Also, make sure topScores are initialized to some very high number
                '
                If skipFreqAnalysis = 0 Then

                    Dim resultScore As Single = 0
                    Dim letterCount As Single = 0
                    Dim percentOfOccurrences As Single = 9999
                    Dim occurrenceDiff As Single = 0
                    Dim normSum As Single = 0
                    Dim totalLetterCount As Single = 0

                    For letter As Integer = 0 To 25
                        letterCount = 0
                        For Each c As Char In resultString
                            If c = FreqAnalysisData(letter) Then
                                letterCount += 1
                                totalLetterCount += 1
                            End If
                        Next
                        percentOfOccurrences = (letterCount / resultString.Length) * 100
                        occurrenceDiff = occurrenceDiff + Math.Pow((percentOfOccurrences - FreqAnalysisScore(letter)), 2)
                    Next

                    normSum = Math.Sqrt(occurrenceDiff / 26)
                    'If totalLetterCount >= resultString.Length * 0.1 Then ' if at least 80% of string aren't alpha chars

                    If normSum < topScore1 Then
                        topScore1 = normSum
                        topScoreResultString1 = resultString
                        topScoreResultKey1 = t
                    ElseIf normSum < topScore2 Then
                        topScore2 = normSum
                        topScoreResultString2 = resultString
                        topScoreResultKey2 = t
                    ElseIf normSum < topScore3 Then
                        topScore3 = normSum
                        topScoreResultString3 = resultString
                        topScoreResultKey3 = t
                    ElseIf normSum < topScore4 Then
                        topScore4 = normSum
                        topScoreResultString4 = resultString
                        topScoreResultKey4 = t
                    ElseIf normSum < topScore5 Then
                        topScore5 = normSum
                        topScoreResultString5 = resultString
                        topScoreResultKey5 = t
                    ElseIf normSum < topScore6 Then
                        topScore6 = normSum
                        topScoreResultString6 = resultString
                        topScoreResultKey6 = t
                    ElseIf normSum < topScore7 Then
                        topScore7 = normSum
                        topScoreResultString7 = resultString
                        topScoreResultKey7 = t
                    ElseIf normSum < topScore8 Then
                        topScore8 = normSum
                        topScoreResultString8 = resultString
                        topScoreResultKey8 = t
                    ElseIf normSum < topScore9 Then
                        topScore9 = normSum
                        topScoreResultString9 = resultString
                        topScoreResultKey9 = t
                    ElseIf normSum < topScore10 Then
                        topScore10 = normSum
                        topScoreResultString10 = resultString
                        topScoreResultKey10 = t
                    End If
                    'End If

                End If
                If t = &HFF Then Exit For 'fix bug where system gives overflow after last loop where t is temporarily &hff+1
            Next
            TextBox8.Text = TextBox8.Text & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString1 & "^" & " Key: " & "^" & Chr(topScoreResultKey1) & _
                "^" & " Score: " & "^" & topScore1 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString2 & "^" & " Key: " & "^" & Chr(topScoreResultKey2) & _
                "^" & " Score: " & "^" & topScore2 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString3 & "^" & " Key: " & "^" & Chr(topScoreResultKey3) & _
                "^" & " Score: " & "^" & topScore3 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString4 & "^" & " Key: " & "^" & Chr(topScoreResultKey4) & _
                "^" & " Score: " & "^" & topScore4 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString5 & "^" & " Key: " & "^" & Chr(topScoreResultKey5) & _
                "^" & " Score: " & "^" & topScore5 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString6 & "^" & " Key: " & "^" & Chr(topScoreResultKey6) & _
                "^" & " Score: " & "^" & topScore6 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString7 & "^" & " Key: " & "^" & Chr(topScoreResultKey7) & _
                "^" & " Score: " & "^" & topScore7 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString8 & "^" & " Key: " & "^" & Chr(topScoreResultKey8) & _
                "^" & " Score: " & "^" & topScore8 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString9 & "^" & " Key: " & "^" & Chr(topScoreResultKey9) & _
                "^" & " Score: " & "^" & topScore9 & Environment.NewLine & _
                row & "^" & "code: " & "^" & topScoreResultString10 & "^" & " Key: " & "^" & Chr(topScoreResultKey10) & _
                "^" & " Score: " & "^" & topScore10
        Next
        singleByteXor = "done" '******* change later to hand results back to main sub
    End Function

End Class
