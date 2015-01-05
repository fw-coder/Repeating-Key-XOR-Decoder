Public Class Form1

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        TextBox5.Text = HammingDistance(TextBox3.Text, TextBox4.Text).ToString ' write hamming distance to result box
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
    Private Sub OpenFileDialog2_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        Dim strm As System.IO.Stream

        strm = OpenFileDialog2.OpenFile()

        TextBox6.Text = OpenFileDialog2.FileName.ToString()

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If

        ' Import strings from text file
        Dim R As New IO.StreamReader(OpenFileDialog2.FileName)
        Dim str As String = R.ReadToEnd() ' Delimiter is vbLF (LineFeed)

        cipherInput.fromB64File = System.Text.ASCIIEncoding.ASCII.GetBytes(str)

        ' ******** convert cipherInput.fromB64File from base64 to hex bytes
        cipherInput.fromFile = Convert.FromBase64String(str)


        RichTextBox1.Text = str ' Put the strings from the imported file into the list box
        R.Close()
    End Sub

    Public Class cipherInput

        Public Shared fromFile() As Byte
        Public Shared fromB64File() As Byte

    End Class

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select a Hex File"
        OpenFileDialog1.InitialDirectory = "C:temp"
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        OpenFileDialog2.Title = "Please Select a Base64 File"
        OpenFileDialog2.InitialDirectory = "C:temp"
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim bestLengthGuess() As Integer = {0, 0, 0, 0, 0}
        Dim shortestNormalizedDistance() As Single = {9999, 9999, 9999, 9999, 9999}
        Dim cipherString As String = System.Text.ASCIIEncoding.ASCII.GetString(cipherInput.fromFile)

        ' Find the three keysizes with the shortest hamming distances and write them to textBox2
        For lengthGuess As Integer = 2 To 160 'should be 2 to 40
            Dim NormalizedHammingDistance As Single = 0
            Dim loops As Integer = (cipherString.Length - (lengthGuess * 2)) + 1
            ProgressBar1.Value = lengthGuess
            For i As Integer = 1 To loops
                NormalizedHammingDistance = NormalizedHammingDistance + _
                    HammingDistance(cipherString.Substring(0, lengthGuess), _
                    cipherString.Substring(lengthGuess, lengthGuess)) / lengthGuess
            Next
            Dim aveNHD As Single = NormalizedHammingDistance / loops
            For s As Integer = 0 To 4
                If aveNHD < shortestNormalizedDistance(s) Then
                    ' shuffle
                    For sh As Integer = 4 To s + 1 Step -1
                        bestLengthGuess(sh) = bestLengthGuess(sh - 1)
                        shortestNormalizedDistance(sh) = shortestNormalizedDistance(sh - 1)
                    Next
                    bestLengthGuess(s) = lengthGuess
                    shortestNormalizedDistance(s) = aveNHD
                    Exit For
                End If
            Next
        Next

        TextBox2.Text = _
            "L1=" & bestLengthGuess(0) & " (" & shortestNormalizedDistance(0) & ")" & Environment.NewLine & _
            "L2=" & bestLengthGuess(1) & " (" & shortestNormalizedDistance(1) & ")" & Environment.NewLine & _
            "L3=" & bestLengthGuess(2) & " (" & shortestNormalizedDistance(2) & ")" & Environment.NewLine & _
            "L4=" & bestLengthGuess(3) & " (" & shortestNormalizedDistance(3) & ")" & Environment.NewLine & _
            "L5=" & bestLengthGuess(4) & " (" & shortestNormalizedDistance(4) & ")" & Environment.NewLine

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
            Dim lowScore() As Single = {999999, 999999, 999999, 999999, 999999, 999999, 999999, 999999, 999999, 999999}
            Dim lowScoreResultString(10) As String
            Dim lowScoreResultKey() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
            Dim FreqAnalysisData() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                                                "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", _
                                                "u", "v", "w", "x", "y", "z"}
            Dim FreqAnalysisScore() As Single = {8.2, 1.5, 2.8, 4.3, 12.7, 2.2, 2.0, 6.1, _
                                                 7.0, 0.15, 0.8, 4.0, 2.4, 6.7, 7.5, 1.9, _
                                                 0.1, 6.0, 6.3, 9.1, 2.8, 1.0, 2.4, 0.15, _
                                                 2.0, 0.07}
            ProgressBar2.Maximum = cipher.GetLength(0) - 1
            ProgressBar2.Value = row
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

                    For s As Integer = 0 To 9
                        If normSum < lowScore(s) Then
                            ' shuffle
                            For sh As Integer = 9 To s + 1 Step -1
                                lowScore(sh) = lowScore(sh - 1)
                                lowScoreResultString(sh) = lowScoreResultString(sh - 1)
                                lowScoreResultKey(sh) = lowScoreResultKey(sh - 1)
                            Next
                            lowScore(s) = normSum
                            lowScoreResultString(s) = resultString
                            lowScoreResultKey(s) = t
                            Exit For
                        End If
                    Next
                    'End If

                End If
                If t = &HFF Then Exit For 'fix bug where system gives overflow after last loop where t is temporarily &hff+1
            Next

            For b As Integer = 0 To 9
                TextBox8.Text = TextBox8.Text & _
                row & "^" & "code: " & "^" & lowScoreResultString(b) & "^" & " Key: " & "^" & Chr(lowScoreResultKey(b)) & _
                "^" & " Score: " & "^" & lowScore(b) & Environment.NewLine
            Next

        Next
        singleByteXor = "done" '******* change later to hand results back to main sub
    End Function

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        ' Getting repeating key and text to encrypt
        Dim key As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(TextBox10.Text)
        Dim fileOut(cipherInput.fromFile.Length) As Byte
        Dim counter As Integer = 0

        For i As Integer = 0 To cipherInput.fromFile.Length - 1 Step key.Length
            For j As Integer = 0 To key.Length - 1
                Dim out As Byte = cipherInput.fromFile(counter) Xor key(j)
                fileOut(counter) = out
                counter = counter + 1
                If counter = cipherInput.fromFile.Length Then Exit For
            Next
        Next
        RichTextBox2.Text = System.Text.ASCIIEncoding.ASCII.GetString(fileOut)
    End Sub
End Class
