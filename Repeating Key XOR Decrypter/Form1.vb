Public Class Form1

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
        TextBox6.Text = str ' Put the strings from the imported file into the list box
        R.Close()

    End Sub

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

        ' Find the three keysizes with the shortest hamming distances and write them to textBox2
        For lengthGuess As Integer = 2 To 40
            Dim NormalizedHammingDistance As Integer = HammingDistance(TextBox6.Text.Substring(0, lengthGuess), TextBox6.Text.Substring(lengthGuess, lengthGuess)) / lengthGuess
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
        ' Grab text to decrypt
        Dim cipher As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(TextBox6.Text) ' Convert inputted string into hex byte array

        ' Grab keysize
        Dim keysize As Integer = TextBox7.Text

        ' 2-D array tutorial: http://www.homeandlearn.co.uk/NET/nets6p5.html
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
        Dim arrayRow As Integer = (cipher.Length / keysize) - 1 ' 0 to several hundred; each block of cipher on separate row
        Dim arrayCol As Integer = keysize - 1                   ' 0 to keysize (minus 1); each column of bytes in block of cipher

        Dim cipherBlocks(,) As Byte = New Byte(arrayRow, arrayCol) {}
        Dim counter As Integer = 0
        ' keep a counter going while creating 2d array & if counter >= #items in cipher, then stop or add empty values
        ' Break text into blocks of size = keysize
        For r As Integer = 0 To arrayRow
            For c As Integer = 0 To arrayCol
                cipherBlocks(r, c) = cipher(counter)
                counter = counter + 1
                If counter >= cipher.Length - 1 Then Exit For
            Next
        Next

        ' Transpose those blocks: block1 = 1st byte of each block, block2 = 2nd byte, etc.
        Dim cipherBlocksTrans(,) As Byte = New Byte(arrayCol, arrayRow) {}
        Dim counter2 As Integer = 0
        For r2 As Integer = 0 To arrayCol
            For c2 As Integer = 0 To arrayRow
                cipherBlocksTrans(r2, c2) = cipherBlocks(c2, r2) ' ****counter shouldnt be in there
                If counter2 >= cipher.Length - 1 Then Exit For
            Next
        Next


        ' Solve each transposed block using single-byte XOR using bytes 00-FF and find byte for each block
        '     that gives best result
        ' Put single byte solutions together to make the key
        ' Use the key to decrypt original text using repeating-text XOR 

        Dim dummy As Integer = 0



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
End Class
