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
        TextBox6.Text = str ' Put the strings from the imported file into the list box
        R.Close()

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = "C:temp"
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        For lengthGuess As Integer = 2 To 40
            TextBox2.Text = TextBox2.Text & "L=" & lengthGuess & ": " & HammingDistance(TextBox6.Text.Substring(0, lengthGuess), TextBox6.Text.Substring(lengthGuess, lengthGuess)) & Environment.NewLine
        Next

    End Sub
End Class
