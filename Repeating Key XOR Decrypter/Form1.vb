Public Class Form1

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim txt1 As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(TextBox3.Text)
        Dim txt2 As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(TextBox4.Text)
        Dim btxt1 As String() = New String(txt1.Length - 1) {}
        Dim btxt2 As String() = New String(txt2.Length - 1) {}
        Dim txtPosition1 As Long = 0
        Dim txtPosition2 As Long = 0

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

        '' resultBin = bin1 XOR testByte each possible hex value 00 thru ff
        'Dim resultBin As String = String.Empty


        ''Performing XOR
        ''Doing first bit out of the loop to initialize resultBin
        'If bin1.Substring(0, 1) = testByte.Substring(0, 1) Then
        '    resultBin = "0"
        'Else
        '    resultBin = "1"
        'End If
        ''Now doing the rest of the 8 bits
        'For j As Integer = 1 To 7
        '    Dim currentBit1 As String = bin1.Substring(j, 1)
        '    Dim currentBit2 As String = testByte.Substring(j, 1)

        '    If currentBit1 = currentBit2 Then
        '        resultBin = resultBin & "0"
        '    Else
        '        resultBin = resultBin & "1"
        '    End If
        'Next



    End Sub
End Class
