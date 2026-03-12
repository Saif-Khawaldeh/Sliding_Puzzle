Public Class frmGame

    Dim picOrderArray As New ArrayList() 'To Track Images Order
    Dim IsGameRunning As Boolean = False 'Flag to Check if game is Running

    Private Sub GameTimer_Tick(sender As Object, e As EventArgs) Handles GameTimer.Tick
        'Tracks Player Time
        lblTime.Text = CInt(lblTime.Text) + 1
    End Sub

    Private Sub StartGame()

        GameTimer.Enabled = False 'Stop Any Previous Game

        lblTime.Text = 0 'Reset Player Time

        'Randomize the Images
        'For i = 1 To 1000
        '    Dim RamdomBox As Integer = Int(8 * Rnd()) + 1
        '    DoMove(RamdomBox, DetectMove(RamdomBox))
        'Next

        lblTotalMoves.Text = 0 'Reset Player Move count

        GameTimer.Enabled = True 'Start Player Time Timer
        IsGameRunning = True

    End Sub

    Private Function isWin() As Boolean

        'This function checks each box if they are in the correct order in the Picture order Array

        Dim ReturnVal As Boolean = False
        If picOrderArray(0) = 1 And picOrderArray(1) = 2 And picOrderArray(2) = 3 And
            picOrderArray(3) = 4 And picOrderArray(4) = 5 And picOrderArray(5) = 6 And
            picOrderArray(6) = 7 And picOrderArray(7) = 8 And picOrderArray(8) = 0 Then
            ReturnVal = True
        End If

        Return ReturnVal

    End Function

    Private Function DetectMove(PicNumber As Integer) As String
        ' This function checks the position of the empty space and also the
        ' position of the currently clicked box 
        ' to check for possible moves 
        ' and it is fully explained in documentaion
        Dim picPosition As Integer = picOrderArray.IndexOf(PicNumber)
        Dim emptyPosition As Integer = picOrderArray.IndexOf(0)



        If emptyPosition - picPosition = 3 Then ' Possible Bottom move
            Return "Bottom"
        ElseIf picPosition - emptyPosition = 3 Then 'Possible Top move
            Return "Top"
        ElseIf picPosition - emptyPosition = 1 And picPosition <> 0 And picPosition <> 3 _
            And picPosition <> 6 Then 'Possible Left move
            Return "Left"
        ElseIf emptyPosition - picPosition = 1 And picPosition <> 2 And picPosition <> 5 _
            And picPosition <> 8 Then 'Possible Right move
            Return "Right"
        Else
            Return "None" 'No Possible Moves

        End If



    End Function

    Private Sub DoMove(PicNumber As Integer, moveType As String)

        'This Function works after calling "DetectMove()" function
        'and moves the currently clicked box accordingly

        If moveType <> "None" Then
            '001 - Swap in Array
            Dim picPosition As Integer = picOrderArray.IndexOf(PicNumber)
            Dim emptyPosition As Integer = picOrderArray.IndexOf(0)

            Dim tempPicNumber As Integer = picOrderArray(picPosition)

            picOrderArray(picPosition) = picOrderArray(emptyPosition)
            picOrderArray(emptyPosition) = tempPicNumber


            '002 - Move Picture physically in UI

            Select Case PicNumber
                Case 1
                    If moveType = "Top" Then
                        PictureBox1.Top = PictureBox1.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox1.Top = PictureBox1.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox1.Left = PictureBox1.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox1.Left = PictureBox1.Left + 150
                    End If


                Case 2
                    If moveType = "Top" Then
                        PictureBox2.Top = PictureBox2.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox2.Top = PictureBox2.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox2.Left = PictureBox2.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox2.Left = PictureBox2.Left + 150
                    End If


                Case 3

                    If moveType = "Top" Then
                        PictureBox3.Top = PictureBox3.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox3.Top = PictureBox3.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox3.Left = PictureBox3.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox3.Left = PictureBox3.Left + 150
                    End If

                Case 4
                    If moveType = "Top" Then
                        PictureBox4.Top = PictureBox4.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox4.Top = PictureBox4.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox4.Left = PictureBox4.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox4.Left = PictureBox4.Left + 150
                    End If

                Case 5
                    If moveType = "Top" Then
                        PictureBox5.Top = PictureBox5.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox5.Top = PictureBox5.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox5.Left = PictureBox5.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox5.Left = PictureBox5.Left + 150
                    End If

                Case 6
                    If moveType = "Top" Then
                        PictureBox6.Top = PictureBox6.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox6.Top = PictureBox6.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox6.Left = PictureBox6.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox6.Left = PictureBox6.Left + 150
                    End If

                Case 7

                    If moveType = "Top" Then
                        PictureBox7.Top = PictureBox7.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox7.Top = PictureBox7.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox7.Left = PictureBox7.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox7.Left = PictureBox7.Left + 150
                    End If

                Case 8

                    If moveType = "Top" Then
                        PictureBox8.Top = PictureBox8.Top - 150
                    ElseIf moveType = "Bottom" Then
                        PictureBox8.Top = PictureBox8.Top + 150
                    ElseIf moveType = "Left" Then
                        PictureBox8.Left = PictureBox8.Left - 150
                    ElseIf moveType = "Right" Then
                        PictureBox8.Left = PictureBox8.Left + 150
                    End If


            End Select


            lblTotalMoves.Text = CInt(lblTotalMoves.Text) + 1 'Increment the Player's Total Moves Label

        End If



    End Sub

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Runs only One Time, Once the program is Loaded
        LoadBoxesArray()
    End Sub

    Private Sub LoadBoxesArray()
        'This Procedure runs only once the program is loaded
        'It assigns the values in the Pictures order array
        picOrderArray.Add(1)
        picOrderArray.Add(2)
        picOrderArray.Add(3)
        picOrderArray.Add(4)
        picOrderArray.Add(5)
        picOrderArray.Add(6)
        picOrderArray.Add(7)
        picOrderArray.Add(8)
        picOrderArray.Add(0) ' Zero is The Empty Box Area
    End Sub

    Private Sub EndGame()

        If IsGameRunning Then
            IsGameRunning = False
            GameTimer.Enabled = False
            MsgBox("Conggratsss!!!!  :)" & vbCrLf & "Total Moves : " & lblTotalMoves.Text _
                   & vbCrLf & "Time Spent : " & lblTime.Text & " Seconds",, "You Have Won !!!")

        End If


    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If IsGameRunning Then DoMove(1, DetectMove(1))
        If isWin() Then EndGame()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If IsGameRunning Then DoMove(2, DetectMove(2))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        If IsGameRunning Then DoMove(3, DetectMove(3))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        If IsGameRunning Then DoMove(4, DetectMove(4))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        If IsGameRunning Then DoMove(5, DetectMove(5))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        If IsGameRunning Then DoMove(6, DetectMove(6))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        If IsGameRunning Then DoMove(7, DetectMove(7))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        If IsGameRunning Then DoMove(8, DetectMove(8))
        If isWin() Then EndGame()

    End Sub

    Private Sub PictureBoxStart_Click(sender As Object, e As EventArgs) Handles PictureBoxStart.Click
        StartGame()
    End Sub


End Class

