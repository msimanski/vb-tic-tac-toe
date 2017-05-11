Public Class Form
    Dim boardFirst(2, 2) As Integer 'class level game board array, 0 is null, 1 is x, 2 is o, X IS PLAYER, O IS CPU

    Public Structure gameMove 'makes a struct to record game moves, i use this like once, i did it cause i could

        Public row As Integer
        Public col As Integer

    End Structure

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load 'init the game

        Randomize(DateTime.Now.Millisecond) 'place an O at a random place
        Dim x, y As Integer
        x = CInt(Int((3 * Rnd()) + 0))
        Randomize(DateTime.Now.Millisecond)
        y = CInt(Int((3 * Rnd()) + 0))
        boardFirst(x, y) = 2

        changeLetters() 'refresh board

    End Sub

    Function anyMovesLeft(ByVal b(,) As Integer) As Boolean 'returns true if a blank spot is on the board

        For i As Integer = 0 To 2
            For j As Integer = 0 To 2

                If b(i, j) = 0 Then
                    Return True
                End If

            Next
        Next

    End Function

    Function scoreBoard(ByVal b(,) As Integer) As Integer 'gives a score for a board, 0 for no winner or draw, 10 for player win, -10 for cpu win

        For row As Integer = 0 To 2 'traverse rows

            If (b(row, 0) = b(row, 1)) And (b(row, 1) = b(row, 2)) Then 'check rows

                If b(row, 0) = 1 Then
                    Return 10
                ElseIf b(row, 0) = 2 Then
                    Return -10
                End If

            End If

        Next

        For col As Integer = 0 To 2 'traverse cols


            If (b(0, col) = b(1, col)) And (b(1, col) = b(2, col)) Then 'check cols

                If b(0, col) = 1 Then
                    Return 10
                ElseIf b(0, col) = 2 Then
                    Return -10
                End If

            End If


        Next


        If (b(0, 0) = b(1, 1)) And (b(1, 1) = b(2, 2)) Then 'check diagonals


            If b(0, 0) = 1 Then
                Return 10
            ElseIf b(0, 0) = 2 Then
                Return -10
            End If


        End If


        If (b(0, 2) = b(1, 1)) And (b(1, 1) = b(2, 0)) Then 'check diagonals

            If b(0, 2) = 1 Then
                Return 10
            ElseIf b(0, 2) = 2 Then
                Return -10
            End If

        End If

        Return 0

    End Function

    Function miniMax(ByVal b(,) As Integer, ByVal depth As Integer, ByVal isMax As Boolean) As Integer 'main algorithm, used since 40s, found it on Wikipedia (https://en.wikipedia.org/wiki/Minimax) or something, helps get best move for cpu

        Dim score As Integer = scoreBoard(b) 'gets score of current board


        If score = 10 Then 'player won
            Return score
        End If


        If score = -10 Then 'cpu won
            Return score
        End If


        If anyMovesLeft(b) = False Then 'draw
            Return 0
        End If


        If isMax = True Then 'minimax assuming player turn

            Dim best As Integer = -1000

            For i As Integer = 0 To 2
                For j As Integer = 0 To 2
                    If b(i, j) = 0 Then
                        b(i, j) = 1 'do move

                        best = Math.Max(best, miniMax(b, depth + 1, Not (isMax))) 'see if its any good

                        b(i, j) = 0 'reset

                    End If
                Next
            Next

            Return best 'send it back

        Else 'minimax assuming cpu turn

            Dim best As Integer = 1000

            For i As Integer = 0 To 2
                For j As Integer = 0 To 2

                    If b(i, j) = 0 Then

                        b(i, j) = 2 'do move

                        best = Math.Max(best, miniMax(b, depth + 1, Not (isMax))) 'see if its any good

                        b(i, j) = 0 'reset

                    End If

                Next
            Next

            Return best 'send it back

        End If

    End Function

    Function findBestMove(ByVal b(,) As Integer) As gameMove 'implementation of minimax, finds best move for cpu given board

        Dim bestVal = -1000
        Dim bestMove As gameMove
        bestMove.row = -1
        bestMove.col = -1

        For i As Integer = 0 To 2 'traverse cells
            For j As Integer = 0 To 2

                If b(i, j) = 0 Then

                    b(i, j) = 1 'do move

                    Dim moveVal As Integer = miniMax(b, 0, False) 'get heuristic

                    b(i, j) = 0 'reset

                    If moveVal > bestVal Then 'see if this is what cpu wants to do based on heuristic

                        bestMove.row = i
                        bestMove.col = j
                        bestVal = moveVal

                    End If


                End If

            Next
        Next

        Return bestMove 'send it back

    End Function

    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click 'button event handler
        boardFirst(0, 0) = 1
        gameTurn()
    End Sub

    Private Sub btn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2.Click 'button event handler
        boardFirst(1, 0) = 1
        gameTurn()
    End Sub

    Private Sub btn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn3.Click 'button event handler
        boardFirst(2, 0) = 1
        gameTurn()
    End Sub

    Private Sub btn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn4.Click 'button event handler
        boardFirst(0, 1) = 1
        gameTurn()
    End Sub

    Private Sub btn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn5.Click 'button event handler
        boardFirst(1, 1) = 1
        gameTurn()
    End Sub

    Private Sub btn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn6.Click 'button event handler
        boardFirst(2, 1) = 1
        gameTurn()
    End Sub

    Private Sub btn7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn7.Click 'button event handler
        boardFirst(0, 2) = 1
        gameTurn()
    End Sub

    Private Sub btn8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn8.Click 'button event handler
        boardFirst(1, 2) = 1
        gameTurn()
    End Sub

    Private Sub btn9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn9.Click 'button event handler
        boardFirst(2, 2) = 1
        gameTurn()
    End Sub

    Sub gameTurn() 'does a turn of the game

        If (scoreBoard(boardFirst) = 0) Then 'if game is still on

            changeLetters() 'refresh just to be sure
            Dim aimove As gameMove = findBestMove(boardFirst) 'get best move for cpu
            Dim x, y As Integer
            x = aimove.row
            y = aimove.col
            boardFirst(x, y) = 2 'set it in board array
            changeLetters() 'refresh

        End If

        If scoreBoard(boardFirst) = 10 Then 'if you won

            MessageBox.Show("A winner is you!")
            resetGame()

        End If

        If scoreBoard(boardFirst) = -10 Then 'if you lost

            MessageBox.Show("You lost!")
            resetGame()

        End If

        If anyMovesLeft(boardFirst) = False Then 'if you draw

            MessageBox.Show("Draw")
            resetGame()

        End If

    End Sub

    Sub changeLetters() 'refreshes board, just a bunch of unrolled for loops, nothing to see here - move along

        If boardFirst(0, 0) = 2 Then

            btn1.Text = "O"
            btn1.Enabled = False

        End If

        If boardFirst(1, 0) = 2 Then

            btn2.Text = "O"
            btn2.Enabled = False

        End If

        If boardFirst(2, 0) = 2 Then

            btn3.Text = "O"
            btn3.Enabled = False

        End If

        If boardFirst(0, 1) = 2 Then

            btn4.Text = "O"
            btn4.Enabled = False

        End If

        If boardFirst(1, 1) = 2 Then

            btn5.Text = "O"
            btn5.Enabled = False

        End If

        If boardFirst(2, 1) = 2 Then

            btn6.Text = "O"
            btn6.Enabled = False

        End If

        If boardFirst(0, 2) = 2 Then

            btn7.Text = "O"
            btn7.Enabled = False

        End If

        If boardFirst(1, 2) = 2 Then

            btn8.Text = "O"
            btn8.Enabled = False

        End If

        If boardFirst(2, 2) = 2 Then

            btn9.Text = "O"
            btn9.Enabled = False

        End If

        If boardFirst(0, 0) = 1 Then

            btn1.Text = "X"
            btn1.ForeColor = Color.Blue

        End If

        If boardFirst(1, 0) = 1 Then

            btn2.Text = "X"
            btn2.ForeColor = Color.Blue

        End If

        If boardFirst(2, 0) = 1 Then

            btn3.Text = "X"
            btn3.ForeColor = Color.Blue

        End If

        If boardFirst(0, 1) = 1 Then

            btn4.Text = "X"
            btn4.ForeColor = Color.Blue

        End If

        If boardFirst(1, 1) = 1 Then

            btn5.Text = "X"
            btn5.ForeColor = Color.Blue

        End If

        If boardFirst(2, 1) = 1 Then

            btn6.Text = "X"
            btn6.ForeColor = Color.Blue

        End If

        If boardFirst(0, 2) = 1 Then

            btn7.Text = "X"
            btn7.ForeColor = Color.Blue

        End If

        If boardFirst(1, 2) = 1 Then

            btn8.Text = "X"
            btn8.ForeColor = Color.Blue

        End If

        If boardFirst(2, 2) = 1 Then

            btn9.Text = "X"
            btn9.ForeColor = Color.Blue

        End If

        If boardFirst(0, 0) = 0 Then

            btn1.Text = " "
            btn1.Enabled = True

        End If

        If boardFirst(1, 0) = 0 Then

            btn2.Text = " "
            btn2.Enabled = True

        End If

        If boardFirst(2, 0) = 0 Then

            btn3.Text = " "
            btn3.Enabled = True

        End If

        If boardFirst(0, 1) = 0 Then

            btn4.Text = " "
            btn4.Enabled = True

        End If

        If boardFirst(1, 1) = 0 Then

            btn5.Text = " "
            btn5.Enabled = True

        End If

        If boardFirst(2, 1) = 0 Then

            btn6.Text = " "
            btn6.Enabled = True

        End If

        If boardFirst(0, 2) = 0 Then

            btn7.Text = " "
            btn7.Enabled = True

        End If

        If boardFirst(1, 2) = 0 Then

            btn8.Text = " "
            btn8.Enabled = True

        End If

        If boardFirst(2, 2) = 0 Then

            btn9.Text = " "
            btn9.Enabled = True

        End If


    End Sub

    Sub resetGame() 'starts a new game

        For i As Integer = 0 To 2
            For j As Integer = 0 To 2

                boardFirst(i, j) = 0 'make all cells null

            Next
        Next

        Randomize(DateTime.Now.Millisecond) 'place an O at random place
        Dim x, y As Integer
        x = CInt(Int((3 * Rnd()) + 0))
        Randomize(DateTime.Now.Millisecond)
        y = CInt(Int((3 * Rnd()) + 0))
        boardFirst(x, y) = 2

        changeLetters() 'refresh screen

    End Sub

End Class
