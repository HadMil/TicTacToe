Public Class Form1
    Dim turn As Integer
    Sub New()
        InitializeComponent()
        'Setup the game board ready to start
        InitialiseGameBoard()
    End Sub

    ''' <summary>
    ''' Handles when a game button has been selected
    ''' </summary>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click

        Dim ButtonSelected As Button = CType(sender, Button)
        'Set the vlaues for this button when it is selected
        SetupTurnValues(ButtonSelected)
        'Check to see if the game has been won
        CheckMoveForWin(ButtonSelected)
    End Sub

    ''' <summary>
    ''' Handles when the user selects to restart the game again the game again
    ''' </summary>
    Private Sub Reset_Click(sender As Object, e As EventArgs) Handles reset.Click
        'Setup the game board ready to start again
        InitialiseGameBoard()
    End Sub

    ''' <summary>
    ''' Initialises the game board ready to start the game
    ''' </summary>
    Private Sub InitialiseGameBoard()
        Button1.Text = " "
        Button2.Text = " "
        Button3.Text = " "
        Button4.Text = " "
        Button5.Text = " "
        Button6.Text = " "
        Button7.Text = " "
        Button8.Text = " "
        Button9.Text = " "
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Button9.Enabled = True
        turn = 1
        Me.BackColor = Color.Red
    End Sub

    ''' <summary>
    ''' Time to check to see if the game has ended in a draw.
    ''' </summary>
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Me.Timer.Enabled = False
        'If we have got to this turn and no one has won then we can assume it's a draw.
        If turn = 9 Then
            MsgBox("Draw")
            turn += 1
        End If
        Me.Timer.Enabled = True
    End Sub

    ''' <summary>
    ''' Set the text for the button selected, disable the button from being selected and what background colour needs to be next.
    ''' </summary>
    Private Sub SetupTurnValues(ButtonSelected As Button)
        'Set the button text and background colour for the next go 
        If turn Mod 2 = 1 Then
            ButtonSelected.Text = "X"
            Me.BackColor = Color.Blue
        Else
            ButtonSelected.Text = "O"
            Me.BackColor = Color.Red
        End If
        'Update the turn counter and disable the button
        turn += 1
        ButtonSelected.Enabled = False
    End Sub

    ''' <summary>
    ''' Check to see if that last move has won the game 
    ''' </summary>
    Private Sub CheckMoveForWin(ButtonSelected As Button)

        Dim GameWon As Boolean = False

        Select Case ButtonSelected.Name
            Case "Button1"
                If Button1.Text = Button2.Text AndAlso Button1.Text = Button3.Text Then
                    GameWon = True
                ElseIf Button1.Text = Button4.Text AndAlso Button1.Text = Button7.Text Then
                    GameWon = True
                ElseIf Button1.Text = Button5.Text AndAlso Button1.Text = Button9.Text Then
                    GameWon = True
                End If
            Case "Button2"
                If Button1.Text = Button2.Text AndAlso Button1.Text = Button3.Text Then
                    GameWon = True
                ElseIf Button2.Text = Button5.Text AndAlso Button2.Text = Button8.Text Then
                    GameWon = True
                End If
            Case "Button3"
                If Button1.Text = Button2.Text AndAlso Button1.Text = Button3.Text Then
                    GameWon = True
                ElseIf Button3.Text = Button6.Text AndAlso Button3.Text = Button9.Text Then
                    GameWon = True
                ElseIf Button3.Text = Button5.Text AndAlso Button3.Text = Button7.Text Then
                    GameWon = True
                End If
            Case "Button4"
                If Button4.Text = Button5.Text AndAlso Button4.Text = Button6.Text Then
                    GameWon = True
                ElseIf Button1.Text = Button4.Text AndAlso Button1.Text = Button7.Text Then
                    GameWon = True
                End If
            Case "Button5"
                If Button4.Text = Button5.Text AndAlso Button4.Text = Button6.Text Then
                    GameWon = True
                ElseIf Button2.Text = Button5.Text AndAlso Button2.Text = Button8.Text Then
                    GameWon = True
                ElseIf Button1.Text = Button5.Text AndAlso Button1.Text = Button9.Text Then
                    GameWon = True
                ElseIf Button3.Text = Button5.Text AndAlso Button3.Text = Button7.Text Then
                    GameWon = True
                End If
            Case "Button6"
                If Button4.Text = Button5.Text And Button4.Text = Button6.Text Then
                    GameWon = True
                ElseIf Button3.Text = Button6.Text And Button3.Text = Button9.Text Then
                    GameWon = True
                End If
            Case "Button7"
                If Button1.Text = Button4.Text AndAlso Button1.Text = Button7.Text Then
                    GameWon = True
                ElseIf Button3.Text = Button5.Text AndAlso Button3.Text = Button7.Text Then
                    GameWon = True
                ElseIf Button7.Text = Button8.Text AndAlso Button7.Text = Button9.Text Then
                    GameWon = True
                End If
            Case "Button8"
                If Button7.Text = Button8.Text AndAlso Button7.Text = Button9.Text Then
                    GameWon = True
                ElseIf Button2.Text = Button5.Text AndAlso Button2.Text = Button8.Text Then
                    GameWon = True
                End If
            Case "Button9"
                If Button7.Text = Button8.Text AndAlso Button7.Text = Button9.Text Then
                    GameWon = True
                ElseIf Button1.Text = Button5.Text AndAlso Button1.Text = Button9.Text Then
                    GameWon = True
                ElseIf Button3.Text = Button6.Text AndAlso Button3.Text = Button9.Text Then
                    GameWon = True
                End If
        End Select

        'The move has won the game so show the message and disable the board
        If (GameWon) Then
            MsgBox(" '" & ButtonSelected.Text & "' Wins")
            DisableBoard()
        End If

    End Sub

    ''' <summary>
    ''' Disable all the buttons for the game
    ''' </summary>
    Private Sub DisableBoard()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click

        Dim msg As String = "This implementation of Tic Tac Toe was developed using VB.Net in Visual Studio 2019, enjoy."
        MessageBox.Show(msg, "About", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

End Class