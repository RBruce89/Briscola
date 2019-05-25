Option Explicit On
Option Strict On
Option Infer Off

Public Class Form1

    Private cardsLeft As Integer
    Private deck(39) As String
    Private player1Cards(2, 1) As String
    Private player2Cards(2, 1) As String
    Private p1DisplayNum(2) As Integer
    Private p2DisplayNum(2) As Integer

    Private playCard(1) As String
    Private winner As Integer
    Private playerScore(1) As Integer

    Private bri As String
    Private briSuit As String
    Private briNum As Integer

    Private activeGame As Boolean


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub

    Private Sub deckPictureBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles deckPictureBox.Click
        'Start new game

        'Confirm starting a new game when a game is in progress
        If activeGame = False Then
            Call startGame()
        Else
            If MessageBox.Show("Start a new game?", "New Game", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                Call startGame()
            End If
        End If

    End Sub

    Private Sub startGame()
        'Clear old game and start new one

        'Variables for randomly picking bri
        Dim randGen As New Random
        Dim briRand As Integer

        'Reset variables
        For rowClear As Integer = 0 To player1Cards.GetUpperBound(0)
            For colClear As Integer = 0 To player1Cards.GetUpperBound(1)
                player1Cards(rowClear, colClear) = Nothing
            Next colClear
        Next rowClear

        For rowClear As Integer = 0 To player2Cards.GetUpperBound(0)
            For colClear As Integer = 0 To player2Cards.GetUpperBound(1)
                player2Cards(rowClear, colClear) = Nothing
            Next colClear
        Next rowClear

        playCard(0) = Nothing
        playCard(1) = Nothing
        playerScore(0) = 0
        playerScore(1) = 0
        bri = Nothing
        briNum = 0
        activeGame = True

        deck = {"B02", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", _
                "C02", "C04", "C05", "C06", "C07", "C08", "C09", "C10", "C11", "C12", _
                "D02", "D04", "D05", "D06", "D07", "D08", "D09", "D10", "D11", "D12", _
                "S02", "S04", "S05", "S06", "S07", "S08", "S09", "S10", "S11", "S12"}

        'Reset picture box images
        deckPictureBox.BringToFront()
        deckLastPictureBox.Visible = True
        briPictureBox.Visible = True
        deckLastPictureBox.Image = cardsImageList.Images.Item(40)
        p2Card1PictureBox.Image = Nothing
        p2Card2PictureBox.Image = Nothing
        p2Card3PictureBox.Image = Nothing
        briPictureBox.Image = Nothing
        p1Card1PictureBox.Image = Nothing
        p1Card2PictureBox.Image = Nothing
        p1Card3PictureBox.Image = Nothing
        playCard1PictureBox.Image = Nothing
        playCard2PictureBox.Image = Nothing
        Me.Refresh()

        'Reset labels
        playersPointsLabel.Text = playerScore(0).ToString
        computersPointsLabel.Text = playerScore(1).ToString
        winnerLabel.Visible = False

        'Deal first cards
        Call CardDealer()
        Call pause(2)
        Me.Refresh()

        'Drawl and display Bri
        briRand = randGen.Next(0, 40)
        Do While deck(briRand) = Nothing
            briRand = randGen.Next(0, 40)
        Loop
        bri = deck(briRand)
        briNum = briRand
        deck(briRand) = Nothing
        briPictureBox.Image = cardsImageList.Images.Item(briRand)
        briSuit = bri.Remove(1)

        If winner = 1 Then
            Call ComputerPlay()
        End If

        'Enable player card controls
        p1Card1PictureBox.Enabled = True
        p1Card2PictureBox.Enabled = True
        p1Card3PictureBox.Enabled = True

    End Sub

    Private Sub CardCounter()
        'Count remaining cards and update deck images

        'Sets cardsLeft variable based on how many cards are left in the deck
        cardsLeft = 0
        For cardCounter As Integer = 0 To 39
            If deck(cardCounter) <> Nothing Then
                cardsLeft += 1
            End If
        Next cardCounter

        'Hides card picture boxes when there aren't enough cards left to be in them
        If cardsLeft < 3 Then
            If cardsLeft = 0 Then
                briPictureBox.Visible = False
                deckPictureBox.Image = Nothing
                deckLastPictureBox.Visible = False
            Else
                deckPictureBox.Image = Nothing
                deckLastPictureBox.BringToFront()
            End If
        End If

    End Sub

    Private Sub CardDealer()
        'Deal a card to every empty slot for both player and computer then display them

        'Variables to set random card and control loops
        Dim cardUsed As Boolean
        Dim randGen As New Random
        Dim randNum As Integer

        'Deal cards to player 1
        Call CardCounter()
        cardUsed = True
        If cardsLeft > 0 Or bri <> Nothing Then
            For deal1 As Integer = 0 To player1Cards.GetUpperBound(0)
                Do While cardUsed = True AndAlso cardsLeft > 0
                    randNum = randGen.Next(0, 40)
                    If deck(randNum) <> Nothing Then
                        cardUsed = False
                    Else
                        cardUsed = True
                    End If
                Loop
                If player1Cards(deal1, 0) = Nothing AndAlso cardsLeft > 1 Then
                    player1Cards(deal1, 0) = deck(randNum)
                    player1Cards(deal1, 1) = randNum.ToString
                    deck(randNum) = Nothing
                    cardUsed = True
                ElseIf player1Cards(deal1, 0) = Nothing AndAlso winner = 1 Then
                    player1Cards(deal1, 0) = bri
                    player1Cards(deal1, 1) = briNum.ToString
                    bri = Nothing
                    cardUsed = True
                ElseIf player1Cards(deal1, 0) = Nothing AndAlso winner = 0 Then
                    player1Cards(deal1, 0) = deck(randNum)
                    player1Cards(deal1, 1) = randNum.ToString
                    deck(randNum) = Nothing
                    cardUsed = True
                End If
                Call CardCounter()
            Next deal1
        End If

        'Deal cards to player 2
        Call CardCounter()
        cardUsed = True
        If cardsLeft > 0 Or bri <> Nothing Then
            For deal2 As Integer = 0 To player2Cards.GetUpperBound(0)
                Do While cardUsed = True AndAlso cardsLeft > 0
                    randNum = randGen.Next(0, 40)
                    If deck(randNum) <> Nothing Then
                        cardUsed = False
                    Else
                        cardUsed = True
                    End If
                Loop
                If player2Cards(deal2, 0) = Nothing AndAlso cardsLeft > 1 Then
                    player2Cards(deal2, 0) = deck(randNum)
                    player2Cards(deal2, 1) = randNum.ToString
                    deck(randNum) = Nothing
                    cardUsed = True
                ElseIf player2Cards(deal2, 0) = Nothing AndAlso winner = 0 Then
                    player2Cards(deal2, 0) = bri
                    player2Cards(deal2, 1) = briNum.ToString
                    bri = Nothing
                    cardUsed = True
                ElseIf player2Cards(deal2, 0) = Nothing AndAlso winner = 1 Then
                    player2Cards(deal2, 0) = deck(randNum)
                    player2Cards(deal2, 1) = randNum.ToString
                    deck(randNum) = Nothing
                    cardUsed = True
                End If
                Call CardCounter()
            Next deal2
        End If

        'Use a number to track player1 cards for display
        For display1 As Integer = 0 To player1Cards.GetUpperBound(0)
            Integer.TryParse(player1Cards(display1, 1), p1DisplayNum(display1))
        Next display1

        'Use a number to track player2 cards for display
        For display2 As Integer = 0 To player2Cards.GetUpperBound(0)
            Integer.TryParse(player2Cards(display2, 1), p2DisplayNum(display2))
        Next display2

        'Display cards
            p2Card1PictureBox.Image = cardsImageList.Images.Item(40)
        Call pause(1)
            Me.Refresh()
            p2Card2PictureBox.Image = cardsImageList.Images.Item(40)
        Call pause(1)
            Me.Refresh()
            p2Card3PictureBox.Image = cardsImageList.Images.Item(40)
        Call pause(1)
            Me.Refresh()
            p1Card1PictureBox.Image = cardsImageList.Images.Item(p1DisplayNum(0))
        Call pause(1)
            Me.Refresh()
            p1Card2PictureBox.Image = cardsImageList.Images.Item(p1DisplayNum(1))
        Call pause(1)
            Me.Refresh()
            p1Card3PictureBox.Image = cardsImageList.Images.Item(p1DisplayNum(2))
        Call pause(1)
            Me.Refresh()

    End Sub

    Private Sub CardEvaluator()


        'Name and set variables to evaluate cards
        Dim card1Suit As String = playCard(0).Remove(1)
        Dim card2Suit As String = playCard(1).Remove(1)
        Dim cardNum(1) As Integer
        Integer.TryParse(playCard(0).Remove(0, 1), cardNum(0))
        Integer.TryParse(playCard(1).Remove(0, 1), cardNum(1))

        'Determine winner
        If card1Suit = card2Suit Then
            If cardNum(0) > cardNum(1) Then
                winner = 0
            Else
                winner = 1
            End If
        ElseIf card1Suit = briSuit Then
            winner = 0
        ElseIf card2Suit = briSuit Then
            winner = 1
        End If

        'Assign appropriate points to winner
        For points As Integer = 0 To 1
            Select Case cardNum(points)
                Case 8
                    playerScore(winner) += 2
                Case 9
                    playerScore(winner) += 3
                Case 10
                    playerScore(winner) += 4
                Case 11
                    playerScore(winner) += 10
                Case 12
                    playerScore(winner) += 11
            End Select
        Next

        'Display and update labels
        If winner = 0 Then
            winnerLabel.Text = "Player takes hand!"
        ElseIf winner = 1 Then
            winnerLabel.Text = "Computer takes hand!"
        End If
        winnerLabel.Visible = True
        playersPointsLabel.Text = playerScore(0).ToString
        computersPointsLabel.Text = playerScore(1).ToString
        Me.Refresh()

        Call pause(20)

        'Remove played cards
        playCard1PictureBox.Image = Nothing
        playCard2PictureBox.Image = Nothing
        playCard(0) = Nothing
        playCard(1) = Nothing
        winnerLabel.Visible = False

        Me.Refresh()
        Call pause(3)

        'Determine if game is over and declare winner
        If player1Cards(0, 0) = Nothing AndAlso player1Cards(1, 0) = Nothing AndAlso player1Cards(2, 0) = Nothing AndAlso
            player2Cards(0, 0) = Nothing AndAlso player2Cards(1, 0) = Nothing AndAlso player2Cards(2, 0) = Nothing Then
            If playerScore(0) > playerScore(1) Then
                winnerLabel.Text = "Player Wins!"
            ElseIf playerScore(0) < playerScore(1) Then
                winnerLabel.Text = "Computer Wins!"
            Else
                winnerLabel.Text = "Draw!"
            End If
            winnerLabel.Visible = True
            activeGame = False
            deckPictureBox.Image = cardsImageList.Images.Item(40)
            Me.Refresh()
        Else
            'Determine if cards need delt
            Call CardCounter()
            If cardsLeft > 0 Then
                Call CardDealer()
            End If
            'Pass control to computer if it won
            If winner = 1 Then
                Call ComputerPlay()
            End If

            'Re-enable player control for boxes that have cards
            If player1Cards(0, 0) <> Nothing Then
                p1Card1PictureBox.Enabled = True
            End If
            If player1Cards(1, 0) <> Nothing Then
                p1Card2PictureBox.Enabled = True
            End If
            If player1Cards(2, 0) <> Nothing Then
                p1Card3PictureBox.Enabled = True
            End If


        End If

    End Sub

    Private Sub pause(ByVal tenthSeconds As Integer)
        For time As Integer = 0 To tenthSeconds
            System.Threading.Thread.Sleep(100)
            Application.DoEvents()
        Next time
    End Sub

    Private Sub PlayerCard_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles p1Card1PictureBox.Click, p1Card2PictureBox.Click, p1Card3PictureBox.Click
        'Play user card then either let computer play, or evaluate cards

        'Determine the clicked card
        Dim clickedBox As PictureBox = TryCast(sender, PictureBox)
        Dim clickedBoxText As String = clickedBox.ToString

        'Disable cards
        p1Card1PictureBox.Enabled = False
        p1Card2PictureBox.Enabled = False
        p1Card3PictureBox.Enabled = False

        'Put card image from hand onto table
        playCard1PictureBox.Image = clickedBox.Image
        clickedBox.Image = Nothing
        Me.Refresh()

        'Tranfer card from hand to eval area
        Select Case clickedBox.AccessibleName
            Case "p1c1"
                playCard(0) = player1Cards(0, 0)
                player1Cards(0, 0) = Nothing
                player1Cards(0, 1) = Nothing
            Case "p1c2"
                playCard(0) = player1Cards(1, 0)
                player1Cards(1, 0) = Nothing
                player1Cards(1, 1) = Nothing
            Case "p1c3"
                playCard(0) = player1Cards(2, 0)
                player1Cards(2, 0) = Nothing
                player1Cards(2, 1) = Nothing
        End Select

        'Determine if the computer should play, or if it's time to eval
        If playCard(1) = Nothing Then
            Call ComputerPlay()
        Else
            winner = 1
            Call CardEvaluator()
        End If

    End Sub

    Private Sub ComputerPlay()
        'Computer selects a card and plays it

        'Pause for .5 seconds between player's card pick and computer's
        Call pause(5)

        'Variables for computer to select cards
        Dim selectedCard As Integer
        Dim play1 As Boolean = True
        Dim play3 As Boolean = True

        'Variables to track value of cards
        Dim handCardNum(2) As Integer
        Dim handCardSuit(2) As String
        Dim playerCardSuit As String
        Dim playerCardNum As Integer

        'To set hand variables
        For cardValues As Integer = 0 To player2Cards.GetUpperBound(0)
            If player2Cards(cardValues, 0) <> Nothing Then
                handCardSuit(cardValues) = player2Cards(cardValues, 0).Remove(1)
                Integer.TryParse(player2Cards(cardValues, 0).Remove(0, 1), handCardNum(cardValues))
            End If
        Next cardValues

        'Set variables to track Bris and 1s
            For cardsLeft As Integer = 0 To 39
                If deck(cardsLeft) <> Nothing Then
                    If deck(cardsLeft).Remove(1) = briSuit Then
                        play1 = False
                        play3 = False
                    End If
                    If deck(cardsLeft).Remove(0, 1) = "12" Then
                        play3 = False
                    End If
                End If
        Next cardsLeft
        If play1 = True Then
            For cardsLeft As Integer = 0 To player1Cards.GetUpperBound(0)
                If player1Cards(cardsLeft, 0) <> Nothing Then
                    If player1Cards(cardsLeft, 0).Remove(1) = briSuit Then
                        play1 = False
                        play3 = False
                    End If
                    If player1Cards(cardsLeft, 0).Remove(0, 1) = "12" Then
                        play3 = False
                    End If
                End If
            Next cardsLeft
        End If

        'If computer goes first
        If winner = 1 Then
            'Try to play 1 or 3
            If play1 = True Then
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = 12 AndAlso handCardSuit(findType) <> briSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            ElseIf play3 = True Then
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = 11 AndAlso handCardSuit(findType) <> briSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            End If
        'Check for non-Bri face card
            For findNum As Integer = 8 To 10
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            Next findNum
        'Check for non-Bri low card
            For findNum As Integer = 2 To 7
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            Next findNum
        'Check for Bri low card or face card
            For findNum As Integer = 2 To 10
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = findNum Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            Next findNum
        'Check for non-Bri 3 or 1
            For findNum As Integer = 11 To 12
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            Next findNum
        'Check for Bri 3 or 1
            For findNum As Integer = 11 To 12
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = findNum Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            Next findNum

        'If computer goes second
        ElseIf winner = 0 Then

            'Set player card variables
        playerCardSuit = playCard(0).Remove(1)
        Integer.TryParse(playCard(0).Remove(0, 1), playerCardNum)

            'Card down is any Bri but 3
        If playerCardSuit = briSuit AndAlso playerCardNum <> 11 Then
            'Check for non-Bri low card
                For findNum As Integer = 2 To 7
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for low Bri
                For findNum As Integer = 2 To 7
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for non-Bri face card
                For findNum As Integer = 8 To 10
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for Bri face card
                For findNum As Integer = 8 To 10
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
                'Check to see if played card is 1
            If playerCardNum = 12 Then
                    'Check for non-Bri 3 or 1
                    For findNum As Integer = 11 To 12
                        For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                            If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                                selectedCard = findType
                                GoTo Play
                            End If
                        Next findType
                    Next findNum
                'Check for Bri 3
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = 11 AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
            Else
                    'Check for Bri 3 or 1
                    For findNum As Integer = 11 To 12
                        For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                            If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                                selectedCard = findType
                                GoTo Play
                            End If
                        Next findType
                    Next findNum
                    'Check for non-Bri 3 or 1
                    For findNum As Integer = 11 To 12
                        For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                            If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                                selectedCard = findType
                                GoTo Play
                            End If
                        Next findType
                    Next findNum
                End If
            End If

            'Card down is Bri 3
        If playerCardNum = 11 AndAlso playerCardSuit = briSuit Then
            'Check for Bri 1
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = 12 AndAlso handCardSuit(findType) = briSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            'Check for non-Bri low card
                For findNum As Integer = 2 To 7
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for low Bri
                For findNum As Integer = 2 To 7
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for non-Bri face card
                For findNum As Integer = 8 To 10
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for Bri face card
                For findNum As Integer = 8 To 10
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for non-Bri 1 or 3
                For findNum As Integer = 11 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
        End If

            'Card down is non-Bri 1
        If playerCardNum = 12 AndAlso playerCardSuit <> briSuit Then
            'Check for Bri
                For findNum As Integer = 2 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for non-Bri
                For findNum As Integer = 2 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
        End If

            'Card down is Non-Bri 3
        If playerCardNum = 11 AndAlso playerCardSuit <> briSuit Then
                'Check for 1 of matching class
                For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                    If handCardNum(findType) = 12 AndAlso handCardSuit(findType) = playerCardSuit Then
                        selectedCard = findType
                        GoTo Play
                    End If
                Next findType
            'Check for Bri
                For findNum As Integer = 2 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            'Check for non-Bri
                For findNum As Integer = 2 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
        End If

            'Card down is non-Bri low or face
        If playerCardNum < 11 AndAlso playerCardSuit <> briSuit Then
                'Check for 1 or 3 of matching suit
                For findNum As Integer = 11 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = playerCardSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
                'Check for non-Bri lower than player card
                For findNum As Integer = 2 To playerCardNum
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
                'Check for non-Bri non matching suit
                For findNum As Integer = playerCardNum To 10
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> playerCardSuit AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
                'Check for matching suit higher then player card
                For findNum As Integer = playerCardNum To 10
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = playerCardSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
                'Check for Bri
                For findNum As Integer = 2 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) = briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
                'Check for non-Bri 3 or 1
                For findNum As Integer = 11 To 12
                    For findType As Integer = 0 To player2Cards.GetUpperBound(0)
                        If handCardNum(findType) = findNum AndAlso handCardSuit(findType) <> briSuit Then
                            selectedCard = findType
                            GoTo Play
                        End If
                    Next findType
                Next findNum
            End If

        End If


Play:
        'Play and show selected card
        Select Case selectedCard
            Case 0
                playCard(1) = player2Cards(0, 0)
                player2Cards(0, 0) = Nothing
                player2Cards(0, 1) = Nothing
                p2Card1PictureBox.Image = Nothing
                playCard2PictureBox.Image = cardsImageList.Images(p2DisplayNum(0))
            Case 1
                playCard(1) = player2Cards(1, 0)
                player2Cards(1, 0) = Nothing
                player2Cards(1, 1) = Nothing
                p2Card2PictureBox.Image = Nothing
                playCard2PictureBox.Image = cardsImageList.Images(p2DisplayNum(1))
            Case 2
                playCard(1) = player2Cards(2, 0)
                player2Cards(2, 0) = Nothing
                player2Cards(2, 1) = Nothing
                p2Card3PictureBox.Image = Nothing
                playCard2PictureBox.Image = cardsImageList.Images(p2DisplayNum(2))
        End Select

        'Show computers card on screen
        Me.Refresh()

        'If player's card is down, send to eval
        If playCard(0) <> Nothing Then
            Call CardEvaluator()
        End If

    End Sub

End Class
