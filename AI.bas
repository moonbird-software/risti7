Attribute VB_Name = "AI"
Option Explicit

Sub AI_SelectCardToGive(Player As CardDeck)
Dim iCard As Integer, iRankPref As Integer
Dim nCards As Integer

Dim iRank(MAX_CARDS) As Integer
Dim fRules(MAX_CARDS) As Integer
Dim iPref(12) As Integer

    nCards = CountCards(Player)
    
    If nCards = 1 Then
        SelCards Player
        Exit Sub
    End If
    
    iPref(0) = 2
    iPref(1) = 3
    iPref(2) = 4
    iPref(3) = 5
    iPref(4) = Q
    iPref(5) = J
    iPref(6) = T
    iPref(7) = 9
    iPref(8) = A
    iPref(9) = K
    iPref(10) = 6
    iPref(11) = 8
    iPref(12) = 7
   
    For iCard = 0 To nCards - 1
        iRank(iCard) = GetRank(Player.Card(iCard))
        fRules(iCard) = CheckRules(Player.Card(iCard))
    Next iCard
    
    For iRankPref = 0 To 12
        For iCard = 0 To nCards - 1
            If fRules(iCard) = False And iRank(iCard) = iRankPref Then
                Player.Mode(iCard) = cmSelected
                Exit Sub
            End If
        Next iCard
    Next iRankPref
    
    iCard = Int(Rnd * nCards)
    
End Sub
Function IsCardFaceUp(Source As CardDeck, Dest As CardDeck) As Boolean
    IsCardFaceUp = IsCardFaceUpBasic(Source, Dest)
    Select Case Dest.Index
    Case IDD_USER, IDD_CLUB_6, IDD_CLUB_7, IDD_CLUB_8, IDD_DIAMOND_6, IDD_DIAMOND_7, IDD_DIAMOND_8, IDD_HEART_6, IDD_HEART_7, IDD_HEART_8, IDD_SPADE_6, IDD_SPADE_7, IDD_SPADE_8
        IsCardFaceUp = True
    End Select
End Function

Sub ActionClick(Player As CardDeck, iCard As Integer)
Dim iDeck As Integer
    Select Case Game.Mode
    Case IDM_NORMAL
        Player.Mode(iCard) = cmSelected
        iDeck = GetDestDeck(Player.Card(iCard))
        AnimMoveSelCards Player, Deck(iDeck)
        If CountCards(Deck(iDeck)) = MAX_CARDS_SHOWN_IN_6_8 And CheckHand(Player) Then
            If MsgBox(IDS_QUERY_CONTINUE_TURN, vbYesNo + vbQuestion) = vbYes Then
                SelPlayableCards Player
                Exit Sub
            End If
        End If
        
    Case IDM_GIVE
        iDeck = GetNextValidPlayer(Player.Index)
        Player.Mode(iCard) = cmSelected
        DrawDeck Player
        Delay IDT_SHOW_CARD
        AnimMoveSelCards Player, Deck(iDeck)
        SortDeck Deck(iDeck)
        Delay IDT_SHOW_CARD
        DrawDeck Deck(iDeck)
        
    Case IDM_TAKE
        Player.Mode(iCard) = cmSelected
        DrawDeck Player
        Delay IDT_SHOW_CARD
        AnimMoveSelCards Player, Deck(IDD_USER)
        SortDeck Deck(IDD_USER)
        Delay IDT_SHOW_CARD
        DrawDeck Deck(IDD_USER)
        Game.Mode = IDM_NORMAL
    
    End Select
    
    Game.FirstTurn = False
    
    RotateTurn
    
End Sub
Sub DeckClick(ByVal iDeck As Integer, Optional ByVal iCard As Integer)
    Select Case Game.Mode
    Case IDM_NORMAL
        Select Case iDeck
        Case IDD_USER
            PlaySound IDSND_CARDCLICK
            If CheckRules(Deck(iDeck).Card(iCard), True) Then
                UnSelCards Deck(iDeck)
                Deck(iDeck).Mode(iCard) = cmSelected
                DrawDeck Deck(iDeck)
                Delay IDT_SHOW_CARD
                ActionClick Deck(iDeck), iCard
            Else
                AnimFlashCard Deck(iDeck), iCard
            End If
        End Select
    
    Case IDM_GIVE
        Select Case iDeck
        Case IDD_USER
            PlaySound IDSND_CARDCLICK
            ActionClick Deck(iDeck), iCard
        End Select
        
    Case IDM_TAKE
        Select Case iDeck
        Case GetPrevValidPlayer(IDD_USER)
            PlaySound IDSND_CARDCLICK
            ActionClick Deck(iDeck), iCard
        End Select
    
    End Select
End Sub
Sub DeckKeyPress(Index As Integer, KeyAscii As Integer)
Dim iCard As Integer
Dim nCards As Integer
    nCards = CountCards(Deck(Index))
    iCard = GetKeyValue(KeyAscii)
    If iCard > 0 And iCard <= nCards Then
        DeckClick Index, iCard - 1
    End If
End Sub

Function CountPlayers() As Integer
    CountPlayers = CountPlayersBasic
End Function

Sub DealCardsDoneHook()

End Sub

Sub DealCardsHook()

End Sub
Function GetCardsInHand() As Integer
    GetCardsInHand = MAX_CARDS_IN_HAND
End Function
Function GetDestDeck(ByVal Card As Integer) As Integer
    Select Case Card
    Case cdAClubs, cd2Clubs, cd3Clubs, cd4Clubs, cd5Clubs, cd6Clubs
        GetDestDeck = IDD_CLUB_6
    Case cd7Clubs
        GetDestDeck = IDD_CLUB_7
    Case cd8Clubs, cd9Clubs, cdTClubs, cdJClubs, cdQClubs, cdKClubs
        GetDestDeck = IDD_CLUB_8
    Case cdADiamonds, cd2Diamonds, cd3Diamonds, cd4Diamonds, cd5Diamonds, cd6Diamonds
        GetDestDeck = IDD_DIAMOND_6
    Case cd7Diamonds
        GetDestDeck = IDD_DIAMOND_7
    Case cd8Diamonds, cd9Diamonds, cdTDiamonds, cdJDiamonds, cdQDiamonds, cdKDiamonds
        GetDestDeck = IDD_DIAMOND_8
    Case cdAHearts, cd2Hearts, cd3Hearts, cd4Hearts, cd5Hearts, cd6Hearts
        GetDestDeck = IDD_HEART_6
    Case cd7Hearts
        GetDestDeck = IDD_HEART_7
    Case cd8Hearts, cd9Hearts, cdTHearts, cdJHearts, cdQHearts, cdKHearts
        GetDestDeck = IDD_HEART_8
    Case cdASpades, cd2Spades, cd3Spades, cd4Spades, cd5Spades, cd6Spades
        GetDestDeck = IDD_SPADE_6
    Case cd7Spades
        GetDestDeck = IDD_SPADE_7
    Case cd8Spades, cd9Spades, cdTSpades, cdJSpades, cdQSpades, cdKSpades
        GetDestDeck = IDD_SPADE_8
    Case Else
        MsgBox Card
    End Select
End Function
Function GetNextValidPlayer(ByVal iPlr As Integer) As Integer
    GetNextValidPlayer = GetNextValidPlayerBasic(iPlr)
End Function
Function GetPrevValidPlayer(ByVal iPlr As Integer) As Integer
    GetPrevValidPlayer = GetPrevValidPlayerBasic(iPlr)
End Function

Function GiveOrTakeCards(Player As CardDeck) As Boolean
Dim iCard As Integer
Dim nCards As Integer
Dim iPrevPlr As Integer
Dim sStatus As String
Dim fFound As Boolean
    
    GiveOrTakeCards = True
    
    iPrevPlr = GetPrevValidPlayer(Player.Index)
    
    Select Case Player.Index
    Case IDD_USER
        If Rules.GiveCards Then
            ' edellinen antaa pelaajalle
            sStatus = Replace(IDS_STATUS_GIVE_TO, "%s1", Game.Title(iPrevPlr))
            sStatus = Replace(sStatus, "%s2", MakeWordTo(Game.Title(Player.Index)))
        Else
            ' pelaaja ottaa edelliseltä
            If Not Game.Demo Then
                Game.Mode = IDM_TAKE
                sStatus = Replace(IDS_STATUS_CHOOSE_CARD_TO_TAKE, "%s", MakeWordFrom(Game.Title(iPrevPlr)))
                GiveOrTakeCards = False
                SetStatus sStatus
                AnimHiliteDeck Player
                Exit Function
            End If
        End If
    
    Case Else
        Select Case iPrevPlr
        Case IDD_USER
            If Rules.GiveCards Then
                ' pelaaja antaa seuraavalle
                If Not Game.Demo Then
                    Game.Mode = IDM_GIVE
                    sStatus = Replace(IDS_STATUS_CHOOSE_CARD_TO_GIVE, "%s1", Game.Title(Player.Index))
                    sStatus = Replace(sStatus, "%s2", MakeWordTo(Game.Title(Player.Index)))
                    Game.Turn = IDD_USER
                    GiveOrTakeCards = False
                    SetStatus sStatus
                    Exit Function
                End If
            Else
                ' seuraava ottaa pelaajalta
                sStatus = Replace(IDS_STATUS_TAKE, "%s1", Game.Title(Player.Index))
                sStatus = Replace(sStatus, "%s2", MakeWordFrom(Game.Title(iPrevPlr)))
            End If
        
        Case Else
            If Rules.GiveCards Then
                ' edellinen antaa tietokoneelle
                sStatus = Replace(IDS_STATUS_GIVE_TO, "%s1", Game.Title(iPrevPlr))
                sStatus = Replace(sStatus, "%s2", MakeWordTo(Game.Title(Player.Index)))
            Else
                ' tietokone ottaa edelliseltä
                sStatus = Replace(IDS_STATUS_TAKE_FROM, "%s1", Game.Title(Player.Index))
                sStatus = Replace(sStatus, "%s2", MakeWordFrom(Game.Title(iPrevPlr)))
            End If
        End Select
    
    End Select
    
    
    ' move card
    SetStatus sStatus
    AnimHiliteDeck Player
    Delay IDT_MOVE_CARD
    Delay IDT_MOVE_CARD
    'DrawDeck Deck(iPrevPlr)
    Delay IDT_SHOW_CARD
    AI_SelectCardToGive Deck(iPrevPlr)
    DrawDeck Player
    AnimMoveSelCards Deck(iPrevPlr), Player
    SortDeck Player
    Delay IDT_SHOW_CARD
    DrawDeck Player
    
    ' check if plr who gave card is now out of them
    If CountCards(Deck(iPrevPlr)) = 0 Then
        Game.Pos(iPrevPlr) = Game.NextPos
        Game.NextPos = Game.NextPos + 1
    End If
   
End Function
Sub PlayerInput(Enabled As Boolean)
Dim iCard As Integer
Dim nCards As Integer
    PlayerInputBasic Enabled
    
    If Enabled Then
        Select Case Game.Mode
        Case IDM_NORMAL, IDM_GIVE
            frmMain.picDeck(IDD_USER).SetFocus
        Case IDM_TAKE
            frmMain.picDeck(GetPrevValidPlayer(IDD_USER)).SetFocus
        End Select
    End If
End Sub
Sub ClearGameData()
Dim iPlr As Integer
    With Game
        .New = False
        .Over = False
        
        .FirstTurn = False
        .NextPos = 1
        
        For iPlr = 0 To MAX_PLAYERS - 1
            .Pos(iPlr) = 0
        Next iPlr
        
        If .Demo Then
            SetStatus IDS_STATUS_DEMO, True
        Else
            SetStatus
        End If
    End With
End Sub
Sub InitGameData()
Dim iPlr As Integer
    With Game
        .Inited = True
        
        .Turn = IDD_USER
        .Dealer = GetPrevPlayer(IDD_USER)
        .NextPos = 1
        .RoundNbr = 1
        
        For iPlr = 0 To MAX_PLAYERS - 1
            .Score(iPlr) = 0
            .Pos(iPlr) = 0
        Next iPlr
    End With
End Sub
Function CheckHand(Player As CardDeck) As Boolean
Dim iCard As Integer
    For iCard = 0 To CountCards(Player) - 1
        If CheckRules(Player.Card(iCard)) Then
            CheckHand = True
            Exit For
        End If
    Next iCard
End Function
Sub AI_SelectCards(Player As CardDeck)
Dim iCard As Integer, iRank As Integer, iDeck As Integer
Dim nCards As Integer
Dim fSel As Boolean

    nCards = CountCards(Player)
    
    ' check for 5 to A
    For iCard = 0 To nCards - 1
        If CheckRules(Player.Card(iCard)) Then
            iRank = GetRank(Player.Card(iCard))
            If iRank >= 1 And iRank <= 5 Then
                Player.Mode(iCard) = cmSelected
                fSel = True
                Exit For
            End If
        End If
    Next iCard
    
    ' check for 9 to K
    If Not fSel Then
        For iCard = 0 To nCards - 1
            If CheckRules(Player.Card(iCard)) Then
                iRank = GetRank(Player.Card(iCard))
                If iRank >= 9 And iRank <= K Then
                    Player.Mode(iCard) = cmSelected
                    fSel = True
                    Exit For
                End If
            End If
        Next iCard
    End If
    
    ' check for 6
    If Not fSel Then
        For iCard = 0 To nCards - 1
            If CheckRules(Player.Card(iCard)) Then
                iRank = GetRank(Player.Card(iCard))
                If iRank = 6 Then
                    Player.Mode(iCard) = cmSelected
                    fSel = True
                    Exit For
                End If
            End If
        Next iCard
    End If
    
    ' check for 8
    If Not fSel Then
        For iCard = 0 To nCards - 1
            If CheckRules(Player.Card(iCard)) Then
                iRank = GetRank(Player.Card(iCard))
                If iRank = 8 Then
                    Player.Mode(iCard) = cmSelected
                    fSel = True
                    Exit For
                End If
            End If
        Next iCard
    End If
    
    ' check for 7
    If Not fSel Then
        For iCard = 0 To nCards - 1
            If CheckRules(Player.Card(iCard)) Then
                iRank = GetRank(Player.Card(iCard))
                If iRank = 7 Then
                    Player.Mode(iCard) = cmSelected
                    fSel = True
                    Exit For
                End If
            End If
        Next iCard
    End If
    
    ' animate card selection
    Player.Mode(iCard) = cmSelected
    DrawDeck Player
    iDeck = GetDestDeck(Player.Card(iCard))
    AnimMoveSelCards Player, Deck(iDeck)
    
    ' continue if possible
    If CountCards(Deck(iDeck)) = MAX_CARDS_SHOWN_IN_6_8 And CheckHand(Player) Then
        AI_SelectCards Player
    End If

End Sub
Function AI_Turn() As Boolean
Dim iPlr As Integer

    DoEvents
    
    PlayerInput False
    
    ' check if out of cards
    If CountCards(Deck(Game.Turn)) = 0 Then
        Game.Pos(Game.Turn) = Game.NextPos
        Game.NextPos = Game.NextPos + 1
    End If
    If CountPlayers = 1 Then
        iPlr = Game.Turn
        SetNextTurn iPlr
        Game.Pos(Game.Turn) = Game.NextPos
        Game.NextPos = Game.NextPos + 1
        Game.Over = True
        Exit Function
    End If
    
    ' find out if next player is CPU or human
    iPlr = Game.Turn
    SetNextTurn iPlr
    AI_Turn = IsPlayerCPU(Game.Turn)
    
    ' play cards by computer or human
    If Game.Mode = IDM_GIVE Then
        Game.Mode = IDM_NORMAL
    Else
        If CheckHand(Deck(Game.Turn)) Then
            If AI_Turn Then
                AI_SelectCards Deck(Game.Turn)
                Game.FirstTurn = False
            Else
                SelPlayableCards Deck(Game.Turn)
                PlayerInput True
            End If
        Else
            AI_Turn = GiveOrTakeCards(Deck(Game.Turn))
            If Not AI_Turn Then
                PlayerInput True
            End If
        End If
    End If
End Function
Function CheckRules(ByVal iCard As Integer, Optional ByVal fUpdateStatus As Boolean) As Boolean
Dim iRank As Integer, iRankNext
Dim iSuite As Integer
Dim sStatus As String
Dim iDeck6 As Integer, iDeck7 As Integer, iDeck8 As Integer
    
    If Game.FirstTurn Then
        If iCard <> Game.FirstCard Then
            If fUpdateStatus Then
                SetStatus IDS_STATUS_CHOOSE_7_OF_CLUBS
            End If
            Exit Function
        End If
    End If
    
    iSuite = GetSuite(iCard)
    iRank = GetRank(iCard)

    Select Case iSuite
    Case suClub
        iDeck6 = IDD_CLUB_6
        iDeck7 = IDD_CLUB_7
        iDeck8 = IDD_CLUB_8
    Case suDiamond
        iDeck6 = IDD_DIAMOND_6
        iDeck7 = IDD_DIAMOND_7
        iDeck8 = IDD_DIAMOND_8
    Case suHeart
        iDeck6 = IDD_HEART_6
        iDeck7 = IDD_HEART_7
        iDeck8 = IDD_HEART_8
    Case suSpade
        iDeck6 = IDD_SPADE_6
        iDeck7 = IDD_SPADE_7
        iDeck8 = IDD_SPADE_8
    End Select
    
    If CountCards(Deck(iDeck7)) = 0 Then
        If iRank = 7 Then
            CheckRules = True
        Else
            sStatus = Replace(IDS_STATUS_CHOOSE_7, "%s", GetSuiteName(iSuite))
        End If
    Else
        Select Case iRank
        Case Is < 7
            If CountCards(Deck(iDeck6)) = 0 Then
                If iRank = 6 Then
                    CheckRules = True
                Else
                    sStatus = Replace(IDS_STATUS_CHOOSE_6, "%s", GetSuiteName(iSuite))
                End If
            Else
                If CountCards(Deck(iDeck8)) = 0 Then
                    sStatus = Replace(IDS_STATUS_CHOOSE_8, "%s", GetSuiteName(iSuite))
                Else
                    iRankNext = GetRank(GetTopCard(Deck(iDeck6))) - 1
                    If iRank = iRankNext Then
                        CheckRules = True
                    Else
                        sStatus = Replace(IDS_STATUS_CHOOSE_IN_ORDER, "%s1", GetSuiteName(iSuite))
                        sStatus = Replace(sStatus, "%s2", GetCardName(iRankNext))
                    End If
                End If
            End If
        
        Case Is > 7
            If CountCards(Deck(iDeck6)) = 0 Then
                sStatus = Replace(IDS_STATUS_CHOOSE_6, "%s", GetSuiteName(iSuite))
            Else
                If CountCards(Deck(iDeck8)) = 0 Then
                    If iRank = 8 Then
                        CheckRules = True
                    Else
                        sStatus = Replace(IDS_STATUS_CHOOSE_8, "%s", GetSuiteName(iSuite))
                    End If
                Else
                    iRankNext = GetRank(GetTopCard(Deck(iDeck8))) + 1
                    If iRank = iRankNext Then
                        CheckRules = True
                    Else
                        sStatus = Replace(IDS_STATUS_CHOOSE_IN_ORDER, "%s1", GetSuiteName(iSuite))
                        sStatus = Replace(sStatus, "%s2", GetCardName(iRankNext))
                    End If
                End If
            End If
        End Select
    End If
    
    If Not CheckRules And fUpdateStatus Then
        SetStatus sStatus
    End If
End Function
Function GetFirstPlayer() As Integer
Dim iPlr As Integer, iCard As Integer
    For iPlr = 0 To MAX_PLAYERS - 1
        For iCard = 0 To CountCards(Deck(iPlr)) - 1
            If Deck(iPlr).Card(iCard) = cd7Clubs Then
                With Game
                    .FirstTurn = True
                    .FirstCard = cd7Clubs
                End With
                GetFirstPlayer = GetPrevPlayer(iPlr)
                Game.Turn = GetFirstPlayer
                Exit Function
            End If
        Next iCard
    Next iPlr
End Function
Sub SelPlayableCards(Player As CardDeck)
Dim iCard As Integer
    For iCard = 0 To CountCards(Player) - 1
        If CheckRules(Player.Card(iCard)) Then
            Player.Mode(iCard) = cmSelected
        End If
    Next iCard
    DrawDeck Player
End Sub
