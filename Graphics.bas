Attribute VB_Name = "Graphics"
Option Explicit

Sub DrawCard(Obj As Object, Deck As CardDeck, iCard As Integer, nCards As Integer, X As Integer, Y As Integer)
    DrawCardBasic Obj, Deck, iCard, nCards, X, Y
    If nCards = 0 Then
        Select Case Deck.Index
        Case IDD_DEALER
            cdtDraw Obj.hdc, X, Y, 0, mdDeckX, IDC_TABLEBG
        Case IDD_CLUB_6, IDD_CLUB_7, IDD_CLUB_8, IDD_DIAMOND_6, IDD_DIAMOND_7, IDD_DIAMOND_8, IDD_HEART_6, IDD_HEART_7, IDD_HEART_8, IDD_SPADE_6, IDD_SPADE_7, IDD_SPADE_8
            GetCardXY Deck, iCard, nCards, X, Y
            cdtDraw Obj.hdc, X, Y, 0, mdGhost, IDC_TABLEBG
        End Select
    Else
        Select Case Deck.Index
        Case IDD_DEALER
            cdtDraw Obj.hdc, X, Y, Game.CardBack, mdFaceDown, IDC_TABLEBG
        Case IDD_CLUB_6, IDD_CLUB_7, IDD_CLUB_8, IDD_DIAMOND_6, IDD_DIAMOND_7, IDD_DIAMOND_8, IDD_HEART_6, IDD_HEART_7, IDD_HEART_8, IDD_SPADE_6, IDD_SPADE_7, IDD_SPADE_8
            If nCards = MAX_CARDS_SHOWN_IN_6_8 Then
                cdtDraw Obj.hdc, X, Y, Game.CardBack, mdFaceDown, IDC_TABLEBG
            Else
                cdtDraw Obj.hdc, X, Y, Deck.Card(iCard), mdFaceUp, IDC_TABLEBG
            End If
        End Select
    End If
End Sub
Sub FormMainResize()
    FormMainResizeBasic
End Sub
Function GetCardStep(Deck As CardDeck, ByVal nCards As Integer, Optional ByRef fStepX As Boolean, Optional ByRef fReverse As Boolean) As Integer
    GetCardStep = GetCardStepBasic(Deck, nCards, fStepX, fReverse)
    Select Case Deck.Index
    Case IDD_DEALER
        GetCardStep = 8
    Case IDD_CLUB_6, IDD_CLUB_8, IDD_DIAMOND_6, IDD_DIAMOND_8, IDD_HEART_6, IDD_HEART_8, IDD_SPADE_6, IDD_SPADE_8
        GetCardStep = cdHeight / 16
    End Select
End Function
Sub GetCardXY(Deck As CardDeck, ByVal iCard As Integer, ByVal nCards As Integer, ByRef X As Integer, ByRef Y As Integer)
Dim iStep As Integer, iPos As Integer
Dim fStepX As Boolean, fReverse As Boolean
    
    GetCardXYBasic Deck, iCard, nCards, X, Y

    iStep = GetCardStep(Deck, nCards, fStepX, fReverse)
    iPos = iStep * iCard
    
    Select Case Deck.Index
    Case IDD_DEALER
        X = iCard / iStep
        Y = iCard / iStep
    Case IDD_CLUB_6, IDD_DIAMOND_6, IDD_HEART_6, IDD_SPADE_6
        X = 0
        Y = iStep * (MAX_CARDS_SHOWN_IN_6_8 - 1 - iCard)
    Case IDD_CLUB_7, IDD_DIAMOND_7, IDD_HEART_7, IDD_SPADE_7
        X = 0
        Y = 0
    Case IDD_CLUB_8, IDD_DIAMOND_8, IDD_HEART_8, IDD_SPADE_8
        X = 0
        Y = iPos
    End Select
    
End Sub
