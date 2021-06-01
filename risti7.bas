Attribute VB_Name = "Risti7"
Option Explicit

' Revision History
'
'   0.0.16  First Public Release    3.7.2002
'
'   0.9.0   9.1.2003
'
'   - sound support added
'   - playable cards are now selected at the beginning of human turn
'   - implemented card selection AI
'   - implemented card giving AI
'   - new game query is no longer always shown in frmsettings
'
'   0.9.3   3.4.2003
'
'   - changed default font to Tahoma
'   - last plr no longer gets 1 point from losing
'   - if plr gives his last card to someone, he now gets a score
'   - sound setting is now saved when changed in file menu
'   - game score and round nbr are zeroed when switching to/from demo mode
'
'   1.0.0   18.10.2003
'
'   - frmSettings: changed antopeli description slightly
'   - fixed human player being cpu instead
'
' TODO:

' RUNKO definitions
Public Const MAX_PLAYERS = 4
Public Const MAX_DECKS = 18
Public Const MAX_CARDS_IN_HAND = 13
Public Const MAX_CARDS_SHOWN_IN_6_8 = 6
Public Const MIN_MAIN_WIDTH = 10575
Public Const MIN_MAIN_HEIGHT = 10605
Public Const RUNKO_APP = 1

' rules
Type RuleBook
    GiveCards As Boolean
End Type

' deck constants
Public Const IDD_PLAYER1 = 0
Public Const IDD_PLAYER2 = 1
Public Const IDD_PLAYER3 = 2
Public Const IDD_PLAYER4 = 3
Public Const IDD_DEALER = 4
Public Const IDD_CLUB_6 = 5
Public Const IDD_CLUB_7 = 9
Public Const IDD_CLUB_8 = 13
Public Const IDD_DIAMOND_6 = 6
Public Const IDD_DIAMOND_7 = 10
Public Const IDD_DIAMOND_8 = 14
Public Const IDD_HEART_6 = 7
Public Const IDD_HEART_7 = 11
Public Const IDD_HEART_8 = 15
Public Const IDD_SPADE_6 = 8
Public Const IDD_SPADE_7 = 12
Public Const IDD_SPADE_8 = 16
Public Const IDD_SPRITE = 17
Public Const IDD_USER = IDD_PLAYER1

Public Const IDD_TRICK = -1
Public Const IDD_TRASH = -2

' game modes
Public Const IDM_NORMAL = 0
Public Const IDM_GIVE = 1
Public Const IDM_TAKE = 2
Sub FormSettingsLoad()
    With frmSettings
        ' rules
        If Rules.GiveCards Then
            .optRule(0).Value = True
        Else
            .optRule(1).Value = True
        End If
    End With
    
    FormSettingsLoadBasic
End Sub
Sub FormSettingsSave()
    With frmSettings
        ' prompt to start new game if rules have changed
        If .optRule(0).Value <> Rules.GiveCards Then
            If .cmdCancel.Enabled Then
                If MsgBox(IDS_QUERY_RESTART_GAME, vbOKCancel + vbQuestion) = vbOK Then
                    Game.New = True
                Else
                    Exit Sub
                End If
            End If
        End If
        ' rules
        Rules.GiveCards = .optRule(0).Value
    End With
    FormSettingsSaveBasic
End Sub
Sub InitLocale()
Dim iPlr As Integer

    InitLocaleBasic

    ' settings
    With frmSettings
        .optRule(0).Caption = IDS_DLG_SETTINGS_RULE_0
        .optRule(1).Caption = IDS_DLG_SETTINGS_RULE_1
    End With
End Sub
Sub ReadSettings()
    ' rules
    Rules.GiveCards = GetSetting(App.Title, "Rules", "GiveCards", False)
    Game.CardSortOrder = csoSuit
    
    ReadSettingsBasic
    UpdateDebug
End Sub
Sub SaveSettings()
    SaveSettingsBasic
    
    ' rules
    SaveSetting App.Title, "Rules", "GiveCards", Rules.GiveCards
End Sub
Sub UpdateDebug()
    With frmMain
    End With
End Sub
