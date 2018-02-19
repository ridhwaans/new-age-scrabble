Attribute VB_Name = "mdlGlobal"
'This module holds almost all of the declarations and definitions of the major variables used in New-Age Scrabble
'The variables below are public scope, so they can be accessed or modified anywhere for convinience

Public strPlayer(0 To 1) As String    'Player Name, 0 is player one, 1 is player two
Public intPlayerTiles(0 To 1, 0 To 6) As Integer    'The letter numbers (example: A= 0) of each of the player's seven tiles
Public intPlayerScore(0 To 1) As Integer

Public strMostValuable As String    'used for determining the most valuable word for the AI to play
Public intMostValuable As Integer
Public intComputerTries As Integer
Public strPivot As String
Public strWordComputer As String

Public intTurn As Integer    'Indicates whose turn it is
Public intWordsPlayed(0 To 1) As Integer
Public blnTwoPlayer As Boolean    'Two player mode

Public intGridTiles(0 To 224) As Integer    'The letter number of each of the game board tiles
Public blnGridTilesPlayed(0 To 224) As Boolean    'This boolean indicates if the tile is locked (already played) or unlocked for each of the game board tile

'Statistic variable for game summary
Public intBestWord As Integer
Public intGameTimeElapsed As Integer
Public strGameSummary(0 To 5) As String
'0-Winner, 1-Best word, 2-Tiles played, 3-Tiles left,4-Time elapsed, 5-Number of words played

'Time elapsed variables
Public intSeconds As Integer
Public intMinutes As Integer
Public strSeconds As String
Public strMinutes As String

Public intCount As Integer    'intCount variable used in most for loops

Public intTile As Integer    'Holds the letter number of the tile being played
Public intTempTile As Integer
Public intBlankTile(0 To 1) As Integer    'Holds the grid position location of where the blank tile was played

Public intTileBank(0 To 99) As Integer    'Compilation of all generated tiles in one array
Public intTilesInPlay(0 To 1, 0 To 6) As Integer    'Set of tile addresses per player

Public intTilesRemaining(0 To 26) As Integer    'Quantity remaining for each of the 26 letters of the alphabet
Public intTileValue(0 To 26) As Integer    'Score value for each of the 26 letters of the alphabet
Public strTileValue(0 To 26) As String    'String value (example: "A", "B" etc) for each of the 26 letters
Public intTotalTiles As Integer    'Used to store the total number of unused tiles remaining in the game
Public strTilePath As String    'The file path string name of the picture for the letter

Public blnExchangeDisabled As Boolean    'for turn only
Public intScore As Integer    'Temporary variable used to store the word score

Public strGameStandings(0 To 9)    'Used for highscores
Public intGameStandings(0 To 9)    'Used for highscores

Public Sub EndGame()
'Update game summary statistics
    If intPlayerScore(0) > intPlayerScore(1) Then    'Player 1 wins
        strGameSummary(0) = "Winner: " & strPlayer(0)
    ElseIf intPlayerScore(0) > intPlayerScore(1) Then    'Player 2 wins
        strGameSummary(0) = "Winner: " & strPlayer(1)
    ElseIf intPlayerScore(0) = intPlayerScore(1) Then    'Tie
        strGameSummary(0) = "Game Tied"
    End If

    Dim intTilesPlayed As Integer
    For intCount = 0 To 224
        If blnGridTilesPlayed(intCount) = True Then
            intTilesPlayed = intTilesPlayed + 1
        End If
    Next intCount

    strGameSummary(2) = "Tiles played: " & Str(intTilesPlayed)    'Tiles played
    strGameSummary(3) = "Tiles left: " & Str(100 - intTilesPlayed)    'Tiles left
    strGameSummary(4) = "Time elapsed: " & Format(Int(intGameTimeElapsed / 3600), "00") & ":" & Format(Int((intGameTimeElapsed Mod 3600) / 60), "00") & ":" & Format(Int(((intGameTimeElapsed Mod 3600) Mod 60)), "00")    'Total game time elapsed
    strGameSummary(5) = "Words played: " & Str(intWordsPlayed(0) + intWordsPlayed(1))    'Total number of words played

    For intCount = 0 To 5
        frmGameOver.lblGameSummary.Caption = frmGameOver.lblGameSummary.Caption & strGameSummary(intCount) & vbNewLine
    Next intCount

    Call updateHighScores
    
    frmGameOver.lblPlayer1.Caption = strPlayer(0)
    frmGameOver.lblPlayer2.Caption = strPlayer(1)
    frmGameOver.lblPlayerScore(0).Caption = intPlayerScore(0)
    frmGameOver.lblPlayerScore(1).Caption = intPlayerScore(1)
End Sub

Public Sub updateHighScores()
'update highscores here
    Dim hsIndex As Integer
    Dim hsName() As String
    Dim hsScore() As Integer

    hsIndex = 0
    ReDim hsName(0 To hsIndex)
    ReDim hsScore(0 To hsIndex)

    'Get the highscores from the file
    intHighScores = FreeFile
    Open App.Path & "\game_files\highscores_screen\highscores.txt" For Input As intHighScores
    Do
        Input #intHighScores, hsName(hsIndex), hsScore(hsIndex)
        hsIndex = hsIndex + 1
        ReDim Preserve hsName(0 To hsIndex)
        ReDim Preserve hsScore(0 To hsIndex)
    Loop Until EOF(intHighScores)

    Close #intHighScores

    hsIndex = hsIndex - 1

    ReDim Preserve hsName(0 To hsIndex)
    ReDim Preserve hsScore(0 To hsIndex)

    'Update highscores
    For Outer = 0 To 1
        For intCount = 0 To hsIndex
            If intPlayerScore(Outer) > hsScore(intCount) Then    'If score is higher
                For nested = (intCount + 1) To hsIndex    'shift array down
                    hsName(nested) = hsName(nested - 1)
                    hsScore(nested) = hsScore(nested - 1)
                Next nested
                hsName(intCount) = strPlayer(Outer)    'insert new highscore
                hsScore(intCount) = intPlayerScore(Outer)
                GoTo check_next:
            End If
        Next intCount
check_next:
    Next Outer

    Kill App.Path & "\game_files\highscores_screen\highscores.txt"

    'Put the highscores back
    intHighScores = FreeFile
    Open App.Path & "\game_files\highscores_screen\highscores.txt" For Append As intHighScores
    For intCount = 0 To hsIndex
        Write #intHighScores, hsName(intCount), hsScore(intCount)
    Next intCount

    Close #intHighScores
End Sub
