'MULTIPLE RANGE TEST

Sub UpdatePartial()

Dim WS As Worksheet
Dim MatchCase As Boolean
With ActiveSheet.UsedRange

.Replace "MarsGameHorde", "Assault", xlPart
.Replace "{MarsGameHorde8P}", "Massive Assault", xlPart
.Replace "MarsGameSurvival", "Kill Every Thing", xlPart
.Replace "MarsGameSurvival2", "Kill Every Thing 2", xlPart
.Replace "MarsGameAST", "Assault Operations", xlPart
.Replace "MarsGameZAST", "Paranormal Operations", xlPart
.Replace "MarsGameDefence", "Threshold Defense", xlPart
.Replace "MarsGameMonster", "Extinction Ops", xlPart
.Replace "MarsGameCM", "Training", xlPart
.Replace "MarsGameQRM", "Team Deathmatch", xlPart
.Replace "MarsGameTDM", "Elimination", xlPart
.Replace "MarsGameDM", "Free-for-All", xlPart
.Replace "MarsGameAnnex", "King of the Hill", xlPart
.Replace "MarsGameDefuse", "Demolition", xlPart
.Replace "MarsGameBighead", "Head Hunter", xlPart
.Replace "MarsGameMVZ", "Mercs vs. Monsters", xlPart
.Replace "MarsGameInfection", "Patient Zero", xlPart

For Each WS In Worksheets
        WS.Cells.Replace MatchCase:=True

End With
End Sub

'SINGLE REPLACE CODE

Sub ChgInfo() 
     
    Dim WS              As Worksheet 
    Dim Search          As String 
    Dim Replacement     As String 
    Dim Prompt          As String 
    Dim Title           As String 
    Dim MatchCase       As Boolean 
     
    Prompt = "What is the original value you want to replace?"
    Title = "Search Value Input"
    Search = InputBox(Prompt, Title)
     
    Prompt = "What is the replacement value?"
    Title = "Search Value Input"
    Replacement = InputBox(Prompt, Title)
     
    For Each WS In Worksheets
        WS.Cells.Replace What:=Search, Replacement:=Replacement, _
        LookAt:=xlPart, MatchCase:=False
    Next
     
End Sub

