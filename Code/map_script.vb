'MULTIPLE RANGE TEST

Sub UpdatePartial()

Dim WS As Worksheet
Dim MatchCase As Boolean
With ActiveSheet.UsedRange

.Replace "MarsGameHorde", "Assault"
.Replace "{MarsGameHorde8P}", "Massive Assault"
.Replace "MarsGameSurvival", "Kill Every Thing"
.Replace "MarsGameSurvival2", "Kill Every Thing 2"
.Replace "MarsGameAST", "Assault Operations"
.Replace "MarsGameZAST", "Paranormal Operations"
.Replace "MarsGameDefence", "Threshold Defense"
.Replace "MarsGameMonster", "Extinction Ops"
.Replace "MarsGameCM", "Training"
.Replace "MarsGameQRM", "Team Deathmatch"
.Replace "MarsGameTDM", "Elimination"
.Replace "MarsGameDM", "Free-for-All"
.Replace "MarsGameAnnex", "King of the Hill"
.Replace "MarsGameDefuse", "Demolition"
.Replace "MarsGameBighead", "Head Hunter"
.Replace "MarsGameMVZ", "Mercs vs. Monsters"
.Replace "MarsGameInfection", "Patient Zero"

For Each WS In Worksheets
        WS.Cells.Replace What:=Search, Replacement:=Replacement, _
        LookAt:=xlPart, MatchCase:=True
Next
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

