'MULTIPLE RANGE TEST

Sub UpdatePartial()

Set rColumnRange = Range("D2", "D75")

With ActiveSheet.UsedRange
.Replace "{MarsGameHorde}", "{Assault}", xlWhole
.Replace "{MarsGameHorde8P}", "{Massive Assault}", xlWhole
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
End With

End Sub