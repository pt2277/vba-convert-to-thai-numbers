Attribute VB_Name = "NewMacros"
Function ReplaceText(Source As String, Dest As String)
'
' ReplaceText
' Written by Papoj Thamjaroenporn
' 2021/07/08
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Source
        .Replacement.Text = Dest
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function
Sub ConvertToThaiNumbers()
Attribute ConvertToThaiNumbers.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' ConvertToThaiNumbers Macro
' Written by Papoj Thamjaroenporn
' 2021/07/08
'
    Call ReplaceText("1", ChrW(3665))
    Call ReplaceText("2", ChrW(3666))
    Call ReplaceText("3", ChrW(3667))
    Call ReplaceText("4", ChrW(3668))
    Call ReplaceText("5", ChrW(3669))
    Call ReplaceText("6", ChrW(3670))
    Call ReplaceText("7", ChrW(3671))
    Call ReplaceText("8", ChrW(3672))
    Call ReplaceText("9", ChrW(3673))
    Call ReplaceText("0", ChrW(3664))
End Sub
