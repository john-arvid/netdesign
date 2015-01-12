Module Rulesets

    Public Sub AddOrUpdateRuleSet()

        Dim RuleSet As Visio.ValidationRuleSet
        Dim RuleSetName As String
        Dim Document As Visio.Document

        Document = Globals.ThisAddIn.Application.ActiveDocument
        RuleSetName = "NetDesign"
        RuleSet = GetRuleSet(Document, RuleSetName)
        If RuleSet Is Nothing Then
            RuleSet = Document.Validation.RuleSets.Add(RuleSetName)
        End If

        RuleSet.Name = "Net Design"
        RuleSet.Description = "Net Design rule set"
        RuleSet.Enabled = True
        RuleSet.RuleSetFlags = Visio.VisRuleSetFlags.visRuleSetDefault

        Call AddDefaultRules(RuleSet, Document)

    End Sub

    Public Function GetRule(ByVal ruleSet As Visio.ValidationRuleSet, ByVal ruleName As String) As Visio.ValidationRule

        Dim ReturnValue As Visio.ValidationRule = Nothing
        Dim Rule As Visio.ValidationRule

        For Each Rule In ruleSet.Rules
            If UCase(Rule.NameU) = UCase(ruleName) Then
                ReturnValue = Rule
                Exit For
            End If
        Next
        GetRule = ReturnValue
    End Function

    Private Function GetRuleSet(ByVal doc As Visio.Document, ByVal ruleSetName As String) As Visio.ValidationRuleSet

        Dim ReturnValue As Visio.ValidationRuleSet = Nothing
        Dim RuleSet As Visio.ValidationRuleSet

        For Each RuleSet In doc.Validation.RuleSets
            If UCase(RuleSet.NameU) = UCase(ruleSetName) Then
                ReturnValue = RuleSet
                Exit For
            End If
        Next
        GetRuleSet = ReturnValue
    End Function

End Module