Module AddRules

    Public Sub AddDefaultRules(ByRef ruleSet As Visio.ValidationRuleSet, _
                               ByRef document As Visio.Document)

        Dim Rule As Visio.ValidationRule
        Dim RuleName As String

        RuleName = "NoShapesOnPage"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Shapes"
        Rule.Description = "A page must atleast contain one shape"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetPage
        Rule.FilterExpression = ""
        Rule.TestExpression = "Aggcount(ShapesOnPage())>0"

        RuleName = "OneWirePerPort"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Shapes"
        Rule.Description = "A port should only have one connected wire"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = "hascategory(""Copper"")"
        Rule.TestExpression = "aggcount(gluedshapes(0))<2"

        RuleName = "WireConnected"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Shapes"
        Rule.Description = "A wire need to be connected in both ends"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = "strsame(left(mastername(750), 5),""cable"")"
        Rule.TestExpression = "or(and(aggcount(gluedshapes(0)) = 1, aggcount(gluedshapes(3)) = 1), aggcount(gluedshapes(0)) = 2, aggcount(gluedshapes(3)) = 2)"


        RuleName = "Pro-Switch"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Hierarchy"
        Rule.Description = "Wire goes from Processor to Switch, not allowed"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        RuleName = "Pro-Pro"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Hierarchy"
        Rule.Description = "Wire goes from Processor to Processor, not allowed"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        RuleName = "WireLoop"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Connection"
        Rule.Description = "Wire goes in a loop"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        RuleName = "MediaType"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Connection"
        Rule.Description = "Wire does not have same media type as port"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        RuleName = "MediaPurpose"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Connection"
        Rule.Description = "Wire does not have same media purpose as port"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        RuleName = "MediaSpeed"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Connection"
        Rule.Description = "Wire does not have same media speed as port"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        RuleName = "UniqueName"

        Rule = GetRule(ruleSet, RuleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(RuleName)
        End If
        Rule.Category = "Shape"
        Rule.Description = "This shape does not have a unique name"
        Rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape
        Rule.FilterExpression = ""
        Rule.TestExpression = ""

        ' This will be the proper way of adding a rule, but can also be with
        ' string variables that have the information instead of the string
        ' itself
        Call AddRule(ruleSet, "OPCToOPC", "Connection", _
                     "A wire cannot be connected to a OPC in both ends", _
                     Visio.VisRuleTargets.visRuleTargetShape)

    End Sub

    Public Function GetIssue(ByVal rule As Visio.ValidationRule, ByVal targetPage As Visio.Page, ByVal targetShape As Visio.Shape) As Visio.ValidationIssue

        Dim ReturnValue As Visio.ValidationIssue = Nothing
        Dim Issue As Visio.ValidationIssue

        ' Goes through every issue in the document, returns the issue if it exist
        For Each Issue In Globals.ThisAddIn.Application.ActiveDocument.Validation.Issues
            If rule.TargetType = Visio.VisRuleTargets.visRuleTargetShape And Not targetShape Is Nothing Then
                If targetShape Is Issue.TargetShape Then
                    ReturnValue = Issue
                    Exit For
                End If
            ElseIf rule.TargetType = Visio.VisRuleTargets.visRuleTargetPage And Not targetPage Is Nothing Then
                If targetPage Is Issue.TargetPage Then
                    ReturnValue = Issue
                    Exit For
                End If
            ElseIf rule.TargetType = Visio.VisRuleTargets.visRuleTargetDocument Then
                ReturnValue = Issue
                Exit For
            End If
        Next
        ' set the function as the value, a kind of return
        GetIssue = ReturnValue
    End Function

    Public Sub AddIssue(ByVal ruleName As String, ByVal ruleSet As Visio.ValidationRuleSet, ByVal page As Visio.Page, Optional ByVal shape As Visio.Shape = Nothing)
        Dim Rule As Visio.ValidationRule
        Dim Issue As Visio.ValidationIssue

        Rule = GetRule(ruleSet, ruleName)
        If Rule Is Nothing Then
            MsgBox("Could not find rule")
            Exit Sub
        End If
        Issue = GetIssue(Rule, page, shape)
        If Issue IsNot Nothing Then
            Issue.Delete()
        End If
        Issue = Rule.AddIssue(page, shape)
    End Sub

    Private Sub AddRule(ByRef ruleSet As Visio.ValidationRuleSet, _
                        ByVal ruleName As String, ByVal ruleCategory As String, _
                        ByVal ruleDescription As String, _
                        ByVal ruleTargetType As Visio.VisRuleTargets, _
                        Optional ByVal ruleFilterExpression As String = "", _
                        Optional ByVal ruleTestExpression As String = "")

        Dim Rule As Visio.ValidationRule

        Rule = GetRule(ruleSet, ruleName)
        If Rule Is Nothing Then
            Rule = ruleSet.Rules.Add(ruleName)
        End If
        Rule.Category = ruleCategory
        Rule.Description = ruleDescription
        Rule.TargetType = ruleTargetType
        Rule.FilterExpression = ruleFilterExpression
        Rule.TestExpression = ruleTestExpression

    End Sub

End Module