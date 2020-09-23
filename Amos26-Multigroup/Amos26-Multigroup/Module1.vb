Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    Public Function Name() As String Implements IPlugin.Name
        Return "Multigroup"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "A plugin to automate the multigroup analysis testing for invariance."
    End Function

    Structure Estimates
        'An estimates struct holds the estimates
        Public Cmin As Double
        Public Df As Double
    End Structure

    Public Function Mainsub() As Integer Implements IPlugin.MainSub

        'Set up model
        pd.GetCheckBox("AnalysisPropertiesForm", "StandardizedCheck").Checked = True

        MsgBox("Clearing Paths")
        ClearPaths()

        'Get the cmin and df for the unconstrained model.
        MsgBox("Unconstrained Test")
        Dim unconstrainedEstimates As Estimates = GetEstimates()

        'Constrain all the relevant paths
        MsgBox("Naming Paths")
        NamePaths()

        'Get the cmin and df for the constrained model.
        MsgBox("Constrained Test")
        Dim constrainedEstimates As Estimates = GetEstimates()

        ClearPaths()
        MsgBox("Cleared Path")

        'The html output table after running local tests.
        CreateOutput(unconstrainedEstimates, constrainedEstimates, LocalTests(unconstrainedEstimates))

    End Function

    'Creates the output for the global and local tests.
    Public Sub CreateOutput(unconstrainedEstimates As Estimates, constrainedEstimates As Estimates, estimatesMatrix As Double(,))

        pd.AnalyzeCalculateEstimates() 'DO NOT REMOVE, resets the estimates table to the unconstrained output.

        If (System.IO.File.Exists("Multigroup.html")) Then 'Check if output file already exists.
            System.IO.File.Delete("Multigroup.html")
        End If

        'Set up the listener to output the debugger.
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("Multigroup.html")
        Trace.Listeners.Add(resultWriter)

        'Chi-squared difference for global test.
        Dim chiDifference As Double = AmosEngine.ChiSquareProbability(Math.Abs(unconstrainedEstimates.Cmin - constrainedEstimates.Cmin), Math.Abs(unconstrainedEstimates.Df - constrainedEstimates.Df))
        Dim names() As String = GroupNames() 'Names of the groups in the model
        Dim significance As String = "" 'Temp variable for significance tests.
        Dim tablerows As New Dictionary(Of String, Boolean)

        'Regression tables of both groups in the model.
        Dim tableRegression1 As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")
        Dim tableRegression2 As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 2]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")
        Dim numRegression As Integer = GetNodeCount(tableRegression1) 'Number of rows in the regression weights table.

        'Html DOC
        debug.PrintX("<html><body><h1>Multigroup Tests</h1><hr/>")

        'Global test
        debug.PrintX("<h2>Global Test</h2>")
        debug.PrintX("<table><tr><td></td><th>X<sup>2</sup></th><th>DF</th></tr>")
        debug.PrintX("<tr><td>Unconstrained</td><td align=center>" + unconstrainedEstimates.Cmin.ToString("#0.000") + "</td><td align=center>" + unconstrainedEstimates.Df.ToString + "</td></tr>")
        debug.PrintX("<tr><td>Constrained</td><td align=center>" + constrainedEstimates.Cmin.ToString("#0.000") + "</td><td align=center>" + constrainedEstimates.Df.ToString + "</td></tr>")
        debug.PrintX("<tr><td>Difference</td><td align=center>" + Math.Abs((constrainedEstimates.Cmin - unconstrainedEstimates.Cmin)).ToString("#0.000") + "</td><td align=center>" + Math.Abs((constrainedEstimates.Df - unconstrainedEstimates.Df)).ToString + "</td></tr>")
        debug.PrintX("<tr><td>P-Value</td><td colspan=2 align=center>" + chiDifference.ToString("#0.000") + "</td></tr>")
        debug.PrintX("</table>")

        'Global test interpretation
        If chiDifference < 0.1 Then
            debug.PrintX("<p><strong>Interpretation:</strong> The p-value of the chi-square difference test is significant; the model differs across groups.</p><hr/>")
        Else
            debug.PrintX("<p><strong>Interpretation:</strong> The p-value of the chi-square difference test is not significant; interpret local tests with caution.</p><hr/>")
        End If

        'Local tests
        debug.PrintX("<h2>Local Tests</h2>")
        debug.PrintX("<table><tr><th>Path Name</th><th>" + names(0) + " Beta</th><th>" + names(1) + " Beta</th><th>Difference in Betas</th><th>P-Value for Difference</th><th>Interpretation</th></tr>")
        For x = 1 To numRegression 'Iterate through regression weights tables.
            Dim check1 = False
            Dim check2 = False
            Dim check3 = False
            If estimatesMatrix(0, x) <> 0 Then
                Dim id As String = "id=""row" + CStr(x) + """"
                debug.PrintX("<tr " + id + "><td>" + MatrixName(tableRegression1, x, 2) + " &#8594 " + MatrixName(tableRegression1, x, 0) + ".") 'Name of path
                For y = 0 To 3
                    If y = 0 Then 'Check significance of beta.
                        If MatrixName(tableRegression1, x, 6) = "***" Then
                            significance = "***"
                        ElseIf MatrixElement(tableRegression1, x, 6) < 0.01 Then
                            significance = "**"
                        ElseIf MatrixElement(tableRegression1, x, 6) < 0.05 Then
                            significance = "*"
                        ElseIf MatrixElement(tableRegression1, x, 6) < 0.1 Then
                            significance = "&#8224;"
                        Else
                            significance = ""
                            check1 = True
                        End If

                        'Group 1 beta
                        debug.PrintX("<td align=center>" + estimatesMatrix(y, x).ToString("#0.000") + significance + "</td>")

                    ElseIf y = 1 Then 'Check significance of beta.
                        If MatrixName(tableRegression2, x, 6) = "***" Then
                            significance = "***"
                        ElseIf MatrixElement(tableRegression2, x, 6) < 0.01 Then
                            significance = "**"
                        ElseIf MatrixElement(tableRegression2, x, 6) < 0.05 Then
                            significance = "*"
                        ElseIf MatrixElement(tableRegression2, x, 6) < 0.1 Then
                            significance = "&#8224;"
                        Else
                            significance = ""
                            check2 = True
                        End If

                        'Group 2 beta
                        debug.PrintX("<td align=center>" + estimatesMatrix(y, x).ToString("#0.000") + significance + "</td>")
                    Else
                        'Difference in betas and chi-squared difference test.
                        debug.PrintX("<td align=center>" + estimatesMatrix(y, x).ToString("#0.000") + "</td>")
                    End If
                Next

                'Local tests interpretations
                debug.PrintX("<td>")

                If (MatrixName(tableRegression1, x, 6) = "***" Or MatrixElement(tableRegression1, x, 6) < 0.1) And (MatrixName(tableRegression2, x, 6) = "***" Or MatrixElement(tableRegression2, x, 6) < 0.1) Then 'Both betas significant
                    If estimatesMatrix(0, x) > 0 And estimatesMatrix(1, x) > 0 Then 'Both betas positive
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            If estimatesMatrix(0, x) > estimatesMatrix(1, x) Then 'First beta is greater
                                debug.PrintX("The positive relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(0) + ".")
                            ElseIf estimatesMatrix(0, x) < estimatesMatrix(1, x) Then 'Second beta is greater
                                debug.PrintX("The positive relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(1) + ".")
                            Else
                                debug.PrintX("Well, this was unexpected... what do you think is happening here?</td>")
                            End If
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("There is no difference.")
                            check3 = True
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    ElseIf estimatesMatrix(0, x) < 0 And estimatesMatrix(1, x) < 0 Then 'Both betas negative
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            If estimatesMatrix(0, x) < estimatesMatrix(1, x) Then 'First beta is lower
                                debug.PrintX("The negative relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(0) + ".")
                            ElseIf estimatesMatrix(0, x) > estimatesMatrix(1, x) Then 'Second beta is lower
                                debug.PrintX("The negative relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(1) + ".")
                            Else
                                debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                            End If
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("There is no difference.")
                            check3 = True
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    ElseIf estimatesMatrix(0, x) > 0 And estimatesMatrix(1, x) < 0 Then 'First beta is positive
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            debug.PrintX("The relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is positive for " + names(0) + " and negative for " + names(1) + ".")
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("The relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is positive for " + names(0) + " and negative for " + names(1) + ", but there is no statistically significant difference.")
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    ElseIf estimatesMatrix(0, x) < 0 And estimatesMatrix(1, x) > 0 Then 'First beta is negative
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            debug.PrintX("The relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is negative for " + names(0) + " and positive for " + names(1) + ".")
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("The relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is negative for " + names(0) + " and positive for " + names(1) + ", but there is no statistically significant difference.")
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    Else
                        debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                    End If
                ElseIf (MatrixName(tableRegression1, x, 6) = "***" Or MatrixElement(tableRegression1, x, 6) < 0.1) And (MatrixName(tableRegression2, x, 6) <> "***" Or MatrixElement(tableRegression2, x, 6) > 0.1) Then 'First significant
                    If estimatesMatrix(0, x) > 0 Then 'First beta is positive
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            debug.PrintX("The positive relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(0) + ".")
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("The positive relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is only significant for " + names(0) + ".")
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    ElseIf estimatesMatrix(0, x) < 0 Then 'First beta is negative
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            debug.PrintX("The negative relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(0) + ".")
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("The negative relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is only significant for " + names(0) + ".")
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    Else
                        debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                    End If
                ElseIf (MatrixName(tableRegression1, x, 6) <> "***" Or MatrixElement(tableRegression1, x, 6) > 0.1) And (MatrixName(tableRegression2, x, 6) = "***" Or MatrixElement(tableRegression2, x, 6) < 0.1) Then 'Second significant
                    If estimatesMatrix(1, x) > 0 Then 'Second beta is positive
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            debug.PrintX("The positive relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(1) + ".")
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("The positive relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is only significant for " + names(1) + ".")
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    ElseIf estimatesMatrix(1, x) < 0 Then 'Second beta is negative
                        If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                            debug.PrintX("The negative relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is stronger for " + names(1) + ".")
                        ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                            debug.PrintX("The negative relationship between " + MatrixName(tableRegression1, x, 0) + " and " + MatrixName(tableRegression1, x, 2) + " is only significant for " + names(1) + ".")
                        Else
                            debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                        End If
                    Else
                        debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                    End If
                ElseIf (MatrixName(tableRegression1, x, 6) <> "***" Or MatrixElement(tableRegression1, x, 6) > 0.1) And (MatrixName(tableRegression2, x, 6) <> "***" Or MatrixElement(tableRegression2, x, 6) > 0.1) Then 'Both not significant
                    If estimatesMatrix(3, x) < 0.1 Then 'Difference in betas is significant
                        debug.PrintX("While there is a difference, that difference cannot be identified because the relationship is essentially not different from zero for both groups.")
                    ElseIf estimatesMatrix(3, x) > 0.1 Then 'Difference in betas is not significant
                        debug.PrintX("There is no difference")
                        check3 = True
                    Else
                        debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                    End If
                Else
                    debug.PrintX("Well, this was unexpected... what do you think is happening here?")
                End If

                debug.PrintX("</td></tr>")
            End If
            If check1 = True And check2 = True And check3 = True Then
                tablerows.Add("row" + CStr(x), True)
            Else
                tablerows.Add("row" + CStr(x), False)
            End If

        Next
        debug.PrintX("</table><hr/>")

        'References
        debug.PrintX("<strong>Significance Indicators:</strong><br>&#8224; p < 0.100<br>* p < 0.050<br>** p < 0.010<br>*** p < 0.001<br>")
        debug.PrintX("<p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2018), ""Multigroup Analysis"", AMOS Plugin. <a href=""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}")
        For Each row In tablerows
            If row.Value = True Then
                debug.PrintX("#" + row.Key + "{background-color:#B0B0B0;}")
            End If
        Next
        debug.PrintX("</style>")
        debug.PrintX("</body></html>")

        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("Multigroup.html")
    End Sub

    'Analyze estimates for every path constrained in the model.
    Public Function LocalTests(unconstrainedEstimates As Estimates) As Double(,)

        pd.AnalyzeCalculateEstimates() 'DO NOT REMOVE, resets the estimates table to the unconstrained output.

        'Two different estimates tables for two different groups
        Dim tableGroup1 As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")
        Dim tableGroup2 As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 2]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")

        Dim numRegression As Integer = GetNodeCount(tableGroup1) 'Number of rows in the regression tables.
        Dim estimatesMatrix(3, numRegression) As Double 'Matrix to hold the betas and significance of difference.

        For x = 1 To numRegression 'Iterate through the regression table.
            For Each variable As PDElement In pd.PDElements
                If variable.IsPath Then 'Iterate through the paths in the model.
                    If (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Or (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Then
                        If variable.Variable1.NameOrCaption = MatrixName(tableGroup1, x, 2) And variable.Variable2.NameOrCaption = MatrixName(tableGroup1, x, 0) Then 'Check if name of connected variables matches regression table.
                            variable.Value1 = variable.Variable1.NameOrCaption + "_" + variable.Variable2.NameOrCaption 'Constrain model with path name
                            Dim localEstimates As Estimates = GetEstimates()
                            estimatesMatrix(0, x) = MatrixElement(tableGroup1, x, 3) 'Group 1 beta
                            estimatesMatrix(1, x) = MatrixElement(tableGroup2, x, 3) 'Group 2 beta
                            estimatesMatrix(2, x) = MatrixElement(tableGroup1, x, 3) - MatrixElement(tableGroup2, x, 3) 'Difference in betas
                            estimatesMatrix(3, x) = AmosEngine.ChiSquareProbability(Math.Abs(unconstrainedEstimates.Cmin - localEstimates.Cmin), Math.Abs(unconstrainedEstimates.Df - localEstimates.Df)) 'Chi-squared difference test.
                            ClearPaths() 'Rest names of all paths.
                        End If
                    End If
                End If
            Next
        Next

        LocalTests = estimatesMatrix

    End Function

    'Get the names of the two groups in the model.
    Public Function GroupNames() As String()


        'Dim Sem As New AmosEngineLib.AmosEngine
        Dim names(1) As String

        '     pd.SpecifyModel(Sem)
        '      Sem.FitModel()

        '        names(0) = Sem.GetGroupName(1)
        '       names(1) = Sem.GetGroupName(2)

        names(0) = "Male"
        names(1) = "Female"


        'Sem.Dispose()

        GroupNames = names

    End Function

    'Gets the cmin and the df for the given model condition.
    Function GetEstimates() As Estimates

        pd.AnalyzeCalculateEstimates()

        'Properties
        Dim baseEstimates As Estimates
        Dim modelNotes As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='modelnotes']/div[@ntype='result']")

        'Use regex to extract the chi-square and df from the result
        Dim result As String = modelNotes.InnerText
        Dim myRegex As Match = Regex.Match(result, "\d+(\.\d{1,3})?", RegexOptions.IgnoreCase)
        If myRegex.Success Then
            baseEstimates.Cmin = Convert.ToDouble(Convert.ToString(myRegex.Value))
            baseEstimates.Df = Convert.ToDouble(Convert.ToString(myRegex.NextMatch))
        Else
            MsgBox(modelNotes.InnerText)
            Exit Function
        End If

        GetEstimates = baseEstimates

    End Function

    'Get a string element from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As Exception
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As Exception
            MatrixElement = 0
        End Try

    End Function

    'Use an output table path to get the xml version of the table.
    Public Function GetXML(path As String) As XmlElement

        'Gets the xpath expression for an output table.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        Return eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get the number of rows in an xml table.
    Public Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        GetNodeCount = nodeCount

    End Function

    'Method to clear names of paths unless value is 1.
    Public Sub ClearPaths()
        'Set paths back to null.
        For Each variable As PDElement In pd.PDElements 'Iterate through the paths in the model
            If variable.IsPath Then
                'If (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Or (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Then
                If variable.Value1 <> "1" Then
                    variable.Value1 = "" 'Clear Path name
                End If    'End If
            End If
        Next

    End Sub

    'Constrain all the paths that are not equal to 1
    Public Sub NamePaths()

        'Just do observed to observed, latent to latent
        For Each variable As PDElement In pd.PDElements 'Iterate through the paths in the model
            If variable.IsPath Then
                If (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Or (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Then
                    variable.Value1 = variable.Variable1.NameOrCaption + "_" + variable.Variable2.NameOrCaption 'Constrain model with path name
                End If
            End If
        Next


    End Sub


End Class
