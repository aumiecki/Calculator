Option Explicit On
Option Strict On
Imports System.Math

Public Class Form1
    'declar global variables
    Dim dblResult As Double             'result of the calculation
    Dim dblCurrentNumber As Double      'stores current number in display
    Dim dblMemory As Double             'stores value placed in memory
    Dim blnStartNewNumber As Boolean      'determins if the new number should be displayed
    Dim strLastMathOperator As String   'stores selected operation

    ''' <summary>
    ''' clear the values of all variables
    ''' required to reset the calculator 
    ''' </summary>
    Private Sub resetCalculator()
        Me.lblResult.Text = "0"
        dblResult = 0
        strLastMathOperator = "Clear"
        blnStartNewNumber = True
        dblCurrentNumber = 0
        btnDecimal.Enabled = True
    End Sub

    ''' <summary>
    ''' builds number in the Result label
    ''' </summary>
    ''' <param name="strNumber">builds number in the display label</param>
    Private Sub buildNumber(ByVal strNumber As String)
        'should we start a new number or add to an existing Number?
        If blnStartNewNumber Then
            'start a new number
            lblResult.Text = strNumber
        Else
            'append to the current number
            lblResult.Text = lblResult.Text & strNumber
        End If
        dblCurrentNumber = Convert.ToDouble(lblResult.Text)
        blnStartNewNumber = False
    End Sub

    ''' <summary>
    ''' Applies the last operator to result using current number
    ''' </summary>
    ''' <param name="strOperation">the math operation to  
    '''     perform +,‐,*,/ or clear
    '''</param>
    Private Sub handleOperator(ByVal strOperation As String)
        Select Case strLastMathOperator.ToUpper
            Case "ADD"
                dblResult = dblResult + dblCurrentNumber
            Case "SUBTRACT"
                dblResult = dblResult - dblCurrentNumber
            Case "MULTIPLE"
                dblResult = dblResult * dblCurrentNumber
            Case "DIVIDE"
                dblResult = dblResult / dblCurrentNumber
            Case Else
                dblResult = dblCurrentNumber 'Equals
        End Select

        strLastMathOperator = strOperation
        dblCurrentNumber = dblResult
        Me.lblResult.Text = dblCurrentNumber.ToString
        Me.blnStartNewNumber = True
        btnDecimal.Enabled = True
    End Sub

    ''' <summary>
    ''' loads form
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">loads form</param>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'clear all variables on initial load
        resetCalculator()
    End Sub

    ''' <summary>
    ''' clear button to clear all numbers from the Result label
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">clear button to clear all numbers from the Result label</param>
    Private Sub btnPower_Click(sender As Object, e As EventArgs) Handles btnPower.Click
        'clear all variables
        resetCalculator()
    End Sub

    ''' <summary>
    ''' enter 0
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 0</param>
    Private Sub btn0_Click(sender As Object, e As EventArgs) Handles btn0.Click
        buildNumber("0")
    End Sub

    ''' <summary>
    ''' enter 1
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 1</param>
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        buildNumber("1")
    End Sub

    ''' <summary>
    ''' enter 2
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 2</param>
    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        buildNumber("2")
    End Sub

    ''' <summary>
    ''' enter 3
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 3</param>
    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        buildNumber("3")
    End Sub

    ''' <summary>
    ''' enter 4
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 4</param>
    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        buildNumber("4")
    End Sub

    ''' <summary>
    ''' enter 5
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 5</param>
    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        buildNumber("5")
    End Sub

    ''' <summary>
    ''' enter 6
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 6</param>
    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        buildNumber("6")
    End Sub

    ''' <summary>
    ''' enter 7
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 7</param>
    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        buildNumber("7")
    End Sub

    ''' <summary>
    ''' enter 8
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 8</param>
    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        buildNumber("8")
    End Sub

    ''' <summary>
    ''' enter 9
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">enter 9</param>
    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        buildNumber("9")
    End Sub

    ''' <summary>
    ''' execute the square MATH function
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">execute the square MATH function</param>
    Private Sub btnSQRT_Click(sender As Object, e As EventArgs) Handles btnSQRT.Click
        dblCurrentNumber = Sqrt(dblCurrentNumber)
        Me.lblResult.Text = Str(dblCurrentNumber)
    End Sub

    ''' <summary>
    ''' calculate percentake
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">calculate percentake</param>
    Private Sub btnPercent_Click(sender As Object, e As EventArgs) Handles btnPercent.Click
        Me.lblResult.Text = Str(dblCurrentNumber / 100)
        dblCurrentNumber = dblCurrentNumber / 100
        blnStartNewNumber = True
    End Sub

    ''' <summary>
    ''' plus/minus operator
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">plus/minus operator</param>
    Private Sub btnPlusMinus_Click(sender As Object, e As EventArgs) Handles btnPlusMinus.Click
        dblCurrentNumber = dblCurrentNumber * (-1)
        Me.lblResult.Text = dblCurrentNumber.ToString()
    End Sub

    ''' <summary>
    ''' add operator
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">add operator</param>
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        handleOperator("Add")
    End Sub

    ''' <summary>
    ''' subtract operator
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">subtract operator</param>
    Private Sub btnSubtract_Click(sender As Object, e As EventArgs) Handles btnSubtract.Click
        handleOperator("Subtract")
    End Sub

    ''' <summary>
    ''' multiple operator
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">multiple operator</param>
    Private Sub btnMultiple_Click(sender As Object, e As EventArgs) Handles btnMultiple.Click
        handleOperator("Multiple")
    End Sub

    ''' <summary>
    ''' divide operator
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">divide operator</param>
    Private Sub btnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        handleOperator("Divide")
    End Sub

    ''' <summary>
    ''' equals button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">equals button</param>
    Private Sub btnEqual_Click(sender As Object, e As EventArgs) Handles btnEqual.Click
        handleOperator("Equal")
    End Sub

    ''' <summary>
    ''' decimal button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">decimal button</param>
    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If blnStartNewNumber Then
            buildNumber("0.")
        Else
            buildNumber(".")
            btnDecimal.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' add to memory button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">add to memory button</param>
    Private Sub btnMemoryAdd_Click(sender As Object, e As EventArgs) Handles btnMemoryAdd.Click
        dblMemory = dblMemory + dblCurrentNumber
        btnMemoryRecall.BackColor = Color.Red 'change color to Red when button is pressed
    End Sub

    ''' <summary>
    ''' subtract from memory button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">subtract from memory button</param>
    Private Sub btnMemorySubtract_Click(sender As Object, e As EventArgs) Handles btnMemorySubtract.Click
        dblMemory = dblMemory - dblCurrentNumber
        btnMemoryRecall.BackColor = Color.Red 'change color to Red when button is pressed
    End Sub

    ''' <summary>
    ''' clear memory button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">clear memory button</param>
    Private Sub btnMemoryClear_Click(sender As Object, e As EventArgs) Handles btnMemoryClear.Click
        dblMemory = 0
        btnMemoryRecall.BackColor = Color.SteelBlue 'reset color to SteelBlue when button is pressed
    End Sub

    ''' <summary>
    ''' recall number from memory button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e">recall number from memory button</param>
    Private Sub btnMemoryRecall_Click(sender As Object, e As EventArgs) Handles btnMemoryRecall.Click
        lblResult.Text = dblMemory.ToString
        dblCurrentNumber = dblMemory
    End Sub
End Class
