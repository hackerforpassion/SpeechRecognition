Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Diagnostics
Imports System
Imports System.Windows.Forms
Imports System.Collections
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports System.Data
'using AxIEPOPObjects;
Imports System.ComponentModel
Imports SpeechLib

Public Class FormCalculator
#Region "User Declarations"
    Dim c As gscore10.mnulh2
    Dim my As New gscore10.audio()
    Dim down As Boolean
    Dim t As Integer
    Dim w As Integer
    Dim iscurrfile As Boolean
#End Region
#Region "My Procedures"
    Private Sub aplay()
        my.audioFile = "C:\slumdog-3.wav"
        ' Label1.Text = "C:\Users\Public\Music\Sample Music\slumdog-1.wav"
        iscurrfile = True
        On Error GoTo err
        'Me.Text = "Musicplayer-" & my.audioFile
        my.audioPlay()

err:
        Exit Sub
    End Sub
    Private Sub astop()
        On Error GoTo err
        Me.Text = "Voice Calculator"
        my.audioStop()
err:
        Exit Sub
    End Sub
#End Region
    Dim voip As New HTTSLib.TextToSpeechClass

    <DllImport("winmm.dll")> _
   Public Shared Function sndPlaySound(ByVal lpszSound As String, _
    ByVal flags As UInt32) As Boolean
    End Function
    Private objRecoContext As SpeechLib.SpSharedRecoContext = Nothing
    Private grammar As SpeechLib.ISpeechRecoGrammar = Nothing
    Private menuRule As SpeechLib.ISpeechGrammarRule = Nothing



    Dim sign_Indicator As Integer = 0
    Dim variable1 As Double
    Dim i As Integer = 0
    Dim muc As Integer
    Dim time As Integer = 0
    Dim variable2 As Double
    Dim addBit As Integer = 0
    Dim subBit As Integer = 0
    Dim multBit As Integer = 0
    Dim divBit As Integer = 0
    Dim modBit As Integer = 0
    Dim powerBit As Integer = 0
    Dim permBit As Integer = 0
    Dim combBit As Integer = 0
    Dim andBit As Integer = 0
    Dim orBit As Integer = 0
    Dim xorBit As Integer = 0
    Dim powerFunctionBit As Integer = 0
    Dim trigFunctionBit As Integer = 0
    Dim HypertrigFunctionBit As Integer = 0
    Dim InversetrigFunctionBit As Integer = 0
    Dim otherFuncsBit As Integer = 0
    Dim logicalFuncsBit As Integer = 0
    Dim memFuncsBit As Integer = 0
    Dim fl As Integer = 0
    Dim memoryVariable As Double = 0
    Dim scientificModeBit As Integer = 0
    Dim tim As Integer = 0
    Dim results As Integer
    Dim ind As Integer = 1
    Dim valc As Double = 0
    Dim sci As Integer = 0
    Dim mem As Integer = 0
    Dim pow As Integer = 0
    Dim logic As Integer = 0
    Dim str As Integer = 0
    Dim htr As Integer = 0
    Dim itr As Integer = 0
    Dim oth As Integer = 0
    Dim clor As Integer = 0
    Dim voc As Integer = 0
    Dim msc As Integer = 0

    Private Sub btn0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn0.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(0)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 0
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("zero")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(1)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 1
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("one")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(2)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 2
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("two")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn3.Click
        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(3)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 3
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("three")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn4.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(4)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 4
            sign_Indicator = 0
        End If
        fl = 1

        If ind = 1 Then
            Try
                With voip
                    .Speak("four")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn5.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(5)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 5
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("five")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn6.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(6)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 6
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("six")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn7.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(7)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 7
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("Seven")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn8.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(8)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 8
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("eight")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btn9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn9.Click

        If sign_Indicator = 0 Then
            displayText.Text = displayText.Text & CStr(9)
        ElseIf sign_Indicator = 1 Then
            displayText.Text = 9
            sign_Indicator = 0
        End If
        fl = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("nine")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub displayText_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles displayText.KeyPress
        e.Handled = True
        If e.KeyChar Like "[0-9]" Or e.KeyChar = Chr(&H8) Then
            e.Handled = False
        End If
    End Sub

    Private Sub btnSign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSign.Click

        If displayText.Text.Length = 0 Then
            displayText.Text = displayText.Text + CStr("-")
        ElseIf displayText.Text <> "." Then
            displayText.Text = displayText.Text * -1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("Sign changed")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click

        displayText.Clear()
        sign_Indicator = 0
        variable1 = 0
        variable2 = 0
        memoryVariable = 0
        reset_SignBits()
        If ind = 1 Then
            Try
                With voip
                    .Speak("clear all")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnClearTextbox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearTextbox.Click

        displayText.Clear()
        If ind = 1 Then
            Try
                With voip
                    .Speak("clear")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Function reset_SignBits()
        addBit = 0
        subBit = 0
        multBit = 0
        divBit = 0
        modBit = 0
        powerBit = 0
        permBit = 0
        combBit = 0
        andBit = 0
        orBit = 0
        xorBit = 0
        fl = 0
        Return 0
    End Function

    Private Function permutation(ByVal variable1, ByVal variable2)
        Dim result As Double
        result = Factorial(variable1) / Factorial(variable1 - variable2)
        Return result
    End Function

    Private Function combination(ByVal variable1, ByVal variable2)
        Dim result As Double
        result = Factorial(variable1) / (Factorial(variable2) * Factorial(variable1 - variable2))
        Return result
    End Function

    Private Function Calculate()
        If displayText.Text <> "." Then

            variable2 = displayText.Text
            If fl = False Then
                variable1 = variable2
            ElseIf addBit = 1 Then
                variable1 = variable1 + variable2
            ElseIf subBit = 1 Then
                variable1 = variable1 - variable2
            ElseIf multBit = 1 Then
                variable1 = variable1 * variable2
            ElseIf divBit = 1 Then
                variable1 = variable1 / variable2
            ElseIf modBit = 1 Then
                variable1 = variable1 Mod variable2
            ElseIf powerBit = 1 Then
                variable1 = Math.Pow(variable1, variable2)
            ElseIf permBit = 1 Then
                variable1 = permutation(variable1, variable2)
            ElseIf combBit = 1 Then
                variable1 = combination(variable1, variable2)
            ElseIf andBit = 1 Then
                variable1 = variable1 And variable2
            ElseIf orBit = 1 Then
                variable1 = variable1 Or variable2
            ElseIf xorBit = 1 Then
                variable1 = variable1 Xor variable2
            Else
                variable1 = variable2
            End If
            displayText.Text = CStr(variable1)

        End If
        Return 0
    End Function

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            addBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("plus")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnSub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSub.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("minus")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            variable2 = displayText.Text
            Calculate()
            reset_SignBits()
            subBit = 1
            sign_Indicator = 1
        End If

    End Sub

    Private Sub btnMult_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMult.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            multBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("multiply")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnDiv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDiv.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            divBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("Divide")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnModulus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModulus.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            modBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("percentage")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnPower_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPower.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            powerBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("power n")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnPerm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPerm.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            permBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("pnr")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnComb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComb.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            combBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("ncr")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnAnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnd.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            andBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("logical And")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnOr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOr.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            orBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("logical or")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnNot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNot.Click

        If displayText.Text.Length <> 0 Then
            variable1 = CDbl(displayText.Text)
            variable1 = Not variable1
            displayText.Text = CStr(variable1)
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("logical not")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Try
                With voip
                    .Speak(displayText.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnXor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnXor.Click

        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
            xorBit = 1
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("logical x or")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnEqual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqual.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("equals")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            Calculate()
            reset_SignBits()
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak(displayText.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnDecimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDecimal.Click
        If displayText.Text = "" Then
            If ind = 1 Then
                Try
                    With voip
                        .Speak(lblInfoDecimal.Text)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            lblInfoDecimal.Visible = True
            Timer1.Enabled = True
        End If
        Dim i As Integer
        Dim chr As Char
        Dim decimal_Indicator As Integer = 0

        If sign_Indicator <> 1 Then
            For i = 0 To displayText.TextLength - 1
                chr = displayText.Text(i)
                If chr = "." Then
                    decimal_Indicator = 1
                End If
            Next

            If decimal_Indicator <> 1 Then
                displayText.Text = displayText.Text & CStr(".")
            End If
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("dot")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text = "" Then
            lblInfoDecimal.Visible = True
        End If
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        lblTime.Text = TimeString
    End Sub
    Private Sub btna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btna.Click

        BackColor = Color.DimGray
        btn0.BackColor = Color.Chocolate
        btn1.BackColor = Color.Chocolate
        btn2.BackColor = Color.Chocolate
        btn3.BackColor = Color.Chocolate
        btn4.BackColor = Color.Chocolate
        btn5.BackColor = Color.Chocolate
        btn6.BackColor = Color.Chocolate
        btn7.BackColor = Color.Chocolate
        btn8.BackColor = Color.Chocolate
        btn9.BackColor = Color.Chocolate
        If ind = 1 Then
            Try
                With voip
                    .Speak("Chocolate")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnb.Click

        BackColor = Color.SlateGray
        btn0.BackColor = Color.DimGray
        btn1.BackColor = Color.DimGray
        btn2.BackColor = Color.DimGray
        btn3.BackColor = Color.DimGray
        btn4.BackColor = Color.DimGray
        btn5.BackColor = Color.DimGray
        btn6.BackColor = Color.DimGray
        btn7.BackColor = Color.DimGray
        btn8.BackColor = Color.DimGray
        btn9.BackColor = Color.DimGray
        If ind = 1 Then
            Try
                With voip
                    .Speak("violet")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnc.Click
        BackColor = Color.Black
        btn0.BackColor = Color.DarkRed
        btn1.BackColor = Color.DarkRed
        btn2.BackColor = Color.DarkRed
        btn3.BackColor = Color.DarkRed
        btn4.BackColor = Color.DarkRed
        btn5.BackColor = Color.DarkRed
        btn6.BackColor = Color.DarkRed
        btn7.BackColor = Color.DarkRed
        btn8.BackColor = Color.DarkRed
        btn9.BackColor = Color.DarkRed
        If ind = 1 Then
            Try
                With voip
                    .Speak("Red")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnBackSpace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackSpace.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("Backspace")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = displayText.Text.Remove(displayText.Text.Length - 1, 1)
        End If
    End Sub

    Private Sub btnInverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInverse.Click

        If displayText.Text.Length <> 0 Then
            displayText.Text = 1 / displayText.Text
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("one by x")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnSin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSin.Click

        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Sin(displayText.Text)
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("sign")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnCos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCos.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("cos")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Cos(displayText.Text)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnTan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTan.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("tangent")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Tan(displayText.Text)
        End If
        sign_Indicator = 1
    End Sub

    Private Sub btnPI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPI.Click

        displayText.Text = Math.PI
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("pai")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnSinh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSinh.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("sign h")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Sinh(displayText.Text)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnCosh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCosh.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("cos h")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Cosh(displayText.Text)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnTanh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTanh.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("tan h")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Tanh(displayText.Text)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnlog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnlog.Click

        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Log10(displayText.Text)
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("logarithm")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnln_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnln.Click

        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Log(displayText.Text)
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("natural logarithm")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btne.Click

        displayText.Text = Math.E
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("Expression")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnex.Click

        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Exp(displayText.Text)
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("exponent")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnsqrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsqrt.Click

        If displayText.Text.Length <> 0 Then
            If displayText.Text <> "." Then
                displayText.Text = Math.Sqrt(displayText.Text)
            End If
            sign_Indicator = 1
        End If
        If ind = 1 Then
            Try
                With voip
                    .Speak("Square root")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnRound_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRound.Click

        If ind = 1 Then

            Try
                With voip
                    .Speak("Round")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            Dim temp As Double = displayText.Text
            displayText.Text = Math.Round(temp)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnTruncate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTruncate.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("Truncate")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            Dim temp As Double = displayText.Text
            displayText.Text = Math.Truncate(temp)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnCeil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCeil.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("ceil")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            Dim temp As Double = displayText.Text
            displayText.Text = Math.Ceiling(temp)
        End If
        sign_Indicator = 1

    End Sub

    Private Sub btnFloor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFloor.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("Floor")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            Dim temp As Double = displayText.Text
            displayText.Text = Math.Floor(temp)
        End If
        sign_Indicator = 1

    End Sub

    Public Function Factorial(ByVal temp)
        Dim result As Double = temp
        If temp = 0 Then
            Return 1
        Else
            For i = temp - 1 To 1 Step -1
                result = result * i
            Next
        End If
        Return result
    End Function

    Private Sub btnFact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFact.Click

        If displayText.Text.Length <> 0 Then
            Dim temp As Double = displayText.Text
            Dim result = Factorial(temp)
            displayText.Text = CStr(result)
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("factorial")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub


    Private Sub btnSquare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSquare.Click
        If displayText.Text.Length <> 0 Then
            displayText.Text = displayText.Text * displayText.Text
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("square")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Try
                With voip
                    .Speak(displayText.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnCube_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCube.Click
        If displayText.Text.Length <> 0 Then
            displayText.Text = displayText.Text * displayText.Text * displayText.Text
        End If
        sign_Indicator = 1
        If ind = 1 Then
            Try
                With voip
                    .Speak("cube")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Try
                With voip
                    .Speak(displayText.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub



    Private Sub btnScientificMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScientificMode.Click
        If scientificModeBit = 0 Then
            If ind = 1 Then
                Try
                    With voip
                        .Speak("Scientific mode enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            btnFact.Enabled = 1
            btnPerm.Enabled = 1
            btnComb.Enabled = 1
            btnPI.Enabled = 1
            btnsqrt.Enabled = 1
            btnlog.Enabled = 1
            btnex.Enabled = 1
            btnln.Enabled = 1
            btnInverse.Enabled = 1
            btnModulus.Enabled = 1
            btne.Enabled = 1
            btnPower.Enabled = 1
            btnSquare.Enabled = 1
            btnCube.Enabled = 1
            btnAnd.Enabled = 1
            btnOr.Enabled = 1
            btnNot.Enabled = 1
            btnXor.Enabled = 1
            btnSin.Enabled = 1
            btnCos.Enabled = 1
            btnTan.Enabled = 1
            btnSinh.Enabled = 1
            btnCosh.Enabled = 1
            btnTanh.Enabled = 1
            btnInvSin.Enabled = 1
            btnInvCos.Enabled = 1
            btnInvTan.Enabled = 1
            btnRound.Enabled = 1
            btnFloor.Enabled = 1
            btnTruncate.Enabled = 1
            btnCeil.Enabled = 1
            btnM.Enabled = 1


            scientificModeBit = 1
        ElseIf scientificModeBit = 1 Then
            btnFact.Enabled = 0
            btnPerm.Enabled = 0
            btnComb.Enabled = 0
            btnPI.Enabled = 0
            btnsqrt.Enabled = 0
            btnlog.Enabled = 0
            btnex.Enabled = 0
            btnln.Enabled = 0
            btnInverse.Enabled = 0
            btnModulus.Enabled = 0
            btne.Enabled = 0
            btnPower.Enabled = 0
            btnSquare.Enabled = 0
            btnCube.Enabled = 0

            btnAnd.Enabled = 0
            btnOr.Enabled = 0
            btnNot.Enabled = 0
            btnXor.Enabled = 0
            btnSin.Enabled = 0
            btnCos.Enabled = 0
            btnTan.Enabled = 0
            btnSinh.Enabled = 0
            btnCosh.Enabled = 0
            btnTanh.Enabled = 0
            btnInvSin.Enabled = 0
            btnInvCos.Enabled = 0
            btnInvTan.Enabled = 0
            btnRound.Enabled = 0
            btnFloor.Enabled = 0
            btnTruncate.Enabled = 0
            btnCeil.Enabled = 0
            btnM.Enabled = 0

            scientificModeBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("Scientific mode disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnPowerFunctions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPowerFunctions.Click
        If powerFunctionBit = 0 Then
            btnPower.Enabled = 1
            btnSquare.Enabled = 1
            btnCube.Enabled = 1
            powerFunctionBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("Power functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf powerFunctionBit = 1 Then
            btnPower.Enabled = 0
            btnSquare.Enabled = 0
            btnCube.Enabled = 0

            powerFunctionBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("power functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnTrignometric_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrignometric.Click
        If trigFunctionBit = 0 Then
            btnSin.Enabled = 1
            btnCos.Enabled = 1
            btnTan.Enabled = 1
            trigFunctionBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("Simple Trignometric functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf trigFunctionBit = 1 Then
            btnSin.Enabled = 0
            btnCos.Enabled = 0
            btnTan.Enabled = 0
            trigFunctionBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("simple Trignometric functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnInverseTrigFuncs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInverseTrigFuncs.Click
        If InversetrigFunctionBit = 0 Then
            btnInvSin.Enabled = 1
            btnInvCos.Enabled = 1
            btnInvTan.Enabled = 1
            InversetrigFunctionBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("inverse trignometric functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf InversetrigFunctionBit = 1 Then
            btnInvSin.Enabled = 0
            btnInvCos.Enabled = 0
            btnInvTan.Enabled = 0
            InversetrigFunctionBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("inverse trignometric functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnHyperbolicTrig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHyperbolicTrig.Click
        If HypertrigFunctionBit = 0 Then
            btnSinh.Enabled = 1
            btnCosh.Enabled = 1
            btnTanh.Enabled = 1
            HypertrigFunctionBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("hyperbolic trignometric functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf HypertrigFunctionBit = 1 Then
            btnSinh.Enabled = 0
            btnCosh.Enabled = 0
            btnTanh.Enabled = 0
            HypertrigFunctionBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("hyperbolic trignometric functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnOtherFuncs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOtherFuncs.Click
        If otherFuncsBit = 0 Then
            btnRound.Enabled = 1
            btnFloor.Enabled = 1
            btnTruncate.Enabled = 1
            btnCeil.Enabled = 1
            otherFuncsBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("other functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf otherFuncsBit = 1 Then
            btnRound.Enabled = 0
            btnFloor.Enabled = 0
            btnTruncate.Enabled = 0
            btnCeil.Enabled = 0
            otherFuncsBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("other functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnLogicalFuncs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogicalFuncs.Click
        If logicalFuncsBit = 0 Then
            btnAnd.Enabled = 1
            btnOr.Enabled = 1
            btnNot.Enabled = 1
            btnXor.Enabled = 1
            logicalFuncsBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("logical functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf logicalFuncsBit = 1 Then
            btnAnd.Enabled = 0
            btnOr.Enabled = 0
            btnNot.Enabled = 0
            btnXor.Enabled = 0
            logicalFuncsBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("logical functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnMemFuncs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMemFuncs.Click
        If memFuncsBit = 0 Then
            btnM.Enabled = 1
            memFuncsBit = 1
            If ind = 1 Then
                Try
                    With voip
                        .Speak("Memory functions enabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        ElseIf memFuncsBit = 1 Then
            btnM.Enabled = 0

            memFuncsBit = 0
            If ind = 1 Then
                Try
                    With voip
                        .Speak("Memory functions disabled")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
        End If
    End Sub

    Private Sub btnInvSin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvSin.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("Secent")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Asin(displayText.Text)
        End If
        sign_Indicator = 1
    End Sub

    Private Sub btnInvCos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvCos.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("coSecent")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Acos(displayText.Text)
        End If
        sign_Indicator = 1
    End Sub

    Private Sub btnInvTan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvTan.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak("cot")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If displayText.Text.Length <> 0 Then
            displayText.Text = Math.Atan(displayText.Text)
        End If
        sign_Indicator = 1
    End Sub

    Private Sub btnM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnM.Click
        If displayText.Text.Length <> 0 Then
            If memoryVariable = 0 Then
                memoryVariable = CDbl(displayText.Text)
            End If
        End If
        sign_Indicator = 1
        displayText.Text = CStr(memoryVariable)
        If ind = 1 Then
            Try
                With voip
                    .Speak("memory")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub btnConstants_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConstants.Click
        Dim new_Form As New FormConstants
        If ind = 1 Then
            Try
                With voip
                    .Speak("Constants values")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        new_Form.Show()

    End Sub
    Private Sub frmDictionary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer5.Enabled = True
        muc = 0
        Timer4.Enabled = True
        CheckBox2.Enabled = False
        CheckBox2.Checked = False
        CheckBox1.Enabled = False

        Label3.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Timer2.Enabled = True
        Button6.Visible = False
        Button5.Visible = False
        tbVoiceSpeed.Visible = False
        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        lblInfoDecimal.Visible = False
        lblTime.Visible = 1
        gbMain1.Visible = 1
        btnTime.Visible = True
        btnColour.Visible = True
        btna.Visible = 0
        btnb.Visible = 0
        btnc.Visible = 0
        advc.Enabled = False
        With tbVoiceSpeed
            .SetRange(0, 10)
            .Value = 5
            lblTrackBarValue.Text = "5"
        End With

        'The default speed
        With voip
            .Speed = ((450 * 50) / 100)
        End With

        Try
            'Sound file that plays on starting of the application
            sndPlaySound("C:\WINNT\Media\Windows Logon Sound.wav", New UInt32)
            'With voip
            '.Speak("Welcome To voice calculator")
            'End With
        Catch ex As Exception
            MsgBox(ex.Message)

            'With voip
            '.Speak("Welcome To voice calculator")
            ' End With
        End Try
        ' With voip
        '.Speak("i'm listening")
        'End With
        label4.Text = "Initializing Speech Engine...."
        ' Get an insance of RecoContext. I am using the shared RecoContext.
        objRecoContext = New SpeechLib.SpSharedRecoContext()
        ' Assign a eventhandler for the Hypothesis Event.
        AddHandler objRecoContext.Hypothesis, AddressOf Hypo_Event
        ' Assign a eventhandler for the Recognition Event.
        AddHandler objRecoContext.Recognition, AddressOf Reco_Event
        'Creating an instance of the grammer object.
        grammar = objRecoContext.CreateGrammar(0)
        'Activate the Menu Commands.			
        menuRule = grammar.Rules.Add("MenuCommands", SpeechRuleAttributes.SRATopLevel Or SpeechRuleAttributes.SRADynamic, 1)
        Dim PropValue As Object = ""
        menuRule.InitialState.AddWordTransition(Nothing, "one", " ", SpeechGrammarWordType.SGLexical, "one", 1, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "two", "", SpeechGrammarWordType.SGLexical, "two", 2, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "three", "", SpeechGrammarWordType.SGLexical, "three", 3, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "four", "", SpeechGrammarWordType.SGLexical, "four", 4, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "five", "", SpeechGrammarWordType.SGLexical, "five", 5, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "six", "", SpeechGrammarWordType.SGLexical, "six", 6, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seven", "", SpeechGrammarWordType.SGLexical, "seven", 7, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eight", "", SpeechGrammarWordType.SGLexical, "eight", 8, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "nine", "", SpeechGrammarWordType.SGLexical, "nine", 9, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "zero", "", SpeechGrammarWordType.SGLexical, "zero", 0, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "plus", "", SpeechGrammarWordType.SGLexical, "plus", 10, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "minus", "", SpeechGrammarWordType.SGLexical, "minus", 11, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "multiply", "", SpeechGrammarWordType.SGLexical, "multiply", 12, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "divide", "", SpeechGrammarWordType.SGLexical, "divide", 13, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "dot", "", SpeechGrammarWordType.SGLexical, "dot", 14, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "equals", "", SpeechGrammarWordType.SGLexical, "equals", 15, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "voice", "", SpeechGrammarWordType.SGLexical, "voice", 16, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "speech", "", SpeechGrammarWordType.SGLexical, "speech", 17, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "floor", "", SpeechGrammarWordType.SGLexical, "floor", 18, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "truncate", "", SpeechGrammarWordType.SGLexical, "truncate", 19, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sine inverse", "", SpeechGrammarWordType.SGLexical, "sine inverse", 20, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "cos inverse", "", SpeechGrammarWordType.SGLexical, "cos inverse", 21, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "tan inverse", "", SpeechGrammarWordType.SGLexical, "tan inverse", 22, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "tan theta", "", SpeechGrammarWordType.SGLexical, "tan theta", 23, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sine theta", "", SpeechGrammarWordType.SGLexical, "sine theta", 24, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "cos theta", "", SpeechGrammarWordType.SGLexical, "cos theta", 25, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "hyperbolic sine", "", SpeechGrammarWordType.SGLexical, "hyperbolic sine", 26, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "hyoerbolic cos", "", SpeechGrammarWordType.SGLexical, "hyperbolic cos", 27, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "hyperbolic tan", "", SpeechGrammarWordType.SGLexical, "hyperbolic tan", 28, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "natural log", "", SpeechGrammarWordType.SGLexical, "natural log", 29, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "logarithm", "", SpeechGrammarWordType.SGLexical, "logarithm", 30, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "expression", "", SpeechGrammarWordType.SGLexical, "expression", 31, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "e power x", "", SpeechGrammarWordType.SGLexical, "e power x", 32, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "backspace", "", SpeechGrammarWordType.SGLexical, "backspace", 33, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "cancel", "", SpeechGrammarWordType.SGLexical, "cancel", 34, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "clear", "", SpeechGrammarWordType.SGLexical, "clear", 35, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "date", "", SpeechGrammarWordType.SGLexical, "date", 36, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "time", "", SpeechGrammarWordType.SGLexical, "time", 37, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "color", "", SpeechGrammarWordType.SGLexical, "color", 38, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "scientific", "", SpeechGrammarWordType.SGLexical, "scientific", 39, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "logicalfunctions", "", SpeechGrammarWordType.SGLexical, "logicalfunctions", 40, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "memoryfunctions", "", SpeechGrammarWordType.SGLexical, "memoryfunctions", 41, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "powerfunctions", "", SpeechGrammarWordType.SGLexical, "powerfunctions", 42, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "simple trignometric functions", "", SpeechGrammarWordType.SGLexical, "simple trignometric functions", 43, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "hyperbolic trignometric functions", "", SpeechGrammarWordType.SGLexical, "hyperbolic trignometric functions", 44, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "inverse trignometric functions", "", SpeechGrammarWordType.SGLexical, "inverse", 45, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "otherfunctions", "", SpeechGrammarWordType.SGLexical, "otherfunctions", 46, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "constants values", "", SpeechGrammarWordType.SGLexical, "constants values", 47, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "memory", "", SpeechGrammarWordType.SGLexical, "memory", 48, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "square", "", SpeechGrammarWordType.SGLexical, "square", 49, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "and", "", SpeechGrammarWordType.SGLexical, "and", 50, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "or", "", SpeechGrammarWordType.SGLexical, "or", 51, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "not", "", SpeechGrammarWordType.SGLexical, "not", 52, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "exor", "", SpeechGrammarWordType.SGLexical, "exor", 53, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "power", "", SpeechGrammarWordType.SGLexical, "power", 54, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sign", "", SpeechGrammarWordType.SGLexical, "sign", 55, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "cube", "", SpeechGrammarWordType.SGLexical, "cube", 56, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "factorial", "", SpeechGrammarWordType.SGLexical, "factorial", 57, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "squareroot", "", SpeechGrammarWordType.SGLexical, "squareroot", 58, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "round", "", SpeechGrammarWordType.SGLexical, "round", 59, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ceil", "", SpeechGrammarWordType.SGLexical, "ceil", 60, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "permutation", "", SpeechGrammarWordType.SGLexical, "permutation", 61, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "combination", "", SpeechGrammarWordType.SGLexical, "combination", 62, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "inverse", "", SpeechGrammarWordType.SGLexical, "inverse", 63, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "pai", "", SpeechGrammarWordType.SGLexical, "pai", 64, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "", "", SpeechGrammarWordType.SGLexical, "", 65, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "billion", "", SpeechGrammarWordType.SGLexical, "billion", 65, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "million", "", SpeechGrammarWordType.SGLexical, "million", 66, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "thousand", "", SpeechGrammarWordType.SGLexical, "thousand", 67, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "hundred", "", SpeechGrammarWordType.SGLexical, "hundred", 68, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "ten", "", SpeechGrammarWordType.SGLexical, "ten", 69, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "twenty", "", SpeechGrammarWordType.SGLexical, "twenty", 70, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "thirty", "", SpeechGrammarWordType.SGLexical, "thirty", 71, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "fourty", "", SpeechGrammarWordType.SGLexical, "fourty", 72, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "fifty", "", SpeechGrammarWordType.SGLexical, "fifty", 73, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "sixty", "", SpeechGrammarWordType.SGLexical, "sixty", 74, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "seventy", "", SpeechGrammarWordType.SGLexical, "seventy", 75, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "eighty", "", SpeechGrammarWordType.SGLexical, "eighty", 76, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "ninety", "", SpeechGrammarWordType.SGLexical, "ninety", 77, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "eleven", "", SpeechGrammarWordType.SGLexical, "eleven", 78, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "twelve", "", SpeechGrammarWordType.SGLexical, "twelve", 79, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "thirteen", "", SpeechGrammarWordType.SGLexical, "thirteen", 80, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "fourteen", "", SpeechGrammarWordType.SGLexical, "fourteen", 81, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "fifteen", "", SpeechGrammarWordType.SGLexical, "fifteen", 82, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "sixteen", "", SpeechGrammarWordType.SGLexical, "sixteen", 83, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "s eventeen", "", SpeechGrammarWordType.SGLexical, "seventeen", 84, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "eighteen", "", SpeechGrammarWordType.SGLexical, "eighteen", 85, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "nineteen", "", SpeechGrammarWordType.SGLexical, "nineteen", 86, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "text to speech", "", SpeechGrammarWordType.SGLexical, "text to speech", 87, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "string to text", "", SpeechGrammarWordType.SGLexical, "string to text", 88, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "equals", "", SpeechGrammarWordType.SGLexical, "equals", 89, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "grey", "", SpeechGrammarWordType.SGLexical, "grey", 90, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "blue", "", SpeechGrammarWordType.SGLexical, "blue", 91, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "violet", "", SpeechGrammarWordType.SGLexical, "violet", 92, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "speechlevel", "", SpeechGrammarWordType.SGLexical, "speechlevel", 93, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "advanced", "", SpeechGrammarWordType.SGLexical, "advanced", 94, PropValue, 1.0F)
        grammar.Rules.Commit()
        grammar.CmdSetRuleState("MenuCommands", SpeechRuleState.SGDSActive)
        label4.Text = "Speech Engine Ready for Input"
        aplay()
    End Sub

    Private Sub frmDictionary_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            With voip
                .StopSpeaking()
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmDictionary_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            With voip
                .StopSpeaking()
            End With
        Catch ex As Exception

        End Try
    End Sub


    Private Sub btnColour_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColour.Click
        btna.Visible = True
        btnb.Visible = True
        btnc.Visible = True
        Timer3.Enabled = True
        If ind = 1 Then
            Try
                With voip
                    .Speak(btnColour.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub btnDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDate.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak(btnDate.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Try
                With voip
                    .Speak(DateTimePicker1.Value.Date)
                End With

            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        If ind = 1 Then
            Try
                With voip
                    .Speak(DateTimePicker1.Value.Date)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button26.Click
        label4.Visible = True
        Button6.Enabled = True
        Button6.Visible = True
        button26.Enabled = False
        btnEqual.Enabled = False
        Button5.Visible = True
        voip.Speak("cahnge voice speed")
        label4.Text = "say the speed level"
        btn1.Enabled = False
        btn2.Enabled = False
        btn3.Enabled = False
        btn4.Enabled = False
        btn5.Enabled = False
        btn6.Enabled = False
        btn7.Enabled = False
        btn8.Enabled = False
        btn9.Enabled = False
        btn0.Enabled = False
        Button1.Visible = True
        Button2.Visible = True
        Button3.Visible = True
        Button4.Visible = True
        tbVoiceSpeed.Visible = True
    End Sub

    Private Sub tbVoiceSpeed_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVoiceSpeed.Scroll
        Try
            With voip
                Select Case tbVoiceSpeed.Value
                    Case 0
                        .StopSpeaking()
                        lblTrackBarValue.Text = "0"

                    Case 1
                        .Speed = ((450 * 10) / 100)
                        lblTrackBarValue.Text = "-1"

                    Case 2
                        .Speed = ((450 * 20) / 100)
                        lblTrackBarValue.Text = "-2"

                    Case 3
                        .Speed = ((450 * 30) / 100)
                        lblTrackBarValue.Text = "-3"

                    Case 4
                        .Speed = ((450 * 40) / 100)
                        lblTrackBarValue.Text = "-4"

                    Case 5
                        .Speed = ((450 * 50) / 100)
                        lblTrackBarValue.Text = "5"

                    Case 6
                        .Speed = ((450 * 60) / 100)
                        lblTrackBarValue.Text = "6"

                    Case 7
                        .Speed = ((450 * 70) / 100)
                        lblTrackBarValue.Text = "7"

                    Case 8
                        .Speed = ((450 * 80) / 100)
                        lblTrackBarValue.Text = "8"

                    Case 9
                        .Speed = ((450 * 90) / 100)
                        lblTrackBarValue.Text = "9"

                    Case 10
                        .Speed = ((450 * 100) / 100)
                        lblTrackBarValue.Text = "10"

                End Select
            End With
        Catch ex As Exception

        End Try
        Try
            With voip
                .Speak(lblTrackBarValue.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub
    Private Sub btnTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTime.Click
        If ind = 1 Then
            Try
                With voip
                    .Speak(btnTime.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Try
                With voip
                    .Speak(lblTime.Text)
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub Reco_Event(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult)
        txtReco.Text = Result.PhraseInfo.GetText(0, -1, True)
        If CheckBox1.Checked = True Then
            If voc = 1 Then
                If txtReco.Text = "two" Then
                    lblTrackBarValue.Text = "2"
                    voip.Speed = ((450 * 40) / 100)
                    tbVoiceSpeed.Value = 4
                    Try
                        With voip
                            .Speak(lblTrackBarValue.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "three" Then
                    voip.Speed = ((450 * 10) / 100)
                    lblTrackBarValue.Text = "3"
                    tbVoiceSpeed.Value = 0
                    Try
                        With voip
                            .Speak(lblTrackBarValue.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "ok" Then
                    voc = 0
                    Try
                        With voip
                            .Speak("ok")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    Button5.Visible = False
                    Button5.Enabled = False
                    label4.Text = "Speechengine ready for input"
                    Button6.Enabled = False
                    Button6.Visible = False
                    button26.Enabled = True
                    btnEqual.Enabled = True
                    Button1.Visible = False
                    Button2.Visible = False
                    Button3.Visible = False
                    Button4.Visible = False
                End If
                If txtReco.Text = "five" Then
                    lblTrackBarValue.Text = "5"
                    voip.Speed = ((450 * 50) / 100)
                    tbVoiceSpeed.Value = 5
                    Try
                        With voip
                            .Speak(lblTrackBarValue.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "seven" Then
                    voip.Speed = ((450 * 70) / 100)
                    lblTrackBarValue.Text = "7"
                    tbVoiceSpeed.Value = 7
                    Try
                        With voip
                            .Speak(lblTrackBarValue.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "ten" Then
                    voip.Speed = ((450 * 100) / 100)
                    lblTrackBarValue.Text = "10"
                    tbVoiceSpeed.Value = 10
                    Try
                        With voip
                            .Speak(lblTrackBarValue.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If voc = 0 Then
                If txtReco.Text = "one" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(1)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 1
                        sign_Indicator = 0
                    End If
                    fl = 1
                    If ind = 1 Then
                        Try
                            With voip
                                .Speak("one")
                            End With
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
                If txtReco.Text = "two" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(2)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 2
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("two")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try

                End If
                If txtReco.Text = "three" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(3)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 3
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("three")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "four" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(4)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 4
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("four")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "five" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(5)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 5
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("five")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "six" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(6)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 6
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("six")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "seven" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(7)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 7
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("Seven")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "eight" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(8)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 8
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("eight")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "nine" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(9)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 9
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("nine")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
                If txtReco.Text = "zero" Then
                    If sign_Indicator = 0 Then
                        displayText.Text = displayText.Text & CStr(0)
                    ElseIf sign_Indicator = 1 Then
                        displayText.Text = 0
                        sign_Indicator = 0
                    End If
                    fl = 1
                    Try
                        With voip
                            .Speak("zero")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "plus" Then
                If displayText.Text.Length <> 0 Then
                    Calculate()
                    reset_SignBits()
                    addBit = 1
                    sign_Indicator = 1
                End If
                Try
                    With voip
                        .Speak("plus")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "minus" Then
                Try
                    With voip
                        .Speak("minus")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                If displayText.Text.Length <> 0 Then
                    variable2 = displayText.Text
                    Calculate()
                    reset_SignBits()
                    subBit = 1
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "multiply" Then
                If displayText.Text.Length <> 0 Then
                    Calculate()
                    reset_SignBits()
                    multBit = 1
                    sign_Indicator = 1
                End If
                Try
                    With voip
                        .Speak("multiply")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "divide" Then
                If displayText.Text.Length <> 0 Then
                    Calculate()
                    reset_SignBits()
                    divBit = 1
                    sign_Indicator = 1
                End If
                Try
                    With voip
                        .Speak("Divide")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "dot" Then
                If displayText.Text = "" Then
                    Try
                        With voip
                            .Speak(lblInfoDecimal.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    lblInfoDecimal.Visible = True
                    Timer1.Enabled = True
                End If
                Dim i As Integer
                Dim chr As Char
                Dim decimal_Indicator As Integer = 0

                If sign_Indicator <> 1 Then
                    For i = 0 To displayText.TextLength - 1
                        chr = displayText.Text(i)
                        If chr = "." Then
                            decimal_Indicator = 1
                        End If
                    Next

                    If decimal_Indicator <> 1 Then
                        displayText.Text = displayText.Text & CStr(".")
                    End If
                End If
                Try
                    With voip
                        .Speak("dot")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                If displayText.Text = "" Then
                    lblInfoDecimal.Visible = True
                End If
            End If
            If txtReco.Text = "equals" Then
                Try
                    With voip
                        .Speak("equals")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                If displayText.Text.Length <> 0 Then
                    Calculate()
                    reset_SignBits()
                End If
                sign_Indicator = 1
                Try
                    With voip
                        .Speak(displayText.Text)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "scientific" Then
                If scientificModeBit = 0 Then
                    Label3.Visible = True
                    sci = 1
                    pow = 1
                    mem = 1
                    logic = 1
                    str = 1
                    itr = 1
                    htr = 1
                    oth = 1
                    Try
                        With voip
                            .Speak("Scientific mode enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    btnFact.BackColor = Color.Coral
                    btnPerm.BackColor = Color.Coral
                    btnComb.BackColor = Color.Coral
                    btnPI.BackColor = Color.Coral
                    btnsqrt.BackColor = Color.Coral
                    btnlog.BackColor = Color.Coral
                    btnex.BackColor = Color.Coral
                    btnln.BackColor = Color.Coral
                    btnInverse.BackColor = Color.Coral
                    btnModulus.BackColor = Color.Coral
                    btne.BackColor = Color.Coral
                    btnPower.BackColor = Color.Coral
                    btnSquare.BackColor = Color.Coral
                    btnCube.BackColor = Color.Coral
                    btnAnd.BackColor = Color.Coral
                    btnOr.BackColor = Color.Coral
                    btnNot.BackColor = Color.Coral
                    btnXor.BackColor = Color.Coral
                    btnSin.BackColor = Color.Coral
                    btnCos.BackColor = Color.Coral
                    btnTan.BackColor = Color.Coral
                    btnSinh.BackColor = Color.Coral
                    btnCosh.BackColor = Color.Coral
                    btnTanh.BackColor = Color.Coral
                    btnInvSin.BackColor = Color.Coral
                    btnInvCos.BackColor = Color.Coral
                    btnInvTan.BackColor = Color.Coral
                    btnRound.BackColor = Color.Coral
                    btnFloor.BackColor = Color.Coral
                    btnTruncate.BackColor = Color.Coral
                    btnCeil.BackColor = Color.Coral
                    btnM.BackColor = Color.Coral


                    scientificModeBit = 1
                ElseIf scientificModeBit = 1 Then
                    Dim sci As Integer = 0
                    Label3.Visible = False
                    btnFact.BackColor = Color.DarkGray
                    btnPerm.BackColor = Color.DarkGray
                    btnComb.BackColor = Color.DarkGray
                    btnPI.BackColor = Color.DarkGray
                    btnsqrt.BackColor = Color.DarkGray
                    btnlog.BackColor = Color.DarkGray
                    btnex.BackColor = Color.DarkGray
                    btnln.BackColor = Color.DarkGray
                    btnInverse.BackColor = Color.DarkGray
                    btnModulus.BackColor = Color.DarkGray
                    btne.BackColor = Color.DarkGray
                    btnPower.BackColor = Color.DarkGray
                    btnSquare.BackColor = Color.DarkGray
                    btnCube.BackColor = Color.DarkGray
                    btnAnd.BackColor = Color.DarkGray
                    btnOr.BackColor = Color.DarkGray
                    btnNot.BackColor = Color.DarkGray
                    btnXor.BackColor = Color.DarkGray
                    btnSin.BackColor = Color.DarkGray
                    btnCos.BackColor = Color.DarkGray
                    btnTan.BackColor = Color.DarkGray
                    btnSinh.BackColor = Color.DarkGray
                    btnCosh.BackColor = Color.DarkGray
                    btnTanh.BackColor = Color.DarkGray
                    btnInvSin.BackColor = Color.DarkGray
                    btnInvCos.BackColor = Color.DarkGray
                    btnInvTan.BackColor = Color.DarkGray
                    btnRound.BackColor = Color.DarkGray
                    btnFloor.BackColor = Color.DarkGray
                    btnTruncate.BackColor = Color.DarkGray
                    btnCeil.BackColor = Color.DarkGray
                    btnM.BackColor = Color.DarkGray

                    scientificModeBit = 0
                    Try
                        With voip
                            .Speak("Scientific mode disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "time" Then
                Try
                    With voip
                        .Speak(btnTime.Text)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                Try
                    With voip
                        .Speak(lblTime.Text)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "advanced" Then
                Try
                    With voip
                        .Speak("Advanced calculator")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                Dim advcal As New ADV
                advcal.Show()
            End If
            If txtReco.Text = "date" Then
                Try
                    With voip
                        .Speak(btnDate.Text)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                Try
                    With voip
                        .Speak(DateTimePicker1.Value.Date)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "color" Then
                btna.Visible = True
                btnb.Visible = True
                btnc.Visible = True
                btna.Enabled = False
                btnb.Enabled = False
                btnc.Enabled = False
                clor = 1
                Try
                    With voip
                        .Speak(btnColour.Text)
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                Timer3.Enabled = True
            End If
            If clor = 1 Then
                If txtReco.Text = "grey" Then
                    BackColor = Color.DimGray
                    btn0.BackColor = Color.Chocolate
                    btn1.BackColor = Color.Chocolate
                    btn2.BackColor = Color.Chocolate
                    btn3.BackColor = Color.Chocolate
                    btn4.BackColor = Color.Chocolate
                    btn5.BackColor = Color.Chocolate
                    btn6.BackColor = Color.Chocolate
                    btn7.BackColor = Color.Chocolate
                    btn8.BackColor = Color.Chocolate
                    btn9.BackColor = Color.Chocolate
                    If ind = 1 Then
                        Try
                            With voip
                                .Speak("grey")
                            End With
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
                If txtReco.Text = "violet" Then
                    BackColor = Color.BlueViolet
                    btn0.BackColor = Color.DimGray
                    btn1.BackColor = Color.DimGray
                    btn2.BackColor = Color.DimGray
                    btn3.BackColor = Color.DimGray
                    btn4.BackColor = Color.DimGray
                    btn5.BackColor = Color.DimGray
                    btn6.BackColor = Color.DimGray
                    btn7.BackColor = Color.DimGray
                    btn8.BackColor = Color.DimGray
                    btn9.BackColor = Color.DimGray
                    If ind = 1 Then
                        Try
                            With voip
                                .Speak("Grey")
                            End With
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
                If txtReco.Text = "blue" Then
                    BackColor = Color.Blue
                    btn0.BackColor = Color.DarkRed
                    btn1.BackColor = Color.DarkRed
                    btn2.BackColor = Color.DarkRed
                    btn3.BackColor = Color.DarkRed
                    btn4.BackColor = Color.DarkRed
                    btn5.BackColor = Color.DarkRed
                    btn6.BackColor = Color.DarkRed
                    btn7.BackColor = Color.DarkRed
                    btn8.BackColor = Color.DarkRed
                    btn9.BackColor = Color.DarkRed
                    If ind = 1 Then
                        Try
                            With voip
                                .Speak("Red")
                            End With
                        Catch ex As Exception
                            MsgBox(ex.Message())
                        End Try
                    End If
                End If
            End If
            If txtReco.Text = "memoryfunctions" Then
                If memFuncsBit = 0 Then
                    Label6.Visible = True
                    btnM.BackColor = Color.Coral
                    memFuncsBit = 1
                    Try
                        With voip
                            .Speak("Memory functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    mem = 1
                ElseIf memFuncsBit = 1 Then
                    mem = 0
                    Label6.Visible = False
                    btnM.BackColor = Color.DarkGray

                    memFuncsBit = 0
                    Try
                        With voip
                            .Speak("Memory functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "powerfunctions" Then
                If powerFunctionBit = 0 Then
                    pow = 1
                    Label7.Visible = True
                    btnPower.BackColor = Color.Coral
                    btnSquare.BackColor = Color.Coral
                    btnCube.BackColor = Color.Coral
                    powerFunctionBit = 1
                    Try
                        With voip
                            .Speak("Power functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                ElseIf powerFunctionBit = 1 Then
                    pow = 0
                    Label7.Visible = False
                    btnPower.BackColor = Color.DarkGray
                    btnSquare.BackColor = Color.DarkGray
                    btnCube.BackColor = Color.DarkGray

                    powerFunctionBit = 0
                    Try
                        With voip
                            .Speak("power functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "logicalfunctions" Then
                If logicalFuncsBit = 0 Then
                    logic = 1
                    Label8.Visible = True
                    btnAnd.BackColor = Color.Coral
                    btnOr.BackColor = Color.Coral
                    btnNot.BackColor = Color.Coral
                    btnXor.BackColor = Color.Coral
                    logicalFuncsBit = 1
                    Try
                        With voip
                            .Speak("logical functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                ElseIf logicalFuncsBit = 1 Then
                    logic = 0
                    Label8.Visible = False
                    btnAnd.BackColor = Color.DarkGray
                    btnOr.BackColor = Color.DarkGray
                    btnNot.BackColor = Color.DarkGray
                    btnXor.BackColor = Color.DarkGray
                    logicalFuncsBit = 0
                    Try
                        With voip
                            .Speak("logical functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "simple trignometric functions" Then
                If trigFunctionBit = 0 Then
                    str = 1
                    Label9.Visible = True
                    btnSin.BackColor = Color.Coral
                    btnCos.BackColor = Color.Coral
                    btnTan.BackColor = Color.Coral
                    trigFunctionBit = 1
                    Try
                        With voip
                            .Speak("Simple Trignometric functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                ElseIf trigFunctionBit = 1 Then
                    str = 0
                    Label9.Visible = False
                    btnSin.BackColor = Color.DarkGray
                    btnCos.BackColor = Color.DarkGray
                    btnTan.BackColor = Color.DarkGray
                    trigFunctionBit = 0
                    Try
                        With voip
                            .Speak("simple Trignometric functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "hyperbolic trignometric functions" Then
                If HypertrigFunctionBit = 0 Then
                    htr = 1
                    Label11.Visible = True
                    btnSinh.BackColor = Color.Coral
                    btnCosh.BackColor = Color.Coral
                    btnTanh.BackColor = Color.Coral
                    HypertrigFunctionBit = 1
                    Try
                        With voip
                            .Speak("hyperbolic trignometric functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                ElseIf HypertrigFunctionBit = 1 Then
                    htr = 0
                    Label11.Visible = False
                    btnSinh.BackColor = Color.DarkGray
                    btnCosh.BackColor = Color.DarkGray
                    btnTanh.BackColor = Color.DarkGray
                    HypertrigFunctionBit = 0
                    Try
                        With voip
                            .Speak("hyperbolic trignometric functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "inverse trignometric functions" Then
                If InversetrigFunctionBit = 0 Then
                    itr = 1
                    Label10.Visible = True
                    btnInvSin.BackColor = Color.Coral
                    btnInvCos.BackColor = Color.Coral
                    btnInvTan.BackColor = Color.Coral
                    InversetrigFunctionBit = 1
                    Try
                        With voip
                            .Speak("inverse trignometric functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                ElseIf InversetrigFunctionBit = 1 Then
                    itr = 0
                    Label10.Visible = False
                    btnInvSin.BackColor = Color.DarkGray
                    btnInvCos.BackColor = Color.DarkGray
                    btnInvTan.BackColor = Color.DarkGray
                    InversetrigFunctionBit = 0
                    Try
                        With voip
                            .Speak("inverse trignometric functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "otherfunctions" Then
                If otherFuncsBit = 0 Then
                    oth = 1
                    btnRound.BackColor = Color.Coral
                    btnFloor.BackColor = Color.Coral
                    btnTruncate.BackColor = Color.Coral
                    btnCeil.BackColor = Color.Coral
                    otherFuncsBit = 1
                    Try
                        With voip
                            .Speak("other functions enabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                ElseIf otherFuncsBit = 1 Then
                    oth = 0
                    btnRound.BackColor = Color.DarkGray
                    btnFloor.BackColor = Color.DarkGray
                    btnTruncate.BackColor = Color.DarkGray
                    btnCeil.BackColor = Color.DarkGray
                    otherFuncsBit = 0
                    Try
                        With voip
                            .Speak("other functions disabled")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "constants values" Then
                Dim new_Form As New FormConstants
                Try
                    With voip
                        .Speak("Constants values")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                new_Form.Show()

            End If
        End If

        If txtReco.Text = "voice" Then
            If CheckBox1.Checked Then
                CheckBox2.Enabled = True
                CheckBox1.Checked = False
                btnFact.BackColor = Color.DarkGray
                btnPerm.BackColor = Color.DarkGray
                btnComb.BackColor = Color.DarkGray
                btnPI.BackColor = Color.DarkGray
                btnsqrt.BackColor = Color.DarkGray
                btnlog.BackColor = Color.DarkGray
                btnex.BackColor = Color.DarkGray
                btnln.BackColor = Color.DarkGray
                btnInverse.BackColor = Color.DarkGray
                btnModulus.BackColor = Color.DarkGray
                btne.BackColor = Color.DarkGray
                btnPower.BackColor = Color.DarkGray
                btnSquare.BackColor = Color.DarkGray
                btnCube.BackColor = Color.DarkGray
                btnAnd.BackColor = Color.DarkGray
                btnOr.BackColor = Color.DarkGray
                btnNot.BackColor = Color.DarkGray
                btnXor.BackColor = Color.DarkGray
                btnSin.BackColor = Color.DarkGray
                btnCos.BackColor = Color.DarkGray
                btnTan.BackColor = Color.DarkGray
                btnSinh.BackColor = Color.DarkGray
                btnCosh.BackColor = Color.DarkGray
                btnTanh.BackColor = Color.DarkGray
                btnInvSin.BackColor = Color.DarkGray
                btnInvCos.BackColor = Color.DarkGray
                btnInvTan.BackColor = Color.DarkGray
                btnRound.BackColor = Color.DarkGray
                btnFloor.BackColor = Color.DarkGray
                btnTruncate.BackColor = Color.DarkGray
                btnCeil.BackColor = Color.DarkGray
                btnM.BackColor = Color.DarkGray
            Else
                CheckBox1.Checked = True
                CheckBox2.Enabled = False
            End If
            Try
                With voip
                    .Speak("voice")
                End With
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        'If txtReco.Text = "text to speech" Then
        'If CheckBox2.Checked Then
        'CheckBox2.Checked = False
        ' Else
        ' CheckBox2.Checked = True
        'End If
        'Try
        'With voip
        '.Speak("text to speech")
        'End With
        'Catch ex As Exception
        'MsgBox(ex.Message())
        'End Try
        'End If
        If CheckBox1.Checked = True Then
           
            If txtReco.Text = "speechlevel" Then
                If voc = 0 Then
                    voc = 1
                Else
                    voc = 0
                End If
                label4.Visible = False
                Button6.Enabled = True
                Button6.Visible = True
                button26.Enabled = False
                btnEqual.Enabled = False
                Button5.Visible = True
                voip.Speak("change voice speed")
                label4.Text = "say voice speed"
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                btn7.Enabled = False
                btn8.Enabled = False
                btn9.Enabled = False
                btn0.Enabled = False
                Button1.Visible = True
                Button2.Visible = True
                Button3.Visible = True
                Button4.Visible = True
                tbVoiceSpeed.Visible = True
            End If
            If txtReco.Text = "memory" Then
                If mem = 1 Then
                    If displayText.Text.Length <> 0 Then
                        If memoryVariable = 0 Then
                            memoryVariable = CDbl(displayText.Text)
                        End If
                    End If
                    sign_Indicator = 1
                    displayText.Text = CStr(memoryVariable)
                    Try
                        With voip
                            .Speak("memory")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "backspace" Then
                Try
                    With voip
                        .Speak("Backspace")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                If displayText.Text.Length <> 0 Then
                    displayText.Text = displayText.Text.Remove(displayText.Text.Length - 1, 1)
                End If

            End If
            If txtReco.Text = "clear" Then
                displayText.Clear()
                Try
                    With voip
                        .Speak("clear")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "cancel" Then
                displayText.Clear()
                sign_Indicator = 0
                variable1 = 0
                variable2 = 0
                memoryVariable = 0
                reset_SignBits()
                Try
                    With voip
                        .Speak("clear all")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            If txtReco.Text = "sign" Then

                If displayText.Text.Length = 0 Then
                    displayText.Text = displayText.Text + CStr("-")
                ElseIf displayText.Text <> "." Then
                    displayText.Text = displayText.Text * -1
                End If
                Try
                    With voip
                        .Speak("Sign changed")
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try

            End If
            If txtReco.Text = "factorial" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Dim temp As Double = displayText.Text
                        results = Factorial(temp)
                        displayText.Text = CStr(results)
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("factorial")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "natural log" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Log(displayText.Text)
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("natural logarithm")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "squareroot" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        If displayText.Text <> "." Then
                            displayText.Text = Math.Sqrt(displayText.Text)
                        End If
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("Square root")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "expression" Then
                If oth = 1 Then
                    displayText.Text = Math.E
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("Expression")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "sine theta" Then
                If str = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Sin(displayText.Text)
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("sign")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "cos theta" Then
                If str = 1 Then
                    Try
                        With voip
                            .Speak("cos")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Cos(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "tan theta" Then
                If str = 1 Then
                    Try
                        With voip
                            .Speak("tangent")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Tan(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "ceil" Then
                If oth = 1 Then
                    Try
                        With voip
                            .Speak("ceil")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        Dim temp As Double = displayText.Text
                        displayText.Text = Math.Ceiling(temp)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "round" Then
                If oth = 1 Then
                    Try
                        With voip
                            .Speak("Round")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        Dim temp As Double = displayText.Text
                        displayText.Text = Math.Round(temp)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "percentage" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        modBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("percentage")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "floor" Then
                If oth = 1 Then
                    Try
                        With voip
                            .Speak("Floor")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        Dim temp As Double = displayText.Text
                        displayText.Text = Math.Floor(temp)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "truncate" Then
                If oth = 1 Then
                    Try
                        With voip
                            .Speak("Truncate")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        Dim temp As Double = displayText.Text
                        displayText.Text = Math.Truncate(temp)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "tan inverse" Then
                If itr = 1 Then
                    Try
                        With voip
                            .Speak("cot")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Atan(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If

            If txtReco.Text = "sine inverse" Then
                If itr = 1 Then
                    Try
                        With voip
                            .Speak("Secent")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Asin(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "cos inverse" Then
                If itr = 1 Then
                    Try
                        With voip
                            .Speak("coSecent")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try

                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Acos(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "inverse" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = 1 / displayText.Text
                    End If
                    Try
                        With voip
                            .Speak("one by x")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "pai" Then
                If oth = 1 Then
                    displayText.Text = Math.PI
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("pai")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "e power x" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Exp(displayText.Text)
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("exponent")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "logarithm" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Log10(displayText.Text)
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("logarithm")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "combination" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        combBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("combination")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "permutation" Then
                If oth = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        permBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("permutation")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "and" Then
                If logic = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        andBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("logical And")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "or" Then
                If logic = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        orBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("logical or")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "not" Then
                If logic = 1 Then
                    If displayText.Text.Length <> 0 Then
                        variable1 = CDbl(displayText.Text)
                        variable1 = Not variable1
                        displayText.Text = CStr(variable1)
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("logical not")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    Try
                        With voip
                            .Speak(displayText.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "exor" Then
                If logic = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        xorBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("logical x or")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "square" Then
                If pow = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = displayText.Text * displayText.Text
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("square")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    Try
                        With voip
                            .Speak(displayText.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "cube" Then
                If pow = 1 Then
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = displayText.Text * displayText.Text * displayText.Text
                    End If
                    sign_Indicator = 1
                    Try
                        With voip
                            .Speak("cube")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    Try
                        With voip
                            .Speak(displayText.Text)
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "power" Then
                If pow = 1 Then
                    If displayText.Text.Length <> 0 Then
                        Calculate()
                        reset_SignBits()
                        powerBit = 1
                        sign_Indicator = 1
                    End If
                    Try
                        With voip
                            .Speak("power n")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End If
            End If
            If txtReco.Text = "hyperbolic sine" Then
                If htr = 1 Then
                    Try
                        With voip
                            .Speak("sign h")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Sinh(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "hyperbolic cos" Then
                If htr = 1 Then
                    Try
                        With voip
                            .Speak("cos h")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Cosh(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If
            If txtReco.Text = "hyperbolic tan" Then
                If htr = 1 Then
                    Try
                        With voip
                            .Speak("tan h")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    If displayText.Text.Length <> 0 Then
                        displayText.Text = Math.Tanh(displayText.Text)
                    End If
                    sign_Indicator = 1
                End If
            End If

        End If
    End Sub 'Reco_Event


    Private Sub Hypo_Event(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal Result As ISpeechRecoResult)
        txtHyp.Text = Result.PhraseInfo.GetText(0, -1, True)
        txtReco.Text = Result.PhraseInfo.GetText(0, -1, True)
    End Sub 'Hypo_Event

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        lblTrackBarValue.Text = "5"
        voip.Speed = ((450 * 50) / 100)
        tbVoiceSpeed.Value = 5
        Try
            With voip
                .Speak(lblTrackBarValue.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        lblTrackBarValue.Text = "-4"
        voip.Speed = ((450 * 40) / 100)
        tbVoiceSpeed.Value = 4
        Try
            With voip
                .Speak(lblTrackBarValue.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        voip.Speed = ((450 * 70) / 100)
        lblTrackBarValue.Text = "7"
        tbVoiceSpeed.Value = 7
        Try
            With voip
                .Speak(lblTrackBarValue.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        voip.Speed = ((450 * 100) / 100)
        lblTrackBarValue.Text = "10"
        tbVoiceSpeed.Value = 10
        Try
            With voip
                .Speak(lblTrackBarValue.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        voip.Speed = ((450 * 10) / 100)
        lblTrackBarValue.Text = "0"
        tbVoiceSpeed.Value = 0
        Try
            With voip
                .Speak(lblTrackBarValue.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub Timer1_Tick_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        tim = tim + 1
        If tim = 5 Then
            lblInfoDecimal.Visible = False
            tim = 0
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            With voip
                .Speak("ok")
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Button5.Visible = False
        Button5.Enabled = False
        label4.Text = "Speechengine ready for input"
        Button6.Enabled = False
        Button6.Visible = False
        button26.Enabled = True
        btnEqual.Enabled = True
        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        tbVoiceSpeed.Visible = False
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        lblTime.Text = TimeString
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckBox2.Checked = False
            CheckBox2.Enabled = False
            btnBackSpace.Enabled = False
            btnClearTextbox.Enabled = False
            btnClearAll.Enabled = False
            btnEqual.Enabled = False
            btnSign.Enabled = False
            btnFact.Enabled = False
            btnSub.Enabled = False
            btnAdd.Enabled = False
            btnMult.Enabled = False
            btnDiv.Enabled = False
            btnTime.Enabled = False
            btnColour.Enabled = False
            btnDate.Enabled = False
            btnScientificMode.Enabled = False
            btnMemFuncs.Enabled = False
            btnPowerFunctions.Enabled = False
            btnLogicalFuncs.Enabled = False
            btnTrignometric.Enabled = False
            btnHyperbolicTrig.Enabled = False
            btnInverseTrigFuncs.Enabled = False
            btnOtherFuncs.Enabled = False
            btnConstants.Enabled = False
            btnPerm.Enabled = False
            btnDecimal.Enabled = False

            btnEqual.Enabled = False
            button26.Enabled = False
            btn1.Enabled = False
            btn2.Enabled = False
            btn3.Enabled = False
            btn4.Enabled = False
            btn5.Enabled = False
            btn6.Enabled = False
            btn7.Enabled = False
            btn8.Enabled = False
            btn9.Enabled = False
            btn0.Enabled = False
            btnComb.Enabled = False
            btnPI.Enabled = False
            btnsqrt.Enabled = False
            btnlog.Enabled = False
            btnex.Enabled = False
            btnln.Enabled = False
            btnInverse.Enabled = False
            btnModulus.Enabled = False
            btne.Enabled = False
            btnPower.Enabled = False
            btnSquare.Enabled = False
            btnCube.Enabled = False
            btnAnd.Enabled = False
            btnOr.Enabled = False
            btnNot.Enabled = False
            btnXor.Enabled = False
            btnSin.Enabled = False
            btnCos.Enabled = False
            btnTan.Enabled = False
            btnSinh.Enabled = False
            btnCosh.Enabled = False
            btnTanh.Enabled = False
            btnInvSin.Enabled = False
            btnInvCos.Enabled = False
            btnInvTan.Enabled = False
            btnRound.Enabled = False
            btnFloor.Enabled = False
            btnTruncate.Enabled = False
            btnCeil.Enabled = False
            btnM.Enabled = False
            DateTimePicker1.Enabled = False
            scientificModeBit = 0
        Else
            CheckBox2.Enabled = True
            btnDecimal.Enabled = True
            DateTimePicker1.Enabled = True
            btnBackSpace.Enabled = True
            btnClearTextbox.Enabled = True
            btnClearAll.Enabled = True
            btnEqual.Enabled = True
            btnScientificMode.Enabled = True
            btnPerm.Enabled = True
            btnSign.Enabled = True
            btnFact.Enabled = True
            btnSub.Enabled = True
            btnAdd.Enabled = True
            btnMult.Enabled = True
            btnDiv.Enabled = True
            btnTime.Enabled = True
            btnColour.Enabled = True
            btnDate.Enabled = True
            btnMemFuncs.Enabled = True
            btnPowerFunctions.Enabled = True
            btnLogicalFuncs.Enabled = True
            btnTrignometric.Enabled = True
            btnHyperbolicTrig.Enabled = True
            btnInverseTrigFuncs.Enabled = True
            btnOtherFuncs.Enabled = True
            btnConstants.Enabled = True

            btnEqual.Enabled = True
            button26.Enabled = True
            btn1.Enabled = True
            btn2.Enabled = True
            btn3.Enabled = True
            btn4.Enabled = True
            btn5.Enabled = True
            btn6.Enabled = True
            btn7.Enabled = True
            btn8.Enabled = True
            btn9.Enabled = True
            btn0.Enabled = True
            btnFact.Enabled = 1
            btnPerm.Enabled = 1
            btnComb.Enabled = 1
            btnPI.Enabled = 1
            btnsqrt.Enabled = 1
            btnlog.Enabled = 1
            btnex.Enabled = 1
            btnln.Enabled = 1
            btnInverse.Enabled = 1
            btnModulus.Enabled = 1
            btne.Enabled = 1
            btnPower.Enabled = 1
            btnSquare.Enabled = 1
            btnCube.Enabled = 1
            btnAnd.Enabled = 1
            btnOr.Enabled = 1
            btnNot.Enabled = 1
            btnXor.Enabled = 1
            btnSin.Enabled = 1
            btnCos.Enabled = 1
            btnTan.Enabled = 1
            btnSinh.Enabled = 1
            btnCosh.Enabled = 1
            btnTanh.Enabled = 1
            btnInvSin.Enabled = 1
            btnInvCos.Enabled = 1
            btnInvTan.Enabled = 1
            btnRound.Enabled = 1
            btnFloor.Enabled = 1
            btnTruncate.Enabled = 1
            btnCeil.Enabled = 1
            btnM.Enabled = 1


            scientificModeBit = 1
        End If
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        time = time + 1
        If time = 5 Then
            time = 0
            Timer3.Enabled = False
            btna.Visible = False
            btnb.Visible = False
            btnc.Visible = False
            clor = 0
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Try
            With voip
                .Speak("text to speech")
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        If CheckBox2.Checked Then
            ind = 1
            CheckBox1.Checked = True

            btnDecimal.Enabled = True
            DateTimePicker1.Enabled = True
            btnBackSpace.Enabled = True
            btnClearTextbox.Enabled = True
            btnClearAll.Enabled = True
            btnEqual.Enabled = True
            btnScientificMode.Enabled = True
            btnPerm.Enabled = True
            btnSign.Enabled = True
            btnFact.Enabled = True
            btnSub.Enabled = True
            btnAdd.Enabled = True
            btnMult.Enabled = True
            btnDiv.Enabled = True
            btnTime.Enabled = True
            btnColour.Enabled = True
            btnDate.Enabled = True
            btnMemFuncs.Enabled = True
            btnPowerFunctions.Enabled = True
            btnLogicalFuncs.Enabled = True
            btnTrignometric.Enabled = True
            btnHyperbolicTrig.Enabled = True
            btnInverseTrigFuncs.Enabled = True
            btnOtherFuncs.Enabled = True
            btnConstants.Enabled = True
            button26.Enabled = True
            btnEqual.Enabled = True
            btn1.Enabled = True
            btn2.Enabled = True
            btn3.Enabled = True
            btn4.Enabled = True
            btn5.Enabled = True
            btn6.Enabled = True
            btn7.Enabled = True
            btn8.Enabled = True
            btn9.Enabled = True
            btn0.Enabled = True
            btnFact.Enabled = 1
            btnPerm.Enabled = 1
            btnComb.Enabled = 1
            btnPI.Enabled = 1
            btnsqrt.Enabled = 1
            btnlog.Enabled = 1
            btnex.Enabled = 1
            btnln.Enabled = 1
            btnInverse.Enabled = 1
            btnModulus.Enabled = 1
            btne.Enabled = 1
            btnPower.Enabled = 1
            btnSquare.Enabled = 1
            btnCube.Enabled = 1
            btnAnd.Enabled = 1
            btnOr.Enabled = 1
            btnNot.Enabled = 1
            btnXor.Enabled = 1
            btnSin.Enabled = 1
            btnCos.Enabled = 1
            btnTan.Enabled = 1
            btnSinh.Enabled = 1
            btnCosh.Enabled = 1
            btnTanh.Enabled = 1
            btnInvSin.Enabled = 1
            btnInvCos.Enabled = 1
            btnInvTan.Enabled = 1
            btnRound.Enabled = 1
            btnFloor.Enabled = 1
            btnTruncate.Enabled = 1
            btnCeil.Enabled = 1
            btnM.Enabled = 1
            CheckBox1.Enabled = False

            scientificModeBit = 1
        End If
        If CheckBox2.Checked = False Then
            ind = 0
            CheckBox1.Checked = False
        End If

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        msc = msc + 1
        If msc = 100 Then
            msc = 0
            Timer4.Enabled = False
            astop()
        End If
    End Sub

    Private Sub advc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles advc.Click
        Try
            With voip
                .Speak("Advanced calculator")
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        Dim advcal As New ADV
        advcal.Show()
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        muc = muc + 1
        If muc = 9 Then
            With voip
                .Speak("Welcome To voice calculator")
            End With
            With voip
                .Speak("i'm listening")
            End With
            muc = 0
            Timer5.Enabled = False
        End If


    End Sub
End Class
