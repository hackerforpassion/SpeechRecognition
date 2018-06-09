Imports SpeechLib
Imports System.Runtime.InteropServices
Public Class FormConstants
    Dim voip As New HTTSLib.TextToSpeechClass
    Private objRecoContext As SpeechLib.SpSharedRecoContext = Nothing
    Private grammar As SpeechLib.ISpeechRecoGrammar = Nothing
    Private menuRule As SpeechLib.ISpeechGrammarRule = Nothing
    Dim s As String

    Private Sub btnSpeed_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeed.MouseEnter
        t1.Text = "3.0 x 10^8 m/s" + "  " + "the symbol is c"
        
    End Sub

    Private Sub btnSpeed_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeed.MouseLeave
        t1.Clear()
    End Sub

    Private Sub btnCoulombs_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCoulombs.MouseEnter
        t1.Text = "8.987 x 10^9 Nm^2/Coulombs^2" + "  " + "the symbol is cc"
        
    End Sub

    Private Sub btnCoulombs_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCoulombs.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnGravitational_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGravitational.MouseEnter
        t1.Text = "9.806 m/s^2" + "  " + "the symbol is g"
        
    End Sub

    Private Sub btnGravitational_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGravitational.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnPlanks_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlanks.MouseEnter
        t1.Text = "6.67 x 10^-11 m^3/kg/s^2" + "  " + "the symbol is h"
        
    End Sub

    Private Sub btnPlanks_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlanks.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnBoltzmans_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBoltzmans.MouseEnter
        t1.Text = "1.3807 x 10^-23 J/K" + "  " + "the symbol is k"
        
    End Sub

    Private Sub btnBoltzmans_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBoltzmans.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnCharge_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCharge.MouseEnter
        t1.Text = "1.6022 x 10^-22 coulombs" + "  " + "the symbol is e"
        
    End Sub

    Private Sub btnCharge_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCharge.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnPermitivity_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPermitivity.MouseEnter
        t1.Text = "8.85 x 10^-12 F/m" + "  " + "the symbol is E"
        
    End Sub

    Private Sub btnPermitivity_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPermitivity.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnPermeability_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPermeability.MouseEnter
        t1.Text = "1.2566 x 10 ^ -6 N/A^2" + "  " + "the symbol is u"
        
    End Sub

    Private Sub btnPermeability_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPermeability.MouseLeave
        t1.Clear()

    End Sub

    Private Sub btnImpedence_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImpedence.MouseEnter
        t1.Text = "3.767 x 10^2 ohms" + "  " + "the symbol is z"
        
    End Sub

    Private Sub btnImpedence_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImpedence.MouseLeave
        t1.Clear()
    End Sub

    Private Sub btnBohrs_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBohrs.MouseEnter
        t1.Text = "5.2918 x 10^-11 m" + "  " + "the symbol is rb"
        
    End Sub

    Private Sub btnBohrs_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBohrs.MouseLeave
        t1.Clear()
    End Sub

    Private Sub btnMolar_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMolar.MouseEnter
        t1.Text = "8.3145 J/mol/K" + "  " + "the symbol is rc"
        
    End Sub

    Private Sub btnMolar_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMolar.MouseLeave
        t1.Clear()
    End Sub

    Private Sub btnAvo_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAvo.MouseEnter
        t1.Text = "6.022 x 10^23 mol^-1" + "  " + "the symbol is na"
        
    End Sub

    Private Sub btnAvo_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAvo.MouseLeave
        t1.Clear()
    End Sub

    Private Sub btnFaradays_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFaradays.MouseEnter
        t1.Text = "9.6485 x 10^4 C/mol" + "  " + "the symbol is f"
        
    End Sub

    Private Sub btnFaradays_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFaradays.MouseLeave
        t1.Clear()
    End Sub

    Private Sub FormConstants_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        btnAvo.Enabled = True
        btnMolar.Enabled = True
        btnBohrs.Enabled = True
        btnImpedence.Enabled = True
        btnPermeability.Enabled = True
        btnPermitivity.Enabled = True
        btnCharge.Enabled = True
        btnBoltzmans.Enabled = True
        btnPlanks.Enabled = True
        btnGravitational.Enabled = True
        btnCoulombs.Enabled = True
        btnSpeed.Enabled = True
        t1.Enabled = False
        CheckBox2.Checked = False
        CheckBox2.Enabled = True
        t1.Visible = True
        t2.Visible = True
        t2.Enabled = False
        btnFaradays.Enabled = True
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
        menuRule.InitialState.AddWordTransition(Nothing, "speed of light", "", SpeechGrammarWordType.SGLexical, "speed of light", 1, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Coulombs Constant", "", SpeechGrammarWordType.SGLexical, "Coulombs Constant", 2, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Gravitational Acceleration", "", SpeechGrammarWordType.SGLexical, "Gravitational Acceleration", 3, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Planks Constant", "", SpeechGrammarWordType.SGLexical, "Planks Constant", 4, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Boltzmans Constant", "", SpeechGrammarWordType.SGLexical, "Boltzmans Constant", 5, PropValue, 1.0F)

        menuRule.InitialState.AddWordTransition(Nothing, " Charge on electron", "", SpeechGrammarWordType.SGLexical, " Charge on electron", 6, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Permitivity of vacuum", "", SpeechGrammarWordType.SGLexical, "Permitivity of vacuum", 7, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Permeability of vacuum", "", SpeechGrammarWordType.SGLexical, "Permeability of vacuum", 8, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Impedence of vacuum", "", SpeechGrammarWordType.SGLexical, "Impedence of vacuum", 9, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Bohrs radius", "", SpeechGrammarWordType.SGLexical, "Bohrs radius", 10, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Molar gas constant", "", SpeechGrammarWordType.SGLexical, "Molar gas constant", 11, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Avogadros Number", "", SpeechGrammarWordType.SGLexical, "Avogadros Number", 12, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "Faradays Constant", "", SpeechGrammarWordType.SGLexical, "Faradays Constant", 13, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "voice", "", SpeechGrammarWordType.SGLexical, "voice", 1, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "", "", SpeechGrammarWordType.SGLexical, "", 1, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "", "", SpeechGrammarWordType.SGLexical, "", 1, PropValue, 1.0F)
        Try
            With voip


                .Speed = ((450 * 40) / 100)



            End With
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Reco_Event(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult)
        t2.Text = Result.PhraseInfo.GetText(0, -1, True)
        If t2.Text = "voice" Then
            If CheckBox2.Checked = True Then
                CheckBox2.Checked = False
                voice(t1.Text)
            Else
                CheckBox2.Checked = True
                voice(t1.Text)
            End If
        End If
        If t2.Text = "speed of light" Then
            t1.Text = btnSpeed.Text + "   " + "3.0 x 10^8 m/s" + "  " + "the symbol is c"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Coulombs Constant" Then
            t1.Text = btnCoulombs.Text + "   " + "8.987 x 10^9 Nm^2/Coulombs^2" + "  " + "the symbol is cc"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Gravitational Acceleration" Then
            t1.Text = btnGravitational.Text + "   " + "9.806 m/s^2" + "  " + "the symbol is g"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Planks Constant" Then
            t1.Text = btnPlanks.Text + "   " + "6.67 x 10^-11 m^3/kg/s^2" + "  " + "the symbol is h"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Boltzmans Constant" Then
            t1.Text = btnBoltzmans.Text + "   " + "1.3807 x 10^-23 J/K" + "  " + "the symbol is k"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Charge on electron" Then
            t1.Text = btnCharge.Text + "   " + "1.6022 x 10^-22 coulombs" + "  " + "the symbol is e"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Permitivity of vacuum" Then
            t1.Text = btnPermitivity.Text + "   " + "8.85 x 10^-12 F/m" + "  " + "the symbol is E"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Permeability of vacuum" Then
            t1.Text = btnPermeability.Text + "   " + "1.2566 x 10 ^ -6 N/A^2" + "  " + "the symbol is u"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Impedence of vacuum" Then
            t1.Text = btnImpedence.Text + "   " + "3.767 x 10^2 ohms" + "  " + "the symbol is z"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Bohrs radius" Then
            t1.Text = btnBohrs.Text + "   " + "5.2918 x 10^-11 m" + "  " + "the symbol is rb"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Molar gas constant" Then
            t1.Text = btnMolar.Text + "   " + "8.3145 J/mol/K" + "  " + "the symbol is rc"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Avogadros Number" Then
            t1.Text = btnAvo.Text + "   " + "6.022 x 10^23 mol^-1" + "  " + "the symbol is na"
            voice(t1.Text)
            voice(t1.Text)
        End If
        If t2.Text = "Faradays Constant" Then
            t1.Text = btnFaradays.Text + "   " + "9.6485 x 10^4 C/mol" + "  " + "the symbol is f"
            voice(t1.Text)
            voice(t1.Text)
        End If
    End Sub 'Reco_Event


    Private Sub Hypo_Event(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal Result As ISpeechRecoResult)
        t2.Text = Result.PhraseInfo.GetText(0, -1, True)
    End Sub 'Hypo_Event


    Private Sub t1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t1.TextChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = False Then
            btnAvo.Enabled = True
            btnMolar.Enabled = True
            btnBohrs.Enabled = True
            btnImpedence.Enabled = True
            btnPermeability.Enabled = True
            btnPermitivity.Enabled = True
            btnCharge.Enabled = True
            btnBoltzmans.Enabled = True
            btnPlanks.Enabled = True
            btnGravitational.Enabled = True
            btnCoulombs.Enabled = True
            btnSpeed.Enabled = True
            t1.Enabled = True
            CheckBox2.Checked = False
            CheckBox2.Enabled = True
            t1.Visible = True
            t2.Visible = True
            t2.Enabled = False
            btnFaradays.Enabled = True
        Else
            btnFaradays.Enabled = False
            btnAvo.Enabled = False
            btnMolar.Enabled = False
            btnBohrs.Enabled = False
            btnImpedence.Enabled = False
            btnPermeability.Enabled = False
            btnPermitivity.Enabled = False
            btnCharge.Enabled = False
            btnBoltzmans.Enabled = False
            btnPlanks.Enabled = False
            btnGravitational.Enabled = False
            btnCoulombs.Enabled = False
            btnSpeed.Enabled = False
            t1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox2.Enabled = False
            t1.Enabled = False
            t2.Enabled = False
        End If
    End Sub
    Sub voice(ByVal s)
        Try
            With voip
                .Speak(s)
            End With
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub

End Class