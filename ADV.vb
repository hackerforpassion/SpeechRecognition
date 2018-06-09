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

Public Class ADV
#Region "User Declarations"
    'Dim c As gscore10.mnulh2
    Dim my As New gscore10.audio()
    Dim down As Boolean
    Dim t As Integer
    Dim w As Integer
    Dim iscurrfile As Boolean
#End Region
#Region "My Procedures"
    Private Sub aplay()
        my.audioFile = "C:\01 Puthiya Manidha-4.wav"
        ' Label1.Text = "C:\Users\Public\Music\Sample Music\slumdog-1.wav"
        iscurrfile = True
        On Error GoTo err
        'Me.Text = "Musicplayer-" & my.audioFile
        my.audioPlay()

err:
        Exit Sub
    End Sub
#End Region
    Dim s As String = ""
    Dim c As Char
    Dim b As Char
    Dim bs As String
    Dim voip As New HTTSLib.TextToSpeechClass
    Private objRecoContext As SpeechLib.SpSharedRecoContext = Nothing
    Private grammar As SpeechLib.ISpeechRecoGrammar = Nothing
    Private menuRule As SpeechLib.ISpeechGrammarRule = Nothing
    Public Shared Function sndPlaySound(ByVal lpszSound As String, _
  ByVal flags As UInt32) As Boolean
    End Function

    Dim muc As Integer = 0
    Dim tx As Double = 0
    Dim res As Double = 0
    Dim rests As Double
    Dim pb As Integer = 0
    Dim mb As Integer = 0
    Dim mlb As Integer = 0
    Dim db As Integer = 0
    Dim pnt As Integer = 0
    Dim i As Integer
    Dim len As Integer
    Dim str As String = ""
    Dim sb As Char
    Dim lng As Integer
    Dim d As Integer
    Dim sig As Integer = 0
    'Dim dotb As Integer = 0
    Private Sub ADV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        aplay()
        Timer5.Enabled = True
        muc = 0
        With voip
            .Speed = ((450 * 40) / 100)
        End With

        Try
            'Sound file that plays on starting of the application
            sndPlaySound("C:\WINNT\Media\Windows Logon Sound.wav", New UInt32)
            'With voip
            '.Speak("Welcome To voice calculator")
            'End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'With voip
        '.Speak("i'm listening")
        'End With
        Label2.Text = "Initializing Speech Engine...."
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
        menuRule.InitialState.AddWordTransition(Nothing, "one", " ", SpeechGrammarWordType.SGLexical, "one", 100, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "two", "", SpeechGrammarWordType.SGLexical, "two", 200, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "three", "", SpeechGrammarWordType.SGLexical, "three", 300, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "four", "", SpeechGrammarWordType.SGLexical, "four", 400, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "five", "", SpeechGrammarWordType.SGLexical, "five", 500, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "six", "", SpeechGrammarWordType.SGLexical, "six", 600, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seven", "", SpeechGrammarWordType.SGLexical, "seven", 700, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eight", "", SpeechGrammarWordType.SGLexical, "eight", 800, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "nine", "", SpeechGrammarWordType.SGLexical, "nine", 900, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "zero", "", SpeechGrammarWordType.SGLexical, "zero", 1000, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "addition", "", SpeechGrammarWordType.SGLexical, "addition", 10, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "minus", "", SpeechGrammarWordType.SGLexical, "minus", 11, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "multiply", "", SpeechGrammarWordType.SGLexical, "multiply", 12, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "divide", "", SpeechGrammarWordType.SGLexical, "divide", 13, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "dot", "", SpeechGrammarWordType.SGLexical, "dot", 14, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "equals", "", SpeechGrammarWordType.SGLexical, "equals", 15, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "two hundred", "", SpeechGrammarWordType.SGLexical, "two hundred", 16, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "three hundred", "", SpeechGrammarWordType.SGLexical, "three hundred", 17, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "four hundred", "", SpeechGrammarWordType.SGLexical, "four hundred", 18, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "five hundred", "", SpeechGrammarWordType.SGLexical, "five hundred", 19, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "six hundred", "", SpeechGrammarWordType.SGLexical, "six hundred", 20, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seven hundred", "", SpeechGrammarWordType.SGLexical, "seven hundred ", 21, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eight hundred", "", SpeechGrammarWordType.SGLexical, "eight hundred", 22, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "nine hundred", "", SpeechGrammarWordType.SGLexical, "nine hundred", 23, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sine theta", "", SpeechGrammarWordType.SGLexical, "sine theta", 24, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "one thousand", "", SpeechGrammarWordType.SGLexical, "one thousand", 25, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "two thousand", "", SpeechGrammarWordType.SGLexical, "two thousand", 26, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "three thousand", "", SpeechGrammarWordType.SGLexical, "three thousand", 27, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "four thousand", "", SpeechGrammarWordType.SGLexical, "four thousand", 28, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "five thousand", "", SpeechGrammarWordType.SGLexical, "five thousand", 29, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "six thousand", "", SpeechGrammarWordType.SGLexical, "six thousand", 30, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seven thousand", "", SpeechGrammarWordType.SGLexical, "seven thousand", 31, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eight thousand", "", SpeechGrammarWordType.SGLexical, "eight thousand", 32, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "remove", "", SpeechGrammarWordType.SGLexical, "remove", 33, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "nine thousand", "", SpeechGrammarWordType.SGLexical, "nine thousand", 34, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "clear", "", SpeechGrammarWordType.SGLexical, "clear", 35, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "new", "", SpeechGrammarWordType.SGLexical, "new", 35, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "date", "", SpeechGrammarWordType.SGLexical, "date", 36, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ten thousand", "", SpeechGrammarWordType.SGLexical, "ten thousand", 37, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty thousand", "", SpeechGrammarWordType.SGLexical, "twenty thousand", 38, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty thousand", "", SpeechGrammarWordType.SGLexical, "thirty thousand", 39, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty thousand", "", SpeechGrammarWordType.SGLexical, "fourty thousand", 40, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty thousand", "", SpeechGrammarWordType.SGLexical, "fifty thousand", 41, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty thousand", "", SpeechGrammarWordType.SGLexical, "sixty thousand", 42, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy thousand", "", SpeechGrammarWordType.SGLexical, "seventy thousand", 43, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty thousand", "", SpeechGrammarWordType.SGLexical, "eighty thousand", 44, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety thousand", "", SpeechGrammarWordType.SGLexical, "ninety thousand", 45, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "otherfunctions", "", SpeechGrammarWordType.SGLexical, "otherfunctions", 46, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "constants values", "", SpeechGrammarWordType.SGLexical, "constants values", 47, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "memory", "", SpeechGrammarWordType.SGLexical, "memory", 48, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "square", "", SpeechGrammarWordType.SGLexical, "square", 49, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "and", "", SpeechGrammarWordType.SGLexical, "and", 50, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "or", "", SpeechGrammarWordType.SGLexical, "or", 51, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "not", "", SpeechGrammarWordType.SGLexical, "not", 52, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "exor", "", SpeechGrammarWordType.SGLexical, "exor", 53, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "power", "", SpeechGrammarWordType.SGLexical, "power", 54, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sign", "", SpeechGrammarWordType.SGLexical, "sign", 55, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "cube", "", SpeechGrammarWordType.SGLexical, "cube", 56, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "factorial", "", SpeechGrammarWordType.SGLexical, "factorial", 57, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "squareroot", "", SpeechGrammarWordType.SGLexical, "squareroot", 58, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "round", "", SpeechGrammarWordType.SGLexical, "round", 59, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "ceil", "", SpeechGrammarWordType.SGLexical, "ceil", 60, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "permutation", "", SpeechGrammarWordType.SGLexical, "permutation", 61, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "combination", "", SpeechGrammarWordType.SGLexical, "combination", 62, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "inverse", "", SpeechGrammarWordType.SGLexical, "inverse", 63, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "pai", "", SpeechGrammarWordType.SGLexical, "pai", 64, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "", "", SpeechGrammarWordType.SGLexical, "", 65, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "billion", "", SpeechGrammarWordType.SGLexical, "billion", 65, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "million", "", SpeechGrammarWordType.SGLexical, "million", 66, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thousand", "", SpeechGrammarWordType.SGLexical, "thousand", 67, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "hundred", "", SpeechGrammarWordType.SGLexical, "hundred", 68, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ten", "", SpeechGrammarWordType.SGLexical, "ten", 69, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty", "", SpeechGrammarWordType.SGLexical, "twenty", 70, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty", "", SpeechGrammarWordType.SGLexical, "thirty", 71, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty", "", SpeechGrammarWordType.SGLexical, "fourty", 72, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty", "", SpeechGrammarWordType.SGLexical, "fifty", 73, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty", "", SpeechGrammarWordType.SGLexical, "sixty", 74, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy", "", SpeechGrammarWordType.SGLexical, "seventy", 75, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty", "", SpeechGrammarWordType.SGLexical, "eighty", 76, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety", "", SpeechGrammarWordType.SGLexical, "ninety", 77, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eleven", "", SpeechGrammarWordType.SGLexical, "eleven", 78, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twelve", "", SpeechGrammarWordType.SGLexical, "twelve", 79, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirteen", "", SpeechGrammarWordType.SGLexical, "thirteen", 80, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourteen", "", SpeechGrammarWordType.SGLexical, "fourteen", 81, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifteen", "", SpeechGrammarWordType.SGLexical, "fifteen", 82, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixteen", "", SpeechGrammarWordType.SGLexical, "sixteen", 83, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventeen", "", SpeechGrammarWordType.SGLexical, "seventeen", 84, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighteen", "", SpeechGrammarWordType.SGLexical, "eighteen", 85, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "nineteen", "", SpeechGrammarWordType.SGLexical, "nineteen", 86, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "text to speech", "", SpeechGrammarWordType.SGLexical, "text to speech", 87, PropValue, 1.0F)
        'menuRule.InitialState.AddWordTransition(Nothing, "string to text", "", SpeechGrammarWordType.SGLexical, "string to text", 88, PropValue, 1.0F)
        '  menuRule.InitialState.AddWordTransition(Nothing, "equals", "", SpeechGrammarWordType.SGLexical, "equals", 89, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "grey", "", SpeechGrammarWordType.SGLexical, "grey", 90, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "blue", "", SpeechGrammarWordType.SGLexical, "blue", 91, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "violet", "", SpeechGrammarWordType.SGLexical, "violet", 92, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "speechlevel", "", SpeechGrammarWordType.SGLexical, "speechlevel", 93, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "point", "", SpeechGrammarWordType.SGLexical, "point", 94, PropValue, 1.0F)
        ' menuRule.InitialState.AddWordTransition(Nothing, "basic", " ", SpeechGrammarWordType.SGLexical, "basic", 101, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty one", "", SpeechGrammarWordType.SGLexical, "twenty one", 201, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty two", "", SpeechGrammarWordType.SGLexical, "twenty two", 202, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty three", "", SpeechGrammarWordType.SGLexical, "twenty three", 203, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty four", "", SpeechGrammarWordType.SGLexical, "twenty four", 204, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty five", "", SpeechGrammarWordType.SGLexical, "twenty five", 205, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty six", "", SpeechGrammarWordType.SGLexical, "twenty six", 206, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty seven", "", SpeechGrammarWordType.SGLexical, "twenty seven", 207, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty eight", "", SpeechGrammarWordType.SGLexical, "twenty eight", 208, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "twenty nine", "", SpeechGrammarWordType.SGLexical, "twenty nine", 209, PropValue, 1.0F)

        menuRule.InitialState.AddWordTransition(Nothing, "thirty one", "", SpeechGrammarWordType.SGLexical, "thirty one", 301, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty two", "", SpeechGrammarWordType.SGLexical, "thirty two", 302, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty three", "", SpeechGrammarWordType.SGLexical, "thirty three", 303, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty four", "", SpeechGrammarWordType.SGLexical, "thirty four", 304, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty five", "", SpeechGrammarWordType.SGLexical, "thirty five", 305, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty six", "", SpeechGrammarWordType.SGLexical, "thirty six", 306, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty seven", "", SpeechGrammarWordType.SGLexical, "thirty seven", 307, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty eight", "", SpeechGrammarWordType.SGLexical, "thirty eight", 308, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "thirty nine", "", SpeechGrammarWordType.SGLexical, "thirty nine", 309, PropValue, 1.0F)

        menuRule.InitialState.AddWordTransition(Nothing, "fourty one", "", SpeechGrammarWordType.SGLexical, "fourty one", 401, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty two", "", SpeechGrammarWordType.SGLexical, "fourty two", 402, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty three", "", SpeechGrammarWordType.SGLexical, "fourty three", 403, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty four", "", SpeechGrammarWordType.SGLexical, "fourty four", 404, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty five", "", SpeechGrammarWordType.SGLexical, "fourty five", 405, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty six", "", SpeechGrammarWordType.SGLexical, "fourty six", 406, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty seven", "", SpeechGrammarWordType.SGLexical, "fourty seven", 407, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty eight", "", SpeechGrammarWordType.SGLexical, "fourty eight", 408, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fourty nine", "", SpeechGrammarWordType.SGLexical, "fourty nine", 409, PropValue, 1.0F)

        menuRule.InitialState.AddWordTransition(Nothing, "fifty one", "", SpeechGrammarWordType.SGLexical, "fifty one", 501, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty two", "", SpeechGrammarWordType.SGLexical, "fifty two", 502, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty three", "", SpeechGrammarWordType.SGLexical, "fifty three", 503, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty four", "", SpeechGrammarWordType.SGLexical, "fifty four", 504, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty five", "", SpeechGrammarWordType.SGLexical, "fifty five", 505, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty six", "", SpeechGrammarWordType.SGLexical, "fifty six", 506, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty seven", "", SpeechGrammarWordType.SGLexical, "fifty seven", 507, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty eight", "", SpeechGrammarWordType.SGLexical, "fifty eight", 508, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "fifty nine", "", SpeechGrammarWordType.SGLexical, "fifty nine", 509, PropValue, 1.0F)

        menuRule.InitialState.AddWordTransition(Nothing, "sixty one", "", SpeechGrammarWordType.SGLexical, "sixty one", 601, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty two", "", SpeechGrammarWordType.SGLexical, "sixty two", 602, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty three", "", SpeechGrammarWordType.SGLexical, "sixty three", 603, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty four", "", SpeechGrammarWordType.SGLexical, "sixty four", 604, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty five", "", SpeechGrammarWordType.SGLexical, "sixty five", 605, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty six", "", SpeechGrammarWordType.SGLexical, "sixty six", 606, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty seven", "", SpeechGrammarWordType.SGLexical, "sixty seven", 607, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty eight", "", SpeechGrammarWordType.SGLexical, "sixty eight", 608, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "sixty nine", "", SpeechGrammarWordType.SGLexical, "sixty nine", 609, PropValue, 1.0F)

        menuRule.InitialState.AddWordTransition(Nothing, "seventy one", "", SpeechGrammarWordType.SGLexical, "seventy one", 701, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy two", "", SpeechGrammarWordType.SGLexical, "seventy two", 702, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy three", "", SpeechGrammarWordType.SGLexical, "seventy three", 703, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy four", "", SpeechGrammarWordType.SGLexical, "seventy four", 704, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy five", "", SpeechGrammarWordType.SGLexical, "seventy five", 705, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy six", "", SpeechGrammarWordType.SGLexical, "seventy six", 706, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy seven", "", SpeechGrammarWordType.SGLexical, "seventy seven", 707, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy eight", "", SpeechGrammarWordType.SGLexical, "seventy eight", 708, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "seventy nine", "", SpeechGrammarWordType.SGLexical, "seventy nine", 709, PropValue, 1.0F)


        menuRule.InitialState.AddWordTransition(Nothing, "eighty one", "", SpeechGrammarWordType.SGLexical, "eighty one", 801, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty two", "", SpeechGrammarWordType.SGLexical, "eighty two", 802, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty three", "", SpeechGrammarWordType.SGLexical, "eighty three", 803, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty four", "", SpeechGrammarWordType.SGLexical, "eighty four", 804, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty five", "", SpeechGrammarWordType.SGLexical, "eighty five", 805, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty six", "", SpeechGrammarWordType.SGLexical, "eighty six", 806, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty seven", "", SpeechGrammarWordType.SGLexical, "eighty seven", 807, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty eight", "", SpeechGrammarWordType.SGLexical, "eighty eight", 808, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "eighty nine", "", SpeechGrammarWordType.SGLexical, "eighty nine", 809, PropValue, 1.0F)


        menuRule.InitialState.AddWordTransition(Nothing, "ninety one", "", SpeechGrammarWordType.SGLexical, "ninety one", 801, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety two", "", SpeechGrammarWordType.SGLexical, "ninety two", 802, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety three", "", SpeechGrammarWordType.SGLexical, "ninety three", 803, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety four", "", SpeechGrammarWordType.SGLexical, "ninety four", 804, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety five", "", SpeechGrammarWordType.SGLexical, "ninety five", 805, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety six", "", SpeechGrammarWordType.SGLexical, "ninety six", 806, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety seven", "", SpeechGrammarWordType.SGLexical, "ninety seven", 807, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety eight", "", SpeechGrammarWordType.SGLexical, "ninety eight", 808, PropValue, 1.0F)
        menuRule.InitialState.AddWordTransition(Nothing, "ninety nine", "", SpeechGrammarWordType.SGLexical, "ninety nine", 809, PropValue, 1.0F)
        grammar.Rules.Commit()
        grammar.CmdSetRuleState("MenuCommands", SpeechRuleState.SGDSActive)
        Label2.Text = "Speech Engine Ready for Input"
    End Sub
    Private Sub Reco_Event(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult)
        TextBox1.Text = Result.PhraseInfo.GetText(0, -1, True)
        If Result.PhraseInfo.GetText(0, -1, True) <> "new" Or Result.PhraseInfo.GetText(0, -1, True) <> "sign" Or Result.PhraseInfo.GetText(0, -1, True) <> "remove" Or Result.PhraseInfo.GetText(0, -1, True) <> "equals" Or Result.PhraseInfo.GetText(0, -1, True) <> "plus" Or Result.PhraseInfo.GetText(0, -1, True) <> "minus" Or Result.PhraseInfo.GetText(0, -1, True) <> "multiply" Or Result.PhraseInfo.GetText(0, -1, True) <> "divide" Then
            Txt.Text = Txt.Text + Result.PhraseInfo.GetText(0, -1, True) + " "
        End If
        If Result.PhraseInfo.GetText(0, -1, True) = "remove" Then
            bck()
        End If

        If Result.PhraseInfo.GetText(0, -1, True) = "sign" Then
            If pb = 1 Or mb = 1 Or mlb = 1 Or db = 1 Then
                If tx1.Text = "" Then
                    If sig = 0 Then
                        sig = 1
                        tx1.Text = CStr("-")
                    Else
                        sig = 0
                        tx1.Text = CStr("")
                    End If
                Else
                    tx1.Text = Val(tx1.Text) * -1
                End If
            Else
                If rest.Text = "" Then
                    If sig = 0 Then
                        sig = 2
                        rest.Text = CStr("-")
                    Else

                        sig = 0
                        rest.Text = CStr("")
                    End If
                Else
                    rest.Text = Val(rest.Text) * -1
                End If
            End If
        End If
        If Result.PhraseInfo.GetText(0, -1, True) = "new" Then
            fr.Clear()
            ans.Clear()
            Txt.Clear()
            tx1.Clear()
            rest.Clear()
            tx = 0
        End If

        If Result.PhraseInfo.GetText(0, -1, True) = "addition" Then
            pnt = 0
            tx = 0
            ans.Clear()
            pb = 1
            s = Txt.Text
            Txt.Clear()
            help()
            tx1.Text = CStr(Val(tx1.Text) + tx)
            'rest.Text = Val(tx1.Text)
            rest.Text = tx1.Text
            tx1.Clear()
            tx = 0
            If sig = 2 Then
                rest.Text = Val(rest.Text) * -1
                sig = 0
            End If
        End If
        If Result.PhraseInfo.GetText(0, -1, True) = "minus" Then
            pnt = 0
            tx = 0
            ans.Clear()
            mb = 1
            s = Txt.Text
            Txt.Clear()
            help()
            tx1.Text = CStr(Val(tx1.Text) + tx)
            'rest.Text = Val(tx1.Text)
            rest.Text = tx1.Text
            tx1.Clear()
            tx = 0
            If sig = 2 Then
                rest.Text = Val(rest.Text) * -1
                sig = 0
            End If
        End If
        If Result.PhraseInfo.GetText(0, -1, True) = "multiply" Then
            tx = 0
            pnt = 0
            ans.Clear()
            mlb = 1
            s = Txt.Text
            Txt.Clear()
            help()
            tx1.Text = CStr(Val(tx1.Text) + tx)
            'rest.Text = Val(tx1.Text)
            rest.Text = tx1.Text
            tx1.Clear()
            tx = 0
            If sig = 2 Then
                rest.Text = Val(rest.Text) * -1
                sig = 0
            End If
        End If
        If Result.PhraseInfo.GetText(0, -1, True) = "divide" Then
            tx = 0
            ans.Clear()
            db = 1
            s = Txt.Text
            Txt.Clear()
            help()
            tx1.Text = CStr(Val(tx1.Text) + tx)
            'rest.Text = Val(tx1.Text)
            rest.Text = tx1.Text
            tx1.Clear()
            tx = 0
            If sig = 2 Then
                rest.Text = Val(rest.Text) * -1
                sig = 0
            End If
        End If
        If Result.PhraseInfo.GetText(0, -1, True) = "equals" Then
            pnt = 0
            ans.Clear()
            s = Txt.Text
            Txt.Clear()
            help()
            tx1.Text = CStr(Val(tx1.Text) + tx)
            tx = 0
            If sig = 1 Then
                tx1.Text = Val(tx1.Text) * -1
                sig = 0
            End If
            If sig = 2 Then
                rest.Text = Val(rest.Text) * -1
                sig = 0
            End If
            If rest.Text <> "" Then
                If pb = 1 Then
                    Try
                        With voip
                            .Speak((rest.Text) + "plus" + (tx1.Text) + "Equals")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    fr.Text = CStr(Val(rest.Text) + Val(tx1.Text))
                    Try
                        With voip
                            .Speak(Val(fr.Text))
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    pb = 0
                End If
                If mb = 1 Then
                    Try
                        With voip
                            .Speak((rest.Text) + "minus" + (tx1.Text) + "Equals")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    fr.Text = CStr(Val(rest.Text) - Val(tx1.Text))
                    Try
                        With voip
                            .Speak(Val(fr.Text))
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    mb = 0
                End If
                If mlb = 1 Then
                    Try
                        With voip
                            .Speak((rest.Text) + "into" + (tx1.Text) + "Equals")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    fr.Text = CStr(Val(rest.Text) * Val(tx1.Text))
                    Try
                        With voip
                            .Speak(Val(fr.Text))
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    mlb = 0
                End If
                If db = 1 Then
                    Try
                        With voip
                            .Speak((rest.Text) + "divide at by" + (tx1.Text) + "Equals")
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    fr.Text = CStr(Val(rest.Text) / Val(tx1.Text))
                    Try
                        With voip
                            .Speak(Val(fr.Text))
                        End With
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                    db = 0
                End If
            Else
                fr.Text = tx1.Text
                Try
                    With voip
                        .Speak(Val(tx1.Text))
                    End With
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
            End If
            ans.Clear()
        End If
    End Sub
    Private Sub Hypo_Event(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal Result As ISpeechRecoResult)
        If Result.PhraseInfo.GetText(0, -1, True) <> "clear" Or Result.PhraseInfo.GetText(0, -1, True) <> "remove" Then
            'Txt.Text = Txt.Text + Result.PhraseInfo.GetText(0, -1, True) + " "
        End If
    End Sub 'Hypo_Event

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        fr.Clear()
        ans.Clear()
        Txt.Clear()
        tx1.Clear()
        rest.Clear()
        tx = 0
    End Sub
    Sub help()
        For Each c In s
            If ans.Text <> "one" Or ans.Text <> "two" Or ans.Text <> "three" Or ans.Text <> "four" Or ans.Text <> "five" Or ans.Text <> "six" Or ans.Text <> "seven" Or ans.Text <> "eight" Or ans.Text <> "nine" Or ans.Text <> "ten" Or ans.Text <> "twenty" Or ans.Text <> "thirty" Or ans.Text <> "fourty" Or ans.Text <> "fifty" Or ans.Text <> "sixty" Or ans.Text <> "seventy" Or ans.Text <> "eighty" Or ans.Text <> "ninety" Or ans.Text <> "eleven" Or ans.Text <> "twelve" Or ans.Text <> "thirteen" Or ans.Text <> "fourteen" Or ans.Text <> "fifteen" Or ans.Text <> "sixteen" Or ans.Text <> "seventeen" Or ans.Text <> "eighteen" Or ans.Text <> "nineteen" Or ans.Text <> "hundred" Or ans.Text <> "thousand" Or ans.Text <> "million" Or ans.Text <> "billion" Or ans.Text <> "point" Then
                If c <> " " Then
                    ans.Text = ans.Text + CStr(c)
                End If
            End If
            'If ans.Text = "one" Or ans.Text = "two" Or ans.Text = "three" Or ans.Text = "four" Or ans.Text = "five" Or ans.Text = "six" Or ans.Text = "seven" Or ans.Text = "eight" Or ans.Text = "nine" Or ans.Text = "ten" Or ans.Text = "twenty" Or ans.Text = "thirty" Or ans.Text = "fourty" Or ans.Text = "fifty" Or ans.Text = "sixty" Or ans.Text = "seventy" Or ans.Text = "eighty" Or ans.Text = "ninety" Or ans.Text <> "eleven" Or ans.Text <> "twelve" Or ans.Text <> "thirteen" Or ans.Text <> "fourteen" Or ans.Text <> "fifteen" Or ans.Text <> "sixteen" Or ans.Text <> "seventeen" Or ans.Text <> "eighteen" Or ans.Text <> "nineteen" Then
            If ans.Text = "one" Then
                If pnt <> 1 Then
                    tx = tx + 1
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(1)
                    ans.Clear()
                End If
            ElseIf ans.Text = "two" Then
                If pnt <> 1 Then
                    tx = tx + 2
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(2)
                    ans.Clear()
                End If

            ElseIf ans.Text = "three" Then
                If pnt <> 1 Then
                    tx = tx + 3
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(3)
                    ans.Clear()
                End If
            ElseIf ans.Text = "four" Then
                If pnt <> 1 Then
                    tx = tx + 4
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(4)
                    ans.Clear()
                End If

            ElseIf ans.Text = "five" Then

                If pnt <> 1 Then
                    tx = tx + 5
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(5)
                    ans.Clear()
                End If


            ElseIf ans.Text = "six" Then

                If pnt <> 1 Then
                    tx = tx + 6
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(6)
                    ans.Clear()
                End If

            ElseIf ans.Text = "seven" Then

                If pnt <> 1 Then
                    tx = tx + 7
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(7)
                    ans.Clear()
                End If
            ElseIf ans.Text = "eight" Then

                If pnt <> 1 Then
                    tx = tx + 8
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(8)
                    ans.Clear()
                End If

            ElseIf ans.Text = "nine" Then
                If pnt <> 1 Then
                    tx = tx + 9
                    ans.Clear()
                Else
                    tx1.Text = tx1.Text + CStr(9)
                    ans.Clear()
                End If

            ElseIf ans.Text = "ten" Then
                tx = tx + 10
                ans.Clear()
            ElseIf ans.Text = "eleven" Then
                tx = tx + 11
                ans.Clear()
            ElseIf ans.Text = "twelve" Then
                tx = tx + 12
                ans.Clear()
            ElseIf ans.Text = "thirteen" Then
                tx = tx + 13
                ans.Clear()
            ElseIf ans.Text = "fourteen" Then
                tx = tx + 14
                ans.Clear()
            ElseIf ans.Text = "fifteen" Then
                tx = tx + 15
                ans.Clear()
            ElseIf ans.Text = "sixteen" Then
                tx = tx + 16
                ans.Clear()
            ElseIf ans.Text = "seventeen" Then
                tx = tx + 17
                ans.Clear()
            ElseIf ans.Text = "eighteen" Then
                tx = tx + 18
                ans.Clear()
            ElseIf ans.Text = "nineteen" Then
                tx = tx + 19
                ans.Clear()
            ElseIf ans.Text = "twenty" Then
                tx = tx + 20
                ans.Clear()
            ElseIf ans.Text = "thirty" Then
                tx = tx + 30
                ans.Clear()
            ElseIf ans.Text = "fourty" Then
                tx = tx + 40
                ans.Clear()
            ElseIf ans.Text = "fifty" Then
                tx = tx + 50
                ans.Clear()
            ElseIf ans.Text = "sixty" Then
                tx = tx + 60
                ans.Clear()
            ElseIf ans.Text = "seventy" Then
                tx = tx + 70
                ans.Clear()
            ElseIf ans.Text = "eighty" Then
                tx = tx + 80
                ans.Clear()
            ElseIf ans.Text = "ninety" Then
                tx = tx + 90

                ans.Clear()
                ' ElseIf ans.Text = "hundred" Or ans.Text = "thousand" Or ans.Text = "million" Or ans.Text = "billion" Then
            ElseIf ans.Text = "hundred" Then
                If tx < 10 And tx > 0 Then
                    tx = tx * 100
                    'tx1.Text = tx1.Text + CStr(tx)
                    ans.Clear()
                End If
            ElseIf ans.Text = "thousand" Then
                If tx < 1000 And tx > 0 Then
                    tx = tx * 1000
                    tx1.Text = CStr(Val(tx1.Text) + tx)
                    tx = 0
                    ans.Clear()
                End If
            ElseIf ans.Text = "million" Then
                If tx < 1000 And tx > 0 Then
                    tx = tx * 1000000
                    tx1.Text = CStr(Val(tx1.Text) + tx)
                    tx = 0
                    ans.Clear()
                End If
            ElseIf ans.Text = "billion" Then
                tx1.Text = CStr(Val(tx1.Text) + tx)
                res = Val(tx1.Text) * 1000000000
                tx1.Text = CStr(res)
                tx = 0
                ans.Clear()
            ElseIf ans.Text = "point" Then
                tx1.Text = Val(tx1.Text) + tx
                tx1.Text = tx1.Text + CStr(".")
                tx = 0
                ans.Clear()
                pnt = 1

            End If
            'End If
        Next
    End Sub
    Sub bck()
        src.Clear()
        Dim hp As String
        Dim h As Char
        hp = Txt.Text
        bsp.Clear()
        For Each h In hp
            If bsp.Text = "one" Or bsp.Text = "two" Or bsp.Text = "three" Or bsp.Text = "four" Or bsp.Text = "five" Or bsp.Text = "six" Or bsp.Text = "seven" Or bsp.Text = "eight" Or bsp.Text = "nine" Or bsp.Text = "ten" Or bsp.Text = "twenty" Or bsp.Text = "thirty" Or bsp.Text = "fourty" Or bsp.Text = "fifty" Or bsp.Text = "sixty" Or bsp.Text = "seventy" Or bsp.Text = "eighty" Or bsp.Text = "ninety" Or bsp.Text = "eleven" Or bsp.Text = "twelve" Or bsp.Text = "thirteen" Or bsp.Text = "fourteen" Or bsp.Text = "fifteen" Or bsp.Text = "sixteen" Or bsp.Text = "seventeen" Or bsp.Text = "eighteen" Or bsp.Text = "nineteen" Or bsp.Text = "hundred" Or bsp.Text = "thousand" Or bsp.Text = "million" Or bsp.Text = "billion" Or bsp.Text = "point" Then
                src.Text = src.Text + bsp.Text + " "
                bsp.Clear()
            End If
            If bsp.Text <> "one" Or bsp.Text <> "two" Or bsp.Text <> "three" Or bsp.Text <> "four" Or bsp.Text <> "five" Or bsp.Text <> "six" Or bsp.Text <> "seven" Or bsp.Text <> "eight" Or bsp.Text <> "nine" Or bsp.Text <> "ten" Or bsp.Text <> "twenty" Or bsp.Text <> "thirty" Or bsp.Text <> "fourty" Or bsp.Text <> "fifty" Or bsp.Text <> "sixty" Or bsp.Text <> "seventy" Or bsp.Text <> "eighty" Or bsp.Text <> "ninety" Or bsp.Text <> "eleven" Or bsp.Text <> "twelve" Or bsp.Text <> "thirteen" Or bsp.Text <> "fourteen" Or bsp.Text <> "fifteen" Or bsp.Text <> "sixteen" Or bsp.Text <> "seventeen" Or bsp.Text <> "eighteen" Or bsp.Text <> "nineteen" Or bsp.Text <> "hundred" Or bsp.Text <> "thousand" Or bsp.Text <> "million" Or bsp.Text <> "billion" Or bsp.Text <> "point" Then
                If h <> " " Then
                    bsp.Text = bsp.Text + CStr(h)
                End If
            End If
        Next
        Txt.Text = src.Text
        src.Clear()
        bs = Txt.Text
        bsp.Clear()
        For Each b In bs
            If bsp.Text <> "one" Or bsp.Text <> "two" Or bsp.Text <> "three" Or bsp.Text <> "four" Or bsp.Text <> "five" Or bsp.Text <> "six" Or bsp.Text <> "seven" Or bsp.Text <> "eight" Or bsp.Text <> "nine" Or bsp.Text <> "ten" Or bsp.Text <> "twenty" Or bsp.Text <> "thirty" Or bsp.Text <> "fourty" Or bsp.Text <> "fifty" Or bsp.Text <> "sixty" Or bsp.Text <> "seventy" Or bsp.Text <> "eighty" Or bsp.Text <> "ninety" Or bsp.Text <> "eleven" Or bsp.Text <> "twelve" Or bsp.Text <> "thirteen" Or bsp.Text <> "fourteen" Or bsp.Text <> "fifteen" Or bsp.Text <> "sixteen" Or bsp.Text <> "seventeen" Or bsp.Text <> "eighteen" Or bsp.Text <> "nineteen" Or bsp.Text <> "hundred" Or bsp.Text <> "thousand" Or bsp.Text <> "million" Or bsp.Text <> "billion" Or bsp.Text <> "point" Then
                If b <> " " Then
                    bsp.Text = bsp.Text + CStr(b)
                End If
            End If
            If bsp.Text = "one" Or bsp.Text = "two" Or bsp.Text = "three" Or bsp.Text = "four" Or bsp.Text = "five" Or bsp.Text = "six" Or bsp.Text = "seven" Or bsp.Text = "eight" Or bsp.Text = "nine" Or bsp.Text = "ten" Or bsp.Text = "twenty" Or bsp.Text = "thirty" Or bsp.Text = "fourty" Or bsp.Text = "fifty" Or bsp.Text = "sixty" Or bsp.Text = "seventy" Or bsp.Text = "eighty" Or bsp.Text = "ninety" Or bsp.Text = "eleven" Or bsp.Text = "twelve" Or bsp.Text = "thirteen" Or bsp.Text = "fourteen" Or bsp.Text = "fifteen" Or bsp.Text = "sixteen" Or bsp.Text = "seventeen" Or bsp.Text = "eighteen" Or bsp.Text = "nineteen" Or bsp.Text = "hundred" Or bsp.Text = "thousand" Or bsp.Text = "million" Or bsp.Text = "billion" Or bsp.Text = "point" Then
                hlp.Text = bsp.Text
                bsp.Clear()
            End If
        Next
        bsp.Text = Txt.Text
        Txt.Clear()
        src.Clear()
        str = bsp.Text
        'hlp.Text = CStr(src)
        lng = hlp.TextLength + 1
        len = bsp.TextLength
        bsp.Clear()
        hlp.Clear()
        d = len - lng
        i = 1
        For Each sb In str
            If i <= d Then
                Txt.Text = Txt.Text + CStr(sb)
                i = i + 1
            End If
        Next
        lng = 0
        len = 0
        str = ""
        i = 0
        bs = ""
        d = 0
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        With voip
            .Speak(TextBox1.Text)
        End With
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        muc = muc + 1
        If muc = 9 Then
            With voip
                .Speak("Advanced voice calculator")
            End With
            With voip
                .Speak("i'm listening")
            End With
            muc = 0
            Timer5.Enabled = False
        End If


    End Sub
End Class