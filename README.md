# SpeechRecognition
This application is based on Microsoft's speech SDK

Voice Calculator is an application which implements Windows Speech API . 
This is a simple application about 5000 lines of code to control a normal and scientific calculator using the Windows Speech API(here the calculator is a customized application). 
I did this project in my early days of programming . 
It took a little bit of time for me to under stand the lexical analysis and speech synthesis process of the Speech API from Microsoft.


**Requirements:**

1) Windows XP or Windows 7 PC

2) A microphone and a speaker

3) Microsoft Speech SDK 5.1 (it is free to download)

4) Microsoft Speech API 

5) Microsoft Visual Studio 2005

We need to tell the key words to the application which are going to be used to control the application using the obeject of following grammer Object

**grammar.Rules.Add("MenuCommands", SpeechRuleAttributes.SRATopLevel Or SpeechRuleAttributes.SRADynamic, 1)**
grammer is created from SpeechLib.ISpeechRecoGrammar

Most of the code is self explanatory.

We have an advanced application which can get the data in the form of string like one , hundred , twenty , five (without commas , a little time gap is considered as a complete word here)

and convert them to numbers and commands, for example

we can say **add fifty five  , twenty two**

The speech recognition system works based on the user's voice input and the precision relates to the clarity of voice and amount of training given to the speech engine.

**Steps to configure and train the speech engine:**
1) Install Windows Speech SDK 5.1 and Speech API 4

2) Launch Speech recognition from start -> Speech recognition

3) Follow the wizard and complete the speech training (this is for precision).

The speech recognition system works based on the level of training and the amount of training given to it. It learns what we teach . So that it can recognize our slang.


**Working Model:**
![Training](https://lh4.googleusercontent.com/RXe7o1Lbre0Abw_Y5N3Z1K6vZFw2SarccbeRDT6jlWssuC0KuDwJZNfZ9I8fEREno-SS-LIm-axfjw=w1853-h951)
![Working Model](https://lh5.googleusercontent.com/WO3ZHhrddaZHXaDr-u6AyLADbsNDaaEVZymikY8lYsSogL6Vr4FxLtVEvH63GIIU1nUHZqTZOrVVAw=w1853-h951)
