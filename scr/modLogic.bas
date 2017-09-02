Attribute VB_Name = "modLogic"
Public Sub LoadQuestions()
Dim i As Byte

For i = 1 To 10
    With Question(i)
        Select Case i
            Case 1
                .Question = "During what years did World War II take place?"
                .AnswerText(1) = "1939 - 1945"
                .AnswerText(2) = "1940 - 1952"
                .AnswerText(3) = "1926 - 1930"
                .AnswerText(4) = "1914 - 1918"
                .Answer = 1 'number 1 out of four
            Case 2
                .Question = "Which country did Germany invade in 1939."
                .AnswerText(1) = "Netherlands"
                .AnswerText(2) = "Poland"
                .AnswerText(3) = "Luxembourg"
                .AnswerText(4) = "Belgium"
                .Answer = 2
            Case 3
                .Question = "What is the name of the second Japanese city on which The USA used an atomic bomb?"
                .AnswerText(1) = "Hiroshima"
                .AnswerText(2) = "Kawasaki"
                .AnswerText(3) = "Nagasaki"
                .AnswerText(4) = "Kyoto"
                .Answer = 3
            Case 4
                .Question = "The attack on the beaches of Normandy, now known as D-Day, was previously referred to as what?"
                .AnswerText(1) = "Operation Market-Garden"
                .AnswerText(2) = "Operation Overlord"
                .AnswerText(3) = "Operation Normandy"
                .AnswerText(4) = "Operation Barbarossa"
                .Answer = 2
            Case 5
                .Question = "Who was the leader of Britain from 1940 to 1945?"
                .AnswerText(1) = "Neville Chamberlain"
                .AnswerText(2) = "Sir Reginald Aylmer Ranfurly Plunkett-Ernle-Erle-Drax"
                .AnswerText(3) = "Joseph Stalin"
                .AnswerText(4) = "Winston Churchill"
                .Answer = 4
            Case 6
                .Question = "From which French city did the british evacuate in mid 1940"
                .AnswerText(1) = "Paris"
                .AnswerText(2) = "Lyon"
                .AnswerText(3) = "Dunkirk"
                .AnswerText(4) = "Le Havre"
                .Answer = 3
            Case 7
                .Question = "Which two countries did Germany invade on April 9, 1940?"
                .AnswerText(1) = "Czechoslovakia and Sweden"
                .AnswerText(2) = "Norway and Denmark"
                .AnswerText(3) = "Greece and Italy"
                .AnswerText(4) = "Albania and Lithuania"
                .Answer = 2
            Case 8
                .Question = "Who was the leader of Italy during World War II"
                .AnswerText(1) = "Rizzuto"
                .AnswerText(2) = "Linguini"
                .AnswerText(3) = "Berlusconi"
                .AnswerText(4) = "Mussolini"
                .Answer = 4
            Case 9
                .Question = "Approximately how many soldiers and civilians died during World War II"
                .AnswerText(1) = "About 70 million"
                .AnswerText(2) = "About 500 thousand"
                .AnswerText(3) = "About 128 million"
                .AnswerText(4) = "About 15 million"
                .Answer = 1
            Case 10
                .Question = "How many countries were affected by World War II in terms of resources, money, and casualties."
                .AnswerText(1) = "About 77"
                .AnswerText(2) = "About 13"
                .AnswerText(3) = "About 26"
                .AnswerText(4) = "About 104"
                .Answer = 4
        End Select
    End With
Next

End Sub

Public Sub LoadQuestion(ByVal Index As Byte)

If Index = 11 Then Exit Sub

With frmMain
    .lblQuestion.Caption = Question(Index).Question
    .optAns(1).Caption = Question(Index).AnswerText(1)
    .optAns(2).Caption = Question(Index).AnswerText(2)
    .optAns(3).Caption = Question(Index).AnswerText(3)
    .optAns(4).Caption = Question(Index).AnswerText(4)
    QuestionOn = Index
End With

End Sub

Public Sub FindSelectedAnswer()
Dim CorrectAnswer As Byte
Dim i As Byte

CorrectAnswer = Question(QuestionOn).Answer

        'If QuestionOn = 2 Then MkDir (App.Path & "\data\etc\lol\")

For i = 1 To 4
    If i = CorrectAnswer And frmMain.optAns(i).Value = True Then
        Call ReadToMe("Correct! Good job!")
        Select Case Random
            Case 1
                MsgBox "YOU MUST BE A FREAKING GENIUS"
            Case 2
                MsgBox "This is clearly too easy for you."
            Case 3
                MsgBox "Do you study this in your free time?", vbYesNo
        End Select
        If QuestionOn + 1 = 11 Then frmEndGame.Show
        Call LoadQuestion(QuestionOn + 1)
        Random = CInt(Int((3 - 1 + 1) * Rnd() + 1))
        Exit Sub
    Else
        If i = 4 Then
            Call ReadToMe("Incorrect. Try again.")
            Select Case Random
                Case 1
                    MsgBox "Try again. Hopefully with a different answer.", vbCritical
                Case 2
                    MsgBox "Wrong answer. Try again.", vbCritical
                Case 3
                    MsgBox "You've almost got it.", vbCritical
            End Select
            Exit Sub
        End If
    End If
    
    Random = CInt(Int((3 - 1 + 1) * Rnd() + 1))
Next

End Sub

Public Sub ReadToMe(ByVal Text As String)
    Dim Msg, sapi
    Msg = Text
    Set sapi = CreateObject("sapi.spvoice")
    
    sapi.speak Msg
End Sub
