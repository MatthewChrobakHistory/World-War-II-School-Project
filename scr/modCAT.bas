Attribute VB_Name = "modCAT"
Option Explicit

Public QuestionOn As Byte
Public Random As Byte
Public PageOn As Byte

Public Question(1 To 10) As QuestionRec

Private Type QuestionRec
    Question As String
    AnswerText(1 To 4) As String
    Answer As Byte
End Type
