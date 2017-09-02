VERSION 5.00
Begin VB.Form frmReading 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label lblText 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmReading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()

If PageOn + 1 > 4 Then Exit Sub

PageOn = PageOn + 1
lblText.Caption = GetPageText(PageOn)

If PageOn + 1 > 4 Then cmdNext.Visible = False
If PageOn > 1 Then cmdPrevious.Visible = True

End Sub

Private Sub cmdPrevious_Click()

If PageOn - 1 < 1 Then Exit Sub

PageOn = PageOn - 1
lblText.Caption = GetPageText(PageOn)

If PageOn - 1 < 1 Then cmdPrevious.Visible = False
If PageOn < 4 Then cmdNext.Visible = True

End Sub

Private Sub Form_Load()

lblText.Caption = GetPageText(1)
cmdPrevious.Visible = False
PageOn = 1

End Sub

Public Function GetPageText(ByVal Page As Byte) As String
Dim LALA As String

Select Case Page
    Case 1
        LALA = "World War II, known as the greatest war to ever take place is officially known to have begun in 1939 and finished in 1945. The war is known to have officially begun when Germany invaded Poland in 1939. Shortly after,"
        LALA = LALA & " they invaded the countries of Luxembourg, Netherlands and Belgium in 1940. Germany was determined to attack France. France had constructed the Maginot line, a massive military structure stretching the length of the border between France and Germany. The German army moved to France through Belgium. The result of this is that German troops were able to attack British troops from behind in France and corner them in the city of Dunkirk. Winston Churchill, the president of Britain from 1940 to 1945, decided to dedicate all of Britain's efforts and resources into rescuing the British troops at Dunkirk. The evacuation of Dunkirk in 1940 was instrumental in the war as it raised the moral of Britain and saved thousands of soldiers. On April 9th, 1940, in order to secure an iron ore route, Germany easily"
        GetPageText = LALA & " conquered the countries of Denmark and Norway, adding to its massive expanding territory. Hitler and Mussolini, leading Germany and Italy respectively during the war, were allied together with the common intention of conquering everything in their path."
    Case 2
        GetPageText = "One of the many plans to retaliate at Germany by the Allies, was the planned attack on the beaches of Normandy. This operation would be very risky because the German troops would have the high ground. This brutal battle is now known as D-Day, the day the Allied troops stormed France to retake it. However, the operation was originally known as Operation Overlord."
    Case 3
        GetPageText = "On the other side of the war, Japan and the United States were also in conflict. After the bombing of Pearl Harbour, the United States retaliated against two Japanese cities who both became victims of the atomic bomb. The United States used the atomic bomb first on Hiroshima, and then on Nagasaki."
    Case 4
        GetPageText = "Looking back on World War II, approximately 70 million civilians and soldiers were killed during the 6 years of war. 104 countries were directly or indirectly involved in terms of resources, money, territory and casualties. The war struck the world and the damage leaving cities in ruins. The world has yet to see another war to match the destructive power of the Second World War."
End Select

End Function
