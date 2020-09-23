VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collision Generator & Detector"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblGen 
      Caption         =   "Click This Label To Generate Collision"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Line Line13 
      X1              =   8
      X2              =   48
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Line Line12 
      X1              =   8
      X2              =   48
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line Line11 
      X1              =   88
      X2              =   88
      Y1              =   56
      Y2              =   144
   End
   Begin VB.Line Line10 
      X1              =   128
      X2              =   128
      Y1              =   112
      Y2              =   192
   End
   Begin VB.Line Line9 
      X1              =   8
      X2              =   8
      Y1              =   56
      Y2              =   192
   End
   Begin VB.Line Line8 
      X1              =   296
      X2              =   8
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Line Line7 
      X1              =   296
      X2              =   296
      Y1              =   56
      Y2              =   192
   End
   Begin VB.Line Line6 
      X1              =   8
      X2              =   296
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Line Line5 
      X1              =   160
      X2              =   240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      X1              =   160
      X2              =   160
      Y1              =   72
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   160
      X2              =   280
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Line Line2 
      X1              =   280
      X2              =   280
      Y1              =   72
      Y2              =   152
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2640
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   160
      X2              =   280
      Y1              =   72
      Y2              =   72
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'This sub checks to what side are trying to move, checks if there will be a collision if you move and lets you, or doesnt let you move
'Note that you cant move in two directions at a time

If CollisionGenerated = False Then Exit Sub 'Doesnt allow you to move untill you generate collision

Select Case KeyCode

Case vbKeyLeft 'If you pressed Left
If CheckCollision(Left, Shape.Left, Shape.Top, Shape.Height, Shape.Width) = True Then Exit Sub 'Calls the CheckCollision function with specified parameters
Shape.Left = Shape.Left - 1 'Moves you by a pixel in a direction

Case vbKeyRight 'If you pressed Right
If CheckCollision(Right, Shape.Left, Shape.Top, Shape.Height, Shape.Width) = True Then Exit Sub 'Calls the CheckCollision function with specified parameters
Shape.Left = Shape.Left + 1 'Moves you by a pixel in a direction

Case vbKeyUp 'If you pressed Up
If CheckCollision(Up, Shape.Left, Shape.Top, Shape.Height, Shape.Width) = True Then Exit Sub 'Calls the CheckCollision function with specified parameters
Shape.Top = Shape.Top - 1 'Moves you by a pixel in a direction

Case vbKeyDown 'If you pressed Down
If CheckCollision(Down, Shape.Left, Shape.Top, Shape.Height, Shape.Width) = True Then Exit Sub 'Calls the CheckCollision function with specified parameters
Shape.Top = Shape.Top + 1 'Moves you by a pixel in a direction

End Select

End Sub

Private Function CheckCollision(Side As Side, ObjectLeft As Long, ObjectTop As Long, ObjectHeight As Long, ObjectWidth As Long) As Boolean
'This function is explained in the documentation

Select Case Side

'--------------------------------------------------------------------------------------
Case Up 'Used to checks collision above the object

    If InStr(1, Mid(CollisionArray(ObjectTop - 1), ObjectLeft + 1, ObjectWidth), "1") <> 0 Then CheckCollision = True

'--------------------------------------------------------------------------------------
Case Down 'Used to checks collision under the object

    If InStr(1, Mid(CollisionArray(ObjectTop + ObjectHeight), ObjectLeft + 1, ObjectWidth), "1") <> 0 Then CheckCollision = True

'--------------------------------------------------------------------------------------
Case Left 'Used to checks collision left from object

    For ArrayPosition = ObjectTop To ObjectTop + ObjectHeight - 1
        If Mid(CollisionArray(ArrayPosition), ObjectLeft, 1) = 1 Then
        CheckCollision = True
        GoTo FoundCollision
        End If
    Next

'--------------------------------------------------------------------------------------
Case Right 'Used to checks collision right from object

    For ArrayPosition = ObjectTop To ObjectTop + ObjectHeight - 1
        If Mid(CollisionArray(ArrayPosition), ObjectLeft + ObjectWidth + 1, 1) = 1 Then
        CheckCollision = True
        GoTo FoundCollision
        End If
    Next
    
    
End Select

FoundCollision:
End Function


Private Sub lblGen_Click()

Screen.MousePointer = 11 'Makes the mouse pointer turn to 'busy' so you know that program is doing something

DoEvents 'Make computer do things without rushing(and trying to calculate faster than its can)

For a = 0 To Form1.ScaleHeight - 1 'Loops through each of forms pixels in a column
For b = 0 To Form1.ScaleWidth - 1 'Loops through each of forms pixels in a row of the a column

If Form1.Point(b, a) = 0 Then 'Check documentation
CollisionArray(a) = CollisionArray(a) & 1
Else
CollisionArray(a) = CollisionArray(a) & 0
End If

Next
'Debug.Print a & " " & CollisionArray(a)
Next

Screen.MousePointer = 0 'Turn mouse pointer to normal so you know its done

CollisionGenerated = True 'Enables you to move your shape
End Sub
