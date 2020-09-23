VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example - Graphics from Chrono Trigger"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   25
      X1              =   120
      X2              =   120
      Y1              =   188
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   24
      X1              =   120
      X2              =   128
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   23
      X1              =   88
      X2              =   88
      Y1              =   188
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   22
      X1              =   88
      X2              =   80
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   21
      X1              =   120
      X2              =   120
      Y1              =   140
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   20
      X1              =   88
      X2              =   88
      Y1              =   140
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   19
      X1              =   120
      X2              =   128
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   18
      X1              =   88
      X2              =   80
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   17
      X1              =   48
      X2              =   88
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   16
      X1              =   88
      X2              =   88
      Y1              =   96
      Y2              =   104
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   15
      X1              =   128
      X2              =   128
      Y1              =   160
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   14
      X1              =   120
      X2              =   88
      Y1              =   188
      Y2              =   188
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   13
      X1              =   128
      X2              =   80
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   12
      X1              =   80
      X2              =   80
      Y1              =   184
      Y2              =   192
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   11
      X1              =   48
      X2              =   48
      Y1              =   96
      Y2              =   192
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   10
      X1              =   48
      X2              =   80
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   9
      X1              =   80
      X2              =   80
      Y1              =   160
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   8
      X1              =   88
      X2              =   120
      Y1              =   140
      Y2              =   140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   7
      X1              =   128
      X2              =   128
      Y1              =   136
      Y2              =   125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   6
      X1              =   128
      X2              =   150
      Y1              =   125
      Y2              =   125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   5
      X1              =   150
      X2              =   150
      Y1              =   125
      Y2              =   80
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   4
      X1              =   150
      X2              =   128
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   3
      X1              =   128
      X2              =   128
      Y1              =   96
      Y2              =   80
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   2
      X1              =   128
      X2              =   104
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   1
      X1              =   104
      X2              =   104
      Y1              =   96
      Y2              =   104
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      Index           =   0
      X1              =   88
      X2              =   104
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1800
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblGen 
      Caption         =   "Click This Label To Generate Collision"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example uses magenta color as unpassable

Option Explicit
Public LineArray As Long 'Used In generating collision


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

If Form1.Point(b, a) = vbMagenta Then 'Check documentation
CollisionArray(a) = CollisionArray(a) & 1
Else
CollisionArray(a) = CollisionArray(a) & 0
End If

Next
'Debug.Print a & " " & CollisionArray(a)
Next
Screen.MousePointer = 0 'Turn mouse pointer to normal so you know its done

For LineArray = Line1.LBound To Line1.UBound 'Loops through all the lines and hide them
Line1(LineArray).Visible = False
Next

CollisionGenerated = True 'Enables you to move your shape

End Sub
