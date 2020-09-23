Attribute VB_Name = "Module1"
Option Explicit 'You dont really need this. It makes sure that every variable is declared or you'll get an error

Public Enum Side 'This is just to make thing easier
    Up = 0
    Down = 1
    Left = 2
    Right = 3
End Enum

Public CollisionGenerated As Boolean 'This will be true when user clicks generate collision

Public CollisionArray(199) As String 'This is the main array that holds collision information
Public ArrayPosition As Long 'This is used when checking collision for moving left and right(check the documentation


Public a As Long 'Check the SinkSub example in the documentation

Public b As Long 'Same as above
