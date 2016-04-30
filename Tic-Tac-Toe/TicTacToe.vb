'------------------------------------------------------------------
'-                File Name : TicTacToe
'-                Part of Project: TicTacToe
'------------------------------------------------------------------
'-                Writen By: David Godi
'-                Written On: 03/15/15
'------------------------------------------------------------------
'- File Purpose:
'-   This file contains the main startup of the program
'------------------------------------------------------------------
'- Program Purpose: 
'-
'- This program demonstrates drag and drop if the classic game of 
'- Tic-Tac-Toe
'------------------------------------------------------------------
'- Global Variable Dictionary (alphabetically)
'- lblFocus     - current gaem cell with focus
'- lblgameCells - array that contains the 3 x 3 cell labels  
'- strCellArray - array contains cells values
'- blnGameOver  - flag for game over
'- strWhoseTurn - determines whose turn
'------------------------------------------------------------------

Public Class frmGame
  Dim lblFocus As Label
  Dim lblgameCells(9) As Label
  Dim strCellArray(9) As String
  Dim blnGameOver As Boolean
  Dim strWhoseTurn As String = "X"

  '------------------------------------------------------------------
  '-                Subprogram Name : frmGame_Load 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the form is loaded and intializes
  '- the array that holds labels containing the values the player drags
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------

  Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ' intitlaize labels into array 
    lblgameCells(0) = addHandlerToLabel(lbl1)
    lblgameCells(1) = addHandlerToLabel(lbl2)
    lblgameCells(2) = addHandlerToLabel(lbl3)
    lblgameCells(3) = addHandlerToLabel(lbl4)
    lblgameCells(4) = addHandlerToLabel(lbl5)
    lblgameCells(5) = addHandlerToLabel(lbl6)
    lblgameCells(6) = addHandlerToLabel(lbl7)
    lblgameCells(7) = addHandlerToLabel(lbl8)
    lblgameCells(8) = addHandlerToLabel(lbl9)

    lblGameOver.Text = "X's Turn"
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : frmGame_Paint
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called to repaint the ouput when an object 
  '-is moved or changed
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- myGraphics - instance of a the Graphics class for drawing objects
  '- myPen      - instance of the Pen class for drawing lines and curves
  '------------------------------------------------------------------
  Private Sub frmGame_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    ' declare drawing variables
    Dim myGraphics As Graphics = Me.CreateGraphics
    Dim myPen As Pen

    ' resize pen and assign a brush color
    myPen = New Pen(Brushes.Gray, 3)

    ' draw grid
    myGraphics.DrawLine(myPen, 277, 10, 277, 345)
    myGraphics.DrawLine(myPen, 387, 10, 387, 345)
    myGraphics.DrawLine(myPen, 160, 125, 506, 125)
    myGraphics.DrawLine(myPen, 160, 230, 506, 230)

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : addHandlerToLabel
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine creates Drag event handelers for the game cells
  '- note: calls method addDragOverHandeler to add DragOver handeler
  '-       calls method addDragEnterHandeler to add DragEnter handeler
  '-       calls method addDragDropHandeler to add DragDrop handeler
  '-       calls method addDragLeaveHandeler to add DragLeave handeler
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- lblObj - label for game cell      
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  '- Return
  '- Label
  '------------------------------------------------------------------
  Private Function addHandlerToLabel(ByRef lblObj As Label) As Label
    lblObj = addDragOverHandlerToLabel(lblObj)
    lblObj = addDragEnterHandlerToLabel(lblObj) '
    lblObj = addDragDropHandlerToLabel(lblObj)
    lblObj = addDragLeaveHandlerToLabel(lblObj)
    Return lblObj
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : addDragOverHandlerToLabel
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- add DragLeave handeler to game cell
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- lblObj - label for game cell             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  '- Return
  '- Label
  '------------------------------------------------------------------
  Private Function addDragOverHandlerToLabel(ByRef lblvalue As Label) As Label
    AddHandler lblvalue.DragOver, AddressOf lbl_DragOver
    Return lblvalue
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : addDragOverHandlerToLabel
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- add DragLeave handeler to game cell
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- lblObj - label for game cell             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  '- Return
  '- Label
  '------------------------------------------------------------------
  Private Function addDragEnterHandlerToLabel(ByRef lblObj As Label) As Label
    AddHandler lblObj.DragEnter, AddressOf lbl_DragEnter
    Return lblObj
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : addDragDropHandlerToLabel
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- add DragDrop handeler to game cell
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- lblObj - label for game cell             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  '- Return
  '- Label
  '------------------------------------------------------------------
  Private Function addDragDropHandlerToLabel(ByRef lblObj As Label) As Label
    AddHandler lblObj.DragDrop, AddressOf lbl_DragDrop
    Return lblObj
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : addDragLeaveHandlerToLabel
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- add DragLeave handeler to game cell
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- lblObj - label for game cell             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  '- Return
  '- Label
  '------------------------------------------------------------------
  Private Function addDragLeaveHandlerToLabel(ByRef lblObj As Label) As Label
    AddHandler lblObj.DragLeave, AddressOf lbl_DragLeave
    Return lblObj
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : lbl_MouseMove
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- calls subroutine when the mouse its moved over lblX or lblO
  '- if mouse button is not clicked then exit sub
  '- elseif: sender's text value != strWhoseTurn then exit sub
  '- else: initiate drage and drop controls on sender
  '- note: if mousde
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the MouseEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- lblSender   - sender object as Label
  '- chrBtnValue - lable value
  '- data        - data Object for DoDragDrop method
  '- effect      - type of DragDropEffect
  '------------------------------------------------------------------
  Private Sub lbl_MouseMove(sender As Object, e As MouseEventArgs) Handles lblX.MouseMove, lblO.MouseMove
    Dim lblSender As Label = DirectCast(sender, Label)
    Dim chrBtnValue As Char = Convert.ToChar(lblSender.Text.Substring(lblSender.Text.Length - 1))

    'If no mouse button was pressed, exit
    If e.Button = 0 Or blnGameOver Then
      Exit Sub
    ElseIf Not chrBtnValue = strWhoseTurn Then
      Exit Sub
    Else
      lblFocus = lblSender
    End If

    '**
    '* set up drag drop control when mouse button is clicked on sender
    '**

    ' copy the data associated with the sender for the DoDragDrop 
    '-method to work properly
    Dim data As New DataObject()

    'set DragDropEffect ...  + sign
    Dim effect As DragDropEffects = DragDropEffects.Copy

    ' set data to copy
    data.SetData("")

    'Initiate the dragdrop on the copy of the target control we made
    effect = lblSender.DoDragDrop(data, effect)

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : lbl_DragEnter
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- calls subroutine when drag event is active on sender
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the DragEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub lbl_DragEnter(sender As Object, e As DragEventArgs)
    e.Effect = DragDropEffects.Copy
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : lbl_DragDrop
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- calls subroutine when drag drop event is active on sender
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the DragEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- lblSender - sender object as Label
  '- intIndex  - game cell value value
  '------------------------------------------------------------------
  Private Sub lbl_DragDrop(sender As Object, e As DragEventArgs)
    ' covert sender object to Label
    Dim lblSender As Label = DirectCast(sender, Label)
    Dim intIndex As Integer = CInt(Convert.ToString(lblSender.Name).Substring(3))

    ' insert X or O if sender text has no vlaue and game is not over
    If lblSender.Text = "" And blnGameOver = False Then
      If lblFocus.Text = "X" Then
        lblSender.Text = "X"
      Else
        lblSender.Text = "O"
      End If
    End If

    ' change back gorund color of label
    lblSender.BackColor = Color.Transparent

    ' add label to array
    lblgameCells(intIndex) = lblSender

    ' add label value to array
    strCellArray(intIndex) = lblSender.Text

    run(lblSender.Text)

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : lbl_DragOver
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- calls subroutine when drag over event is active on sender
  '- note: toggles background color of cell based on text value
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the DragEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- lblSender - sender object as Label
  '------------------------------------------------------------------
  Private Sub lbl_DragOver(sender As Object, e As DragEventArgs)
    ' covert sender object to Label
    Dim lblSender As Label = DirectCast(sender, Label)

    ' set background color of label to red if sender text is empty
    ' else: set background color to green
    If Not lblSender.Text = "" Then
      lblSender.BackColor = Color.Red
    Else
      lblSender.BackColor = Color.Green
    End If
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : lbl_DragLeave
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- calls subroutine when drag leave event is active on sender
  '- note: background color of cell is to back to default color 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the DragEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- lblSender - sender object as Label
  '------------------------------------------------------------------
  Private Sub lbl_DragLeave(sender As Object, e As EventArgs)
    ' covert sender object to Label
    Dim lblSender As Label = DirectCast(sender, Label)
    lblSender.BackColor = Color.Transparent
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : run
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- main run of game
  '- note: calls gameWon to dertermine if a player won
  '-       calls isFull to determine a tie
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strCellValue - label value            
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub run(ByVal strCellValue As String)
    ' if get is not over check for a winner or a tie
    If Not blnGameOver Then
      ' if a plyer wins end game
      If gameWon(strCellValue) Then
        lblGameOver.Text = "Game goes to " + strCellValue
        blnGameOver = True
        ' if no player wins end game
      ElseIf isFull() Then
        lblGameOver.Text = "Draw! Game Over"
        blnGameOver = True

        ' toggle the players turn
      Else
        If strCellValue = "X" Then
          lblGameOver.Text = "O's Turn"
          strWhoseTurn = "O"
        Else
          strWhoseTurn = "X"
          lblGameOver.Text = "X's Turn"
        End If
      End If

    End If
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : isFull
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- returns true if all cells are full
  '- else: return false
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the DragEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Function isFull() As Boolean

    ' loop determines if cells are empty
    For inti As Integer = 0 To 8
      If strCellArray(inti) = "" Then
        Return False
      End If
    Next

    Return True
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : gameWon
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- returns true if labels compared are within limits, winner found
  '- else: false, can't find winner
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strValue - value X or O            
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Function gameWon(ByVal strvalue As String) As Boolean
    ' loop for finding matches horizontal or vertical
    For inti As Integer = 0 To 2
      If strCellArray(1 + (3 * inti)) = strvalue And strCellArray(2 + (3 * inti)) = strvalue And strCellArray(3 + (3 * inti)) = strvalue Then
        lblgameCells(1 + (3 * inti)).BackColor = Color.Blue
        lblgameCells(2 + (3 * inti)).BackColor = Color.Blue
        lblgameCells(3 + (3 * inti)).BackColor = Color.Blue
        Return True
      End If

      ' vertical matching
      If strCellArray(inti + 1 + (3 * 0)) = strvalue And strCellArray(inti + 1 + (3 * 1)) = strvalue And strCellArray(inti + 1 + (3 * 2)) = strvalue Then
        lblgameCells(inti + 1 + (3 * 0)).BackColor = Color.Blue
        lblgameCells(inti + 1 + (3 * 1)).BackColor = Color.Blue
        lblgameCells(inti + 1 + (3 * 2)).BackColor = Color.Blue
        Return True
      End If
    Next

    ' dianglegonal matching
    If strCellArray(1) = strvalue And strCellArray(5) = strvalue And strCellArray(9) = strvalue Then
      lblgameCells(1).BackColor = Color.Blue
      lblgameCells(5).BackColor = Color.Blue
      lblgameCells(9).BackColor = Color.Blue
      Return True
    End If

    If strCellArray(3) = strvalue And strCellArray(5) = strvalue And strCellArray(7) = strvalue Then
      lblgameCells(3).BackColor = Color.Blue
      lblgameCells(5).BackColor = Color.Blue
      lblgameCells(7).BackColor = Color.Blue
      Return True
    End If

    Return False

  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : btnReset_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/29/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- resets game
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the DragEventArgs object sent to the routine             
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click

    blnGameOver = False
    lblGameOver.Text = ""
    strWhoseTurn = "X"

    ' loop for clearing game cell labels
    For inti As Integer = 0 To strCellArray.Length() - 1
      lblgameCells(inti).Text = Nothing
      lblgameCells(inti).BackColor = Color.Transparent
    Next

    ' loop clears values in array
    For inti As Integer = 0 To strCellArray.Length() - 1
      strCellArray(inti) = ""
    Next

  End Sub
End Class
