'------------------------------------------------------------------
'-                File Name : MDI_Calculator
'-                Part of Project: MDI_Calculator
'------------------------------------------------------------------
'-                Writen By: David Godi
'-                Written On: 03/15/15
'------------------------------------------------------------------
'- File Purpose:
'-   This file contains the main application form where the user 
'-   interaction affects the program output based on their actions
'------------------------------------------------------------------
'- Program Purpose: 
'-
'- This program demonstrates MDI forms by allowing the user to 
'- create multiple forms containing a calcualotr for displaying
'- a number represented in binary, decimal, and hex.  Also this program can
'- demonstrate bit operators and, or, Xor, and not.
'------------------------------------------------------------------
'- Global Variable Dictionary (alphabetically)
'- (None)
'------------------------------------------------------------------

Public Class frmMDI_Calculator

  '------------------------------------------------------------------
  '-                Subprogram Name : mnuExit_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called whenever the user clicks the "Exit" button  
  '- located in the menu. Terminates program 
  '- 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
    Me.Close()
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : mnuAbout_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called whenever the user clicks the "About" button  
  '- located in the menu. Displays information about the program 
  '- 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub mnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
    'Create a new instance of the About form
    Dim objAboutBox As New frmAbout
    objAboutBox.ShowDialog()

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : mnuNew_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called whenever the user clicks the "New" button  
  '- located in the menu. Creates and displays new form 
  '- 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub mnuNew_Click(sender As Object, e As EventArgs) Handles mnuNew.Click
    'Create another child
    Dim objChildForm As New frmCalculator() ' instance of frmChild

    ' increment the number of forms 
    My.Application.intFormCount += 1

    ' creates a name for the form
    objChildForm.Name &= "-" & Trim(CStr(My.Application.intFormCount))
    objChildForm.Text = "Calc " & My.Application.intFormCount

    'Hook the child to the parent
    objChildForm.MdiParent = Me

    'Show the child
    objChildForm.Show()

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : mnuCascade_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called whenever the user clicks the sub menu "Cascade" 
  '- button located in the menu. Displays the forms by overlaping
  '- 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub mnuCascade_Click(sender As Object, e As EventArgs) Handles mnuCascade.Click
    Me.LayoutMdi(MdiLayout.Cascade)
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : mnuHorizontal_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called whenever the user clicks the sub menu "Horizontal" 
  '- button located in the menu. Displays the forms as horizontal rows
  '- 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub mnuHorizontal_Click(sender As Object, e As EventArgs) Handles mnuHorizontal.Click
    Me.LayoutMdi(MdiLayout.TileHorizontal)
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : mnuVertical_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called whenever the user clicks the sub menu "Vertical" 
  '- button located in the menu. Displays the forms as vertical columns
  '- 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub mnuVertical_Click(sender As Object, e As EventArgs) Handles mnuVertical.Click
    Me.LayoutMdi(MdiLayout.TileVertical)
  End Sub
End Class
