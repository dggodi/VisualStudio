Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms

'-------------------------------------------------------------------
'-                File Name : CollegeInfo.vb                    
'-                Part of Project: CollegeInfo                    
'------------------------------------------------------------------
'-                Writen By: David Godi
'-                Written On: 4/7/15
'------------------------------------------------------------------
'- File Purpose:                                            
'- This file contains the main form. were 
'- all the input and output will be entered
'------------------------------------------------------------
'- Program Purpose: 
'- 
'- main drive of the program
'------------------------------------------------------------------
'- Parameter Dictionary (in parameter order):               
'- (None)     
'------------------------------------------------------------------
'- Local Variable dictionary (alphabetically)
'- dsColleges     - memory for Colleges
'- dsDegrees      - memory for Degrees
'- binder         - binds Degrees table with the datGridView
'- strConn        - connection name
'- DBCmd          - commands or sql for dataset
'- DBAdaptCollege - commands used to fill the data set dsColleges
'- DBAdaptDegree  - commands used to fill the data set dsDegrees
'- intCurrentPos  - current pos of data
'- blnAdd         - flag for add and update
'------------------------------------------------------------------
Public Class frmCollegeDegreeInfo
  Dim dsColleges As New DataSet
  Dim dsDegrees As New DataSet
  Dim binder As New BindingSource

  ' connection string
  Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBfile.mdb"

  Dim DBCmd As OleDbCommand = New OleDbCommand
  Dim DBAdaptCollege As OleDbDataAdapter
  Dim DBAdaptDegree As OleDbDataAdapter

  Dim intCurrentPos As Integer
  Dim blnAdd As Boolean = True

  '------------------------------------------------------------------
  '-                Subprogram Name : frmInfoManager_Load 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when frmCollegeDegreeInfo is loaded.  
  '- The programs loads daa form database and binds the data
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- strSQLCommand - query string
  '- DBConn        - database connection
  '------------------------------------------------------------------
  Private Sub frmInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ' current row being displayed
    txtIDNumber.Text = 1

    ' set up query
    Dim strSQLCommand As String = "SELECT * FROM Colleges"
    Dim DBConn As OleDbConnection = New OleDbConnection(strConn)

    ' fill dataset with Colleges
    DBAdaptCollege = New OleDbDataAdapter(strSQLCommand, strConn)
    DBAdaptCollege.Fill(dsColleges, "Colleges")
    DBCmd.Connection = DBConn

    ' bind data
    txtName.DataBindings.Add(New Binding("Text", dsColleges, "Colleges.collegeName"))
    txtAddress.DataBindings.Add(New Binding("Text", dsColleges, "Colleges.address"))
    txtCity.DataBindings.Add(New Binding("Text", dsColleges, "Colleges.city"))
    txtState.DataBindings.Add(New Binding("Text", dsColleges, "Colleges.state"))
    txtZipCode.DataBindings.Add(New Binding("Text", dsColleges, "Colleges.zipCode"))
    txtIDNumber.DataBindings.Add(New Binding("Text", dsColleges, "Colleges.CID"))

    ' toggle readonly
    ToggleReadOnly(True)

    pnlSave.Visible = False

    ' update dataGridView
    UpdateGrid()
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnFirst_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the button <| 
  '-
  '- note : Displays the first row
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------

  Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
    BindingContext(dsColleges, "Colleges").Position = (BindingContext(dsColleges, "Colleges").Position = 0)
    intCurrentPos = 0
    UpdateGrid()
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnPrev_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the button << 
  '-
  '- note : Displays the previous row
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
    BindingContext(dsColleges, "Colleges").Position = (BindingContext(dsColleges, "Colleges").Position - 1)
    intCurrentPos -= 1
    UpdateGrid()
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnNext_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the button >> 
  '-
  '- note : Displays the next row
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
    BindingContext(dsColleges, "Colleges").Position = (BindingContext(dsColleges, "Colleges").Position + 1)
    intCurrentPos += 1
    UpdateGrid()
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnLast_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the button |> 
  '-
  '- note : Displays the last row
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
    BindingContext(dsColleges, "Colleges").Position = (dsColleges.Tables("Colleges").Rows.Count - 1)
    UpdateGrid()
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnAdd_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the Add button  
  '-
  '- note : adds a row to the Colleges table
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
    ' toggle readonly
    ToggleReadOnly(False, True)

    ' current row pos
    intCurrentPos = BindingContext(dsColleges, "Colleges").Position
    txtIDNumber.Text = dsColleges.Tables("Colleges").Rows.Count + 1

    ' update grid and flags
    UpdateGrid()
    pnlOption.Visible = False
    pnlSave.Visible = True
    blnAdd = True

    intCurrentPos += 1
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnUpdate_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the Ipdate button  
  '-
  '- note : update the current row to the Colleges table
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
    ' current row pos
    intCurrentPos = BindingContext(dsColleges, "Colleges").Position

    ' toggle readonly
    ToggleReadOnly(False)
    pnlOption.Visible = False
    pnlSave.Visible = True
    blnAdd = False
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnSave_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the Save button  
  '-
  '- note : save a new row to Colleges if flag blnAdd is true
  '-        else: upadte the current row
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- DBConn - database connection
  '------------------------------------------------------------------
  Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
    If blnAdd Then
      DBCmd.CommandText = "INSERT INTO Colleges (CID, CollegeName, Address, City, State, ZipCode) " & _
     "VALUES (" & txtIDNumber.Text & ", '" & txtName.Text & "', '" & txtAddress.Text & "', '" & txtCity.Text & "', '" & txtState.Text & _
    "', '" & txtZipCode.Text & "')"
    Else
      DBCmd.CommandText = "UPDATE Colleges SET CollegeName = '" & txtName.Text & _
                          "', Address = '" & txtAddress.Text & "', City = '" & txtCity.Text & _
                          "', State = '" & txtState.Text & "', ZipCode = '" & txtZipCode.Text & _
                          "'WHERE CID = " & txtIDNumber.Text
    End If

    BindingContext(dsColleges, "Colleges").EndCurrentEdit()

    ' connect and write to database
    Dim DBConn As OleDbConnection = New OleDbConnection(strConn)
    DBConn.Open()
    DBCmd.Connection = DBConn
    DBCmd.ExecuteNonQuery()
    DBConn.Close()

    'update dataset
    ToggleReadOnly(False)
    dsColleges.Clear()
    DBAdaptCollege.Fill(dsColleges, "Colleges")
    BindingContext(dsColleges, "Colleges").Position = intCurrentPos

    ' update grid and flags
    UpdateGrid()
    pnlSave.Visible = False
    pnlOption.Visible = True
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnCancel_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the Cancel button  
  '-
  '- note : cancels any update or new row being added
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (none)
  '------------------------------------------------------------------
  Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
    ' toggle readonly
    ToggleReadOnly(True)

    ' resets grid with current dataset and the current row pos
    dsColleges.Clear()
    DBAdaptCollege.Fill(dsColleges, "Colleges")
    BindingContext(dsColleges, "Colleges").Position = intCurrentPos
    UpdateGrid()

    pnlSave.Visible = False
    pnlOption.Visible = True

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnDelete_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the Delete button  
  '-
  '- note : cancels any update or new row being added
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- DBConn - database connection
  '------------------------------------------------------------------
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
    If MessageBox.Show("Are you sure you want to delete this record?",
                       "Delete Record", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then

      Dim DBConn As OleDbConnection = New OleDbConnection(strConn)
      DBConn.Open()

      DBCmd.CommandText = "DELETE FROM Degrees WHERE CollegeTUID = " & txtIDNumber.Text
      DBCmd.Connection = DBConn
      DBCmd.ExecuteNonQuery()

      DBCmd.CommandText = "DELETE FROM Colleges WHERE CID = " & txtIDNumber.Text
      DBCmd.Connection = DBConn
      DBCmd.ExecuteNonQuery()

      DBConn.Close()

      dsColleges.Clear()
      DBAdaptCollege.Fill(dsColleges, "Colleges")
      BindingContext(dsColleges, "Colleges").Position = intCurrentPos - 1
      UpdateGrid()

    End If
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnUpdateDegrees_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the 
  '- "Update Degrees Information" button.  
  '-
  '- note : update the degree table from the grid
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                       
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- DBConn - database connection
  '------------------------------------------------------------------
  Private Sub btnUpdateDegrees_Click(sender As Object, e As EventArgs) Handles btnUpdateDegrees.Click
    Dim DBConn As OleDbConnection = New OleDbConnection(strConn)

    Dim irow As Integer = dgInfo.CurrentRow.Index
    Dim icount As Integer = dsDegrees.Tables.Count
    BindingContext(dsDegrees, "Degrees").EndCurrentEdit()

    DBConn.Open()

    Debug.WriteLine("btnUpdateDegree    " & icount)
    'DBAdaptDegree.Update(dsDegrees, "Degrees")
    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
    "VALUES (" & icount & ", '" & dgInfo.Item(0, irow).Value & "', '" & dgInfo.Item(1, irow).Value & "', '" & dgInfo.Item(2, irow).Value & "', '" & dgInfo.Item(3, irow).Value & "', '" & dgInfo.Item(4, irow).Value & "')"
    DBCmd.Connection = DBConn
    DBCmd.ExecuteNonQuery()
    DBConn.Close()
    dsDegrees.AcceptChanges()

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : UpdateGrid
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- This subprogram is called to update the dataGridView
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- listTemp - represents the ListBox selected 
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- strSQLCommand - query
  '- table         - table in memory
  '------------------------------------------------------------------

  Sub UpdateGrid()
    Dim strSQLCommand As String = "SELECT * FROM Colleges"
    strSQLCommand = "SELECT DegreeName, DegreeDesignator, CreditsRequired, EstimatedTimeOfCompletion, CollegeTUID FROM Degrees WHERE CollegeTUID = " & _
        txtIDNumber.Text & ""

    DBAdaptDegree = New OleDbDataAdapter(strSQLCommand, strConn)
    DBAdaptDegree.Fill(dsDegrees, "Degrees")

    Dim table As New DataTable
    'Adds or refreshes rows in a specified range in the DataSet
    DBAdaptDegree.Fill(table)
    binder.DataSource = table
    dgInfo.DataSource = binder

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : ToggleReadOnly
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- This subprogram toggles the readonly values for textboxes
  '-
  '- also if the flage blnsave is true clear those textboxes 
  '- and make save and cancel appear
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- listTemp - represents the ListBox selected 
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Sub ToggleReadOnly(ByVal blnValue As Boolean, Optional ByVal blnSave As Boolean = False)
    txtName.ReadOnly = blnValue
    txtAddress.ReadOnly = blnValue
    txtCity.ReadOnly = blnValue
    txtState.ReadOnly = blnValue
    txtZipCode.ReadOnly = blnValue

    If blnSave Then
      txtName.Text = ""
      txtAddress.Text = ""
      txtCity.Text = ""
      txtState.Text = ""
      txtZipCode.Text = ""
      pnlSave.Visible = True
    End If
  End Sub
End Class


