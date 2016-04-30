'------------------------------------------------------------------
'-                File Name : Module1.vb
'-                Part of Project: CollegeInfo
'------------------------------------------------------------------
'-                Writen By: David Godi
'-                Written On: 04/7/15
'------------------------------------------------------------------
'- File Purpose:
'- This file creates and populates the database that is called 
'- from the CollegeInfo form
'------------------------------------------------------------------
'- Program Purpose: 
'-
'- This program creates a database conssting of colleges and the 
'- degrees they provide
'------------------------------------------------------------------
'- Global Variable Dictionary (alphabetically)
'- (None)
'------------------------------------------------------------------

Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms

Module Module1
  '------------------------------------------------------------------
  '-                Subprogram Name : Main 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- main drive of the program
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- (None)     
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- strConn   - connection name
  '- strDBName - name of database
  '- myForm    - object main form
  '------------------------------------------------------------------
  Sub Main()
    Dim strConn As String
    Dim strDBName As String = "DBfile.mdb"
    Dim myForm As New frmCollegeDegreeInfo

    'Pointer to the database being used
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBName

    'If the database doesn't exist, create and populate a new access database
    If Not (File.Exists(strDBName)) Then
      CreateDatabase(strConn)
      PopulateCollegesTable(strConn)
      PopulateDegreesTable(strConn)
    End If

    myForm.ShowDialog()

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : CreateDatabase 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- creates the tables
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strConn - name of connection     
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- DBCat  - creates access database
  '- DBConn - connect to database
  '- DBCmd  - commands used to execute against a data source
  '------------------------------------------------------------------
  Sub CreateDatabase(ByVal strConn As String)
    'Let's build an Access database
    Dim DBCat As New ADOX.Catalog()   ' creates access database
    Dim DBConn As OleDbConnection     'connect to database
    Dim DBCmd As OleDbCommand = New OleDbCommand() ' tell it what to do

    'try to build an empty database
    Try
      DBCat.Create(strConn)
      MessageBox.Show("Created database")
    Catch Ex As Exception
      MessageBox.Show("Database already exists")
    End Try

    ' connect to database
    DBConn = New OleDbConnection(strConn)
    DBConn.Open()

    'Build the college Table
    DBCmd.CommandText = "CREATE TABLE Colleges (" & _
                        "CID INT NOT NULL, " & _
                        "CollegeName VARCHAR(50), " & _
                        "Address VARCHAR(50), " & _
                        "City VARCHAR(50), " & _
                        "State VARCHAR(2), " & _
                        "ZipCode VARCHAR(5), " & _
                        "PRIMARY KEY(CID))"

    DBCmd.Connection = DBConn

    Try
      DBCmd.ExecuteNonQuery()
      MessageBox.Show("Created College Table")
    Catch Ex As Exception
      MessageBox.Show("College Table Already Exists")
    End Try

    'Build the degree Table
    DBCmd.CommandText = "CREATE TABLE Degrees (" & _
                        "DID INT NOT NULL, " & _
                        "DegreeName VARCHAR(50), " & _
                        "DegreeDesignator VARCHAR(50), " & _
                        "CreditsRequired VARCHAR(10), " & _
                        "EstimatedTimeOfCompletion VARCHAR(10), " & _
                        "CollegeTUID INT, " & _
                        "PRIMARY KEY(DID), " & _
                        "FOREIGN KEY (CollegeTUID) REFERENCES Colleges(CID))"

    DBCmd.Connection = DBConn
    Try
      DBCmd.ExecuteNonQuery()
      MessageBox.Show("Created Degrees Table")
    Catch Ex As Exception
      MessageBox.Show("Degrees Table Already Exists")
    End Try

    ' close connection
    DBConn.Close()

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : PopulateCollegesTable 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- poulates Colleges table
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strConn - name of connection     
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- DBConn - connect to database
  '- DBCmd  - commands used to execute against a data source
  '------------------------------------------------------------------
  Sub PopulateCollegesTable(ByVal strConn As String)
    Dim DBConn As OleDbConnection = New OleDbConnection(strConn)
    Dim DBCmd As OleDbCommand = New OleDbCommand()

    'open up a connection to the database
    DBConn.Open()
    DBCmd.Connection = DBConn

    'Add a college using SQL
    DBCmd.CommandText = "INSERT INTO Colleges (CID, CollegeName, Address, City, State, ZipCode) " & _
       "VALUES (1, 'SVSU','7400 Bay Road','University Center','MI','48710')"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Colleges (CID, collegeName, address, city, state, zipCode) " & _
    "VALUES (2, 'Delta College','1961 Delta Road','University Center','MI','48710')"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Colleges (CID, collegeName, address, city, state, zipCode) " & _
      "VALUES (3, 'James Technical','100 Coder Blvd','Saginaw','MI','48604')"
    DBCmd.ExecuteNonQuery()

    DBConn.Close()
    MessageBox.Show("Added a College To Colleges Table")

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : PopulateDegreesTable 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 4/7/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- poulates Colleges table
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strConn - name of connection     
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- DBConn - connect to database
  '- DBCmd  - commands used to execute against a data source
  '------------------------------------------------------------------
  Sub PopulateDegreesTable(ByVal strConn As String)
    Dim DBConn As OleDbConnection = New OleDbConnection(strConn)
    Dim DBCmd As OleDbCommand = New OleDbCommand()

    'open up a connection to the database
    DBConn.Open()
    DBCmd.Connection = DBConn

    'Add a degree using SQL
    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
       "VALUES (1, 'CIS','BS','124','4', 1)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
 "VALUES (2, 'CS','BS','124','4', 1)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
"VALUES (3, 'Business Administration','MBA','48','2', 1)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
"VALUES (4, 'Auto Mechanic','AA','62','2',2)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
"VALUES (5, 'CS','AS','62','2',2)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
"VALUES (6, 'Coding Wizard','BS','124','5',3)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
"VALUES (7, 'Hacking For Fun','AS','62','3',3)"
    DBCmd.ExecuteNonQuery()

    DBCmd.CommandText = "INSERT INTO Degrees (DID, degreeName, degreeDesignator, creditsRequired, estimatedTimeOfCompletion, collegeTUID) " & _
"VALUES (8, 'Business Analysis','BA','124','5',3)"
    DBCmd.ExecuteNonQuery()

    DBConn.Close()
    MessageBox.Show("Added a Degree To Degreess Table")

  End Sub

End Module
