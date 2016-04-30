'------------------------------------------------------------------
'-                File Name : Calculator
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
'- This program displays an calculator for converting in binary, decimal, 
'- and hex.  Also this program can demonstrate bit operators and, or, Xor, 
'- and not operator
'------------------------------------------------------------------
'- Global Variable Dictionary (alphabetically)
'- btnBinBtnArray     - binary button array
'- btnHexBtnArray     - hex button array
'- btnDecimalBtnArray - decimal button array
'- txtAllTxtvalues    - txt boxes in value columns
'- txtFocusForCalc    - txtbox with current focus
'- blnValue           - flag for value 1 and value 2
'- myCalcType         - determines base type
'------------------------------------------------------------------


Public Class frmCalculator

  ' data structure that holds the current focus
  Enum udtCalcType
    BINARYVALUE
    DECIMALVALUE
    HEXVALUE
  End Enum

  ' declare variables
  Dim btnBinBtnArray(1) As Button
  Dim btnHexBtnArray(5) As Button
  Dim btnDecimalBtnArray(7) As Button
  Dim txtAllTxtvalues(1) As Object
  Dim txtFocusForCalc As TextBox
  Dim blnValue As Boolean
  Dim myCalcType As udtCalcType

  '------------------------------------------------------------------
  '-                Subprogram Name : frmCalculator_FormClosing
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the user clicks the close 'x' on 
  '- the form.  Also this program is called when exit is selected
  '- from the menu located on teh parent or the close 'x'
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- blnDirty   - flag for dirty form
  '- strMessage - message dissplaying quit message
  '------------------------------------------------------------------

  Private Sub frmCalculator_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    Dim blnDirty As Boolean
    Dim strMessage As String = "Are you sure you want to quit " & Me.Text & "?"

    ' loop through to find any textBoxes with a value
    For intOuterCount As Integer = 0 To 1
      For intInnerCount As Integer = 0 To 2
        ' if textbox has a value then set flag
        If txtAllTxtvalues(intOuterCount)(intInnerCount).Text.Length > 0 Then
          blnDirty = True
        End If
      Next
    Next

    ' display message if form is dirty
    If blnDirty Then
      If MessageBox.Show(strMessage, "Confirm", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Exclamation) = DialogResult.No Then
        e.Cancel = True
      Else
        ' permanently remove source from form
        Me.Dispose()
      End If
    End If

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : frmCalculator_Load 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when 'new' is selected from 
  '- frmMDI_Calculator menu and initializes values and button handlers 
  '- for calculator.
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- txtValueArray1 - contains text boxes for value 1
  '- txtValueArray2 - contains text boxes for value 2
  '------------------------------------------------------------------

  Private Sub frmCalculator_Load(sender As Object, e As EventArgs) Handles Me.Load

    ' array contains text boxes for value 1
    Dim txtValueArray1(2) As TextBox
    txtValueArray1(0) = txtBinary1
    txtValueArray1(1) = txtDecimal1
    txtValueArray1(2) = txtHex1

    ' array contains text boxes for value 2
    Dim txtValueArray2(2) As TextBox
    txtValueArray2(0) = txtBinary2
    txtValueArray2(1) = txtDecimal2
    txtValueArray2(2) = txtHex2

    txtAllTxtvalues(0) = txtValueArray1
    txtAllTxtvalues(1) = txtValueArray2

    ' binary button array
    btnBinBtnArray(0) = btn0
    btnBinBtnArray(1) = btn1

    ' decimal button array
    btnDecimalBtnArray(0) = btn2
    btnDecimalBtnArray(1) = btn3
    btnDecimalBtnArray(2) = btn4
    btnDecimalBtnArray(3) = btn5
    btnDecimalBtnArray(4) = btn6
    btnDecimalBtnArray(5) = btn7
    btnDecimalBtnArray(6) = btn8
    btnDecimalBtnArray(7) = btn9

    ' hex button array
    btnHexBtnArray(0) = btnA
    btnHexBtnArray(1) = btnB
    btnHexBtnArray(2) = btnC
    btnHexBtnArray(3) = btnD
    btnHexBtnArray(4) = btnE
    btnHexBtnArray(5) = btnF

    ' create event handlers for buttons in calculator
    AddButtonHandleler(btnBinBtnArray)
    AddButtonHandleler(btnDecimalBtnArray)
    AddButtonHandleler(btnHexBtnArray)

    ' Set flags to disable conver and bit operator buttons
    resetBitOperator(False)
    btnConvert.Enabled = False

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : txtBinary_GotFocus 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the txtBinary1 or txtBinary2 has focus
  '- which disables decimal and hex buttons.   
  '-
  '- note: txtFocusForCalc obtains the address of the textbox focus
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- txtTextBoxFocus - textbox object that contains the focus
  '- strSubString    - name of textbox selected
  '------------------------------------------------------------------
  Private Sub txtBinary_GotFocus(sender As Object, e As EventArgs) Handles txtBinary1.GotFocus, txtBinary2.GotFocus

    txtFocusForCalc = Nothing
    ' disable convert button, to prevent conversion without values
    btnConvert.Enabled = False

    ' get textbox object that has current focus
    Dim txtTextBoxFocus As TextBox = sender
    Dim strSubString As String = Convert.ToString(txtTextBoxFocus.Name).Substring(3)

    ' determine which binary textbox is selected and set flag for value category
    If String.Compare(strSubString, "Binary1") = 0 Then
      txtFocusForCalc = txtBinary1
      blnValue = True
    Else
      txtFocusForCalc = txtBinary2
      blnValue = False
    End If

    ' set buttons on calculator according to binary input
    ButtonControl(btnHexBtnArray, False)
    ButtonControl(btnDecimalBtnArray, False)
    myCalcType = udtCalcType.BINARYVALUE

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : txtDecimal_GotFocus 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the txtDecimal1 or txtDeimal2 has focus,
  '- which disables hex buttons.   
  '-
  '- note: txtFocusForCalc obtains the address of the textbox that has focus
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- txtTextBoxFocus - textbox object that contains the focus
  '- strSubString    - name of textbox selected
  '------------------------------------------------------------------
  Private Sub txtDecimal_GotFocus(sender As Object, e As EventArgs) Handles txtDecimal1.GotFocus, txtDecimal2.GotFocus
    txtFocusForCalc = Nothing
    ' disable convert button, to prevent conversion without values
    btnConvert.Enabled = False

    ' get textbox object that has current focus
    Dim txtTextBoxFocus As TextBox = sender
    Dim strSubString As String = Convert.ToString(txtTextBoxFocus.Name).Substring(3)

    ' determine which binary textbox is selected and set flag for value category
    If String.Compare(strSubString, "Decimal1") = 0 Then
      txtFocusForCalc = txtDecimal1
      blnValue = True
    Else
      txtFocusForCalc = txtDecimal2
      blnValue = False
    End If

    ' set buttons on calculator according to binary input
    ButtonControl(btnHexBtnArray, False)
    ButtonControl(btnDecimalBtnArray, True)
    myCalcType = udtCalcType.DECIMALVALUE

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : txtHex_GotFocus 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when the txtHex1 or txtHex2 has focus
  '- which enables hex and decimal buttons   
  '-
  '- note: txtFocusForCalc obtains the address of the textbox that has focus
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- txtTextBoxFocus - textbox object that contains the focus
  '- strSubString    - name of textbox selected
  '------------------------------------------------------------------
  Private Sub txtHex_GotFocus(sender As Object, e As EventArgs) Handles txtHex1.GotFocus, txtHex2.GotFocus
    txtFocusForCalc = Nothing

    ' disable convert button, to prevent conversion without values
    btnConvert.Enabled = False

    ' get textbox object that has current focus
    Dim txtTextBoxFocus As TextBox = sender
    Dim strSubString As String = Convert.ToString(txtTextBoxFocus.Name).Substring(3)

    ' determine which binary textbox is selected and set flag for value category
    If String.Compare(strSubString, "Hex1") = 0 Then
      txtFocusForCalc = txtHex1
      blnValue = True
    Else
      txtFocusForCalc = txtHex2
      blnValue = False
    End If

    ' set buttons on calculator according to binary input
    ButtonControl(btnHexBtnArray, True)
    ButtonControl(btnDecimalBtnArray, True)
    myCalcType = udtCalcType.HEXVALUE
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnClear_Click 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- This subrotine is called when "Clear Value 1", "Clear Value 2", "Clear Result"  
  '-
  '- note: clears value in value 2, vlue 2, and results category
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- btnButton   - receives the button obj that was clicked
  '- strTemp     - button name
  '- chrBtnValue - button value
  '------------------------------------------------------------------
  Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear1.Click, btnClear2.Click, btnClearResult.Click
    ' get button obj that clicked
    Dim btnButton = DirectCast(sender, Button)
    Dim strTemp As String = Convert.ToString(btnButton.Name)
    Dim chrBtnValue As Char = Convert.ToChar(strTemp.Substring(strTemp.Length - 1))

    ' determine what category to clear and reset flags for 
    ' bitoperator buttons and convert button
    Select Case Asc(chrBtnValue)
      Case 49
        txtBinary1.Text = ""
        txtDecimal1.Text = ""
        txtHex1.Text = ""
        resetBitOperator(False)
        btnConvert.Enabled = False
      Case 50
        txtBinary2.Text = ""
        txtDecimal2.Text = ""
        txtHex2.Text = ""
        resetBitOperator(False)
        btnConvert.Enabled = False
      Case Else
        txtBinaryResult.Text = ""
        txtDecimalResult.Text = ""
        txtHexResult.Text = ""
    End Select

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : resetBitOperator 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- enables bitoperator buttons if blnValue is true: else disables them  
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- blnValue - flag for disable or enable bit operator buttons
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub resetBitOperator(ByVal blnValue As Boolean)
    btnAdd1.Enabled = blnValue
    btnOr2.Enabled = blnValue
    btnXor3.Enabled = blnValue
    btnNotValue.Enabled = blnValue
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnConvert_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- toggles the disable and enable convert button values if  
  '-
  '- Note: calls method ConvertValues to convert binary, decimal, and hex 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- blnValue - flag for disable or enable bit operator buttons
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- blnDirty - flag for dirty text boxes in category "value 1" and "value 2"
  '------------------------------------------------------------------
  Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
    Dim blnDirty As Boolean

    ' converts binary, decimal, and hex values
    ConvertValues()

    ' sets flag to convert button
    For intOuterCount As Integer = 0 To 1
      For intInnerCount As Integer = 0 To 2
        If txtAllTxtvalues(intOuterCount)(intInnerCount).Text.Length > 0 Then
          blnDirty = True
        Else
          blnDirty = False
        End If
      Next
    Next

    ' enables convert button if blnDirty is true : else disable
    If blnDirty Then
      resetBitOperator(True)
    Else
      resetBitOperator(False)
    End If

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : Button_Click 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- 
  '- add the buttons value to the text box that has focus 
  '-
  '- note: txtFocusForCalc contins the address or the text box with focus
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine       
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- btnButton   - receives the button obj that was clicked
  '- strBtnValue - button value
  '------------------------------------------------------------------
  Private Sub Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ' get button name
    Dim btnButton = DirectCast(sender, Button)
    Dim strBtnValue As String = Convert.ToString(btnButton.Name).Substring(3)
    btnConvert.Enabled = True

    ' sets how the binary is displayed
    If udtCalcType.BINARYVALUE = myCalcType Then
      Dim intCount As Integer = txtFocusForCalc.Text.Length + 1
      If intCount Mod 5 = 0 Then txtFocusForCalc.Text += " "
    End If

    ' adds button value to the text box with focus
    txtFocusForCalc.Text += strBtnValue
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnBitOperator_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- caculates and or and Xor from two decimal values
  '-
  '- Note: calls method DecimalToBinary to covert decimal to binary
  '-       calls method built in method Hex to convert decimal to hex
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine 
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- btnButton   - receives the button obj that was clicked 
  '- strBtnTemp  - button name
  '- intBtnValue - button value
  '- strTemp     - binary string
  '------------------------------------------------------------------
  Private Sub btnBitOperator_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click, btnOr2.Click, btnXor3.Click
    Dim btnButton = DirectCast(sender, Button)
    Dim strBtnTemp As String = Convert.ToString(btnButton.Name)
    Dim intBtnValue As Integer = CInt(strBtnTemp.Substring(strBtnTemp.Length - 1))
    Dim strTemp As String = Nothing

    Select Case intBtnValue
      Case 1
        strTemp = CInt(txtDecimal1.Text) And CInt(txtDecimal2.Text)
      Case 2
        strTemp = CInt(txtDecimal1.Text) Or CInt(txtDecimal2.Text)
      Case 3
        strTemp = CInt(txtDecimal1.Text) Xor CInt(txtDecimal2.Text)
    End Select

    txtDecimalResult.Text = strTemp
    txtBinaryResult.Text = BinaryFormat(DecimalToBinary(strTemp))
    txtHexResult.Text = Hex(CInt(strTemp))
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : btnNotValue_Click
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- caculates the not value
  '-
  '- Note: converts neg decimal also by using two complimentent
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- sender – Identifies which particular control raised the  
  '-          click event                                     
  '- e – Holds the EventArgs object sent to the routine 
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- strDecimal  - decimal value
  '- strBinary   - binary string
  '------------------------------------------------------------------
  Private Sub btnNotValue_Click(sender As Object, e As EventArgs) Handles btnNotValue.Click
    Dim strDecimal As Integer = Not CInt(txtDecimal1.Text)
    Dim strBinary As String = ""

    If strDecimal < 0 Then
      strBinary = Convert.ToString(strDecimal, 2).PadLeft(32, "0"c)
    End If

    txtDecimalResult.Text = CStr(strDecimal)
    txtBinaryResult.Text = BinaryFormat(strBinary)
    txtHexResult.Text = Hex(CInt(strDecimal))

  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : ConvertValues 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- converts binary, deciaml and hex values if txtFocusForCalc
  '- enables bitoperator buttons if blnValue is true: else disables them  
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- blnValue - flag for disable or enable bit operator buttons
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub ConvertValues()
    ' determine the index for the current category taht has focus
    Dim intIndex As Integer = IIf(blnValue, 0, 1)

    ' if txtFocusForCalc has nothing disable convert button
    ' else : enable convert button and convert binary, decimal and hex
    If IsNothing(txtFocusForCalc) Then
      btnConvert.Enabled = False
    Else
      btnConvert.Enabled = True

      ' converts decimal and hex 
      Select Case myCalcType
        Case udtCalcType.BINARYVALUE
          Dim strBinaryValue As String = txtFocusForCalc.Text
          For intCountLength As Integer = 0 To 31 - strBinaryValue.Length
            strBinaryValue = "0" & strBinaryValue
          Next

          ' convert binary to decimal
          Dim strDecimal As String = BinaryToDecimal(strBinaryValue)

          ' convert decimal to hex
          Dim strHex As String = Hex(CInt(strDecimal))

          ' displays decimal and hex
          txtAllTxtvalues(intIndex)(0).Text = BinaryFormat(strBinaryValue)
          txtAllTxtvalues(intIndex)(1).Text = strDecimal
          txtAllTxtvalues(intIndex)(2).Text = strHex

        Case udtCalcType.DECIMALVALUE
          ' convert deciaml to binary and display 
          txtAllTxtvalues(intIndex)(0).Text = BinaryFormat(DecimalToBinary(txtFocusForCalc.Text))

          ' convert decimal to hex and display
          txtAllTxtvalues(intIndex)(2).Text = Hex(CInt(txtFocusForCalc.Text))

        Case udtCalcType.HEXVALUE
          ' convert hex to decimal
          Dim strDecimal As String = Convert.ToString(Convert.ToInt64(txtFocusForCalc.Text, 16))

          ' convert decimal to binary
          Dim strBinary As String = BinaryFormat(DecimalToBinary(strDecimal))

          ' dsiplay binary and decimal values
          txtAllTxtvalues(intIndex)(0).Text = strBinary
          txtAllTxtvalues(intIndex)(1).Text = strDecimal
      End Select

      ' disable convert button
      btnConvert.Enabled = False

    End If
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : AddButtonHandleler 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- add event handeler to button  
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- objButtonArray - button array
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub AddButtonHandleler(ByRef objButtonArray() As Button)
    For intCount As Integer = 0 To objButtonArray.Length - 1
      AddHandler objButtonArray(intCount).Click, AddressOf Me.Button_Click
    Next
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : ButtonControl 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- toggle enable and disable buttons  
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- objButtonArray - button array
  '- blnBtnEnabled  - button enabled value
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- (None)
  '------------------------------------------------------------------
  Private Sub ButtonControl(ByRef objButtonArray() As Button, ByVal blnBtnEnabled As Boolean)
    For intCount As Integer = 0 To objButtonArray.Length - 1
      objButtonArray(intCount).Enabled = blnBtnEnabled
    Next
  End Sub

  '------------------------------------------------------------------
  '-                Subprogram Name : DecimalToBinary 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- returns the conversion of decimal to binary 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strValue - decimal value to be coverted
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- strBinaryValue - binary value returned
  '- intDecimal     - decimal value
  '------------------------------------------------------------------
  '- Return value
  '- String
  '------------------------------------------------------------------
  Private Function DecimalToBinary(ByVal strValue As String) As String
    Dim strBinaryValue As String = ""
    Dim intDecimal As Integer = CInt(strValue)

    ' loop to convert deciaml to binary
    For intCount As Integer = 1 To 32
      Dim intTemp As Integer = intDecimal Mod 2

      ' add mod value to sring strBinaryValue
      strBinaryValue = CStr(intTemp) + strBinaryValue

      intDecimal \= 2
    Next

    strBinaryValue = strBinaryValue.Trim()
    Return strBinaryValue

  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : BinaryToDecimal 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- returns the conversion of binary to decimal 
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strValue - binary value to be coverted
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- intDecimal     - decimal value
  '------------------------------------------------------------------
  '- Return value
  '- String
  '------------------------------------------------------------------
  Private Function BinaryToDecimal(ByVal strValue As String) As String
    Dim intDecimal As Integer

    ' loop for converting binary to decimal
    For intCount As Integer = strValue.Length - 1 To 0 Step -1
      Dim intTemp As Integer = CStr(strValue(intCount))

      ' adds the power of two of intCount if bit is 1
      If intTemp = 1 Then
        intDecimal = 2 ^ ((strValue.Length - 1) - intCount) + intDecimal
      End If
    Next
    Return intDecimal
  End Function

  '------------------------------------------------------------------
  '-                Subprogram Name : BinaryFormat 
  '------------------------------------------------------------------
  '-                Writen By: David Godi
  '-                Written On: 03/15/15
  '------------------------------------------------------------------
  '- Program Purpose: 
  '- returns a binary number format of 32 bits
  '- 0000 0000 0000 0000 0000 0000 0000 0000
  '------------------------------------------------------------------
  '- Parameter Dictionary (in parameter order):               
  '- strValue - binary value to be coverted
  '------------------------------------------------------------------
  '- Local Variable dictionary (alphabetically)
  '- strBinary     - new binary value
  '------------------------------------------------------------------
  '- Return value
  '- String
  '------------------------------------------------------------------
  Private Function BinaryFormat(ByVal strValue As String)
    Dim strBinary As String = strValue

    For intCount As Integer = 0 To 37
      If intCount Mod 5 = 0 Then
        strBinary = strBinary.Insert(intCount, " ")
      End If
    Next

    Return strBinary.Trim()
  End Function

  

End Class