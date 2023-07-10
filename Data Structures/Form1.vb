Option Strict On
Option Infer Off
Option Explicit On
Imports System.Collections.Immutable

'Program: Subs and Functions Practice
'Purpose: Chapter 6 of Visual Basic Book
'Name:Mason Merritt
'Date: Feb 18th,2022
Public Class Form1

    'Structure For What is A Data Structure Appilcation
    Structure Employee
        Public id As String
        Public firstName As String
        Public lastName As String
        Public pay As Double
    End Structure

    'Structure for Norbert Pool and Spa
    Structure Dimensions
        Public length As Double
        Public width As Double
        Public depth As Double
    End Structure

    'Structure for Wallpaper Application
    Structure Wallpaper
        Public length As Double
        Public width As Double
        Public height As Double
        Public coverage As Double
    End Structure

    'Structure for Cable Application
    Structure Cable
        Public premChannels As Integer
        Public connections As Integer
    End Structure

    'Structure for Paper WareHouse Application
    Structure ProductInfo
        Public id As String
        Public price As Double
    End Structure

    'Structure for Net Income App
    Structure MonthlyIncome
        Public userIncome As Double
        Public userExpenses As Double
        Public totalIncome As Double
    End Structure

    'Structure for PayRoll App
    Structure Taxes
        Public userName As String
        Public hours As Double
        Public payRate As Double
        Public userAllowance As Double
        Public grossPay As Double
        Public ficaTax As Double
        Public taxableIncome As Double
        Public fwt As Double
        Public netPay As Double
    End Structure

    'Structure for Excercise 18 Color Search App
    Structure ItemInfo
        Public id As String
        Public color As String
        Public price As Double
    End Structure

    Private items(9) As ItemInfo


    Private Sub btnWhatIs_Click(sender As Object, e As EventArgs) Handles btnWhatIs.Click

        'Data Structures
        'You can create your own datatype using the "STRUCTURE STATEMENT":
        '(   Structure structureName
        '         Public memberVariable as Datatype
        '         Public memberVariable as Datatype
        '    End Structure  )
        'Datatypes created using the structure statement are referred to as "USER-DEFINED DATATYPES" or "STRUCTURES"
        ' The structure is composed of "MEMEBERS" that are defined between the structure and end structure statement. They can be variables, constants or even procedures but more often than not they will be variables. the datatype
        'of the memeberVariable will determine the way the data will be stored it can even be ANOTHER STRUCTURE datatype
        ' Once the Structure has been built it still needs to be declared in code to do this use the "SYNTAX STATEMENT":
        '
        '(  " {Dim | Private} structureVariable as structureName "  )
        ' Similar to the way saying Dim age as Integer would store a variable the "HOURLY" variable will now contain 4 variables as you will see in the example to follow. To access these variables in code you use the (.) DOT Operator to
        'seperate the structureVariable name from the memberVariable name
        '
        '(  " structureVariableName.memeberVariableName = Value "  )
        ' the Structure is typically assigned in the Form's Class Decleration Section

        'Decalring a Structure Variable one is for Hourly Employees and the other for Salaried Employees
        Dim hourly As Employee
        Dim salaried As Employee

        'Setting the structureVariables values of the Hourly Employee structureVariable by using the dot operator to access the membervariables of the Structure Employee
        hourly.id = "H01"
        hourly.firstName = "Mason"
        hourly.lastName = "Merritt"
        hourly.pay = 20.5

        'Setting the structureVariables values of the Salaried Employee structureVariable by using the dot operator to access the membervariables of the Structure Employee
        salaried.id = "S01"
        salaried.firstName = "Veronica"
        salaried.lastName = "Chan"
        salaried.pay = 125750.0

        'Using structureVariables to assign calculated values
        'This will take the value in "hourly.pay" and multiple it by 1.5 and then reassign that value back to "hourly.pay"
        hourly.pay *= 1.5
        salaried.pay += 5000

        'Displaying structureVariables to User
        txtWhatISResults.Text = hourly.id & vbCrLf & hourly.firstName & vbCrLf & hourly.lastName & vbCrLf & hourly.pay.ToString &
            vbCrLf & vbCrLf & salaried.id & vbCrLf & salaried.firstName & vbCrLf & salaried.lastName & vbCrLf & salaried.pay.ToString

        'Programmers use structure variables when they need to pass a group of related items to a procedure for further processing . This is because it is easier to pass one structure variable than it is to pass many individual variables.
        'Programmers also use structure variables to store related items in an Array, even when memebers have different dataTypes. In the next Two Sections I will show you examples of how to pass a structureVariable to a procedure and also
        'how to store a structureVariable in an Array
    End Sub

    Private Sub btnNorbertPoolandSpa_Click(sender As Object, e As EventArgs) Handles btnNorbertPoolandSpa.Click

        'Create an Application that will take in UserInput of Pool "Dimensions" and then calculate how many gallons will be needed to fill the pool

        Dim poolSize As Dimensions

        Double.TryParse(txtLengthpool.Text, poolSize.length)
        Double.TryParse(txtWidthpool.Text, poolSize.width)
        Double.TryParse(txtDepthpool.Text, poolSize.depth)

        'Display gallons to the user

        txtGallonspool.Text = GetGallons(poolSize).ToString

    End Sub
    Private Function GetGallons(ByVal pool As Dimensions) As Double

        'Function to find the Gallons needed to Fill the Pool "GET GALLONS". I am passing into the function the memberVariables of the Structure "DIMENSIONS". The user inputs their DATA which is stored in a structureVariable "POOLSIZE". When I call the function
        'in the click event procedure and pass in the values of "POOLSIZE" into the Function the values of "Pool.Length, Pool.Width, Pool.Depth" will become the values of "poolSize.Length, poolSize.Width, poolSize.Depth" and then will calculate the amount of gallons
        'needed using the formula. The txtlabel will the display this value to the User.

        'Variable for Gallons per Cubic foot
        Const cubicFoot As Double = 7.48

        'Caluclation and Return Value
        Return pool.length * pool.width * pool.depth * cubicFoot
    End Function

    Private Sub btnWallpaperCalc_Click(sender As Object, e As EventArgs) Handles btnWallpaperCalc.Click

        'Excercise 12--Wallpaper Warehouse-- Build A program that will take in a users Length, Width and Height then allow them to select a paper size coverage
        'Calculate how many rolls will be needed to cover the room based upon the rolls coverage size. Use an Independent Sub to make the calculation.

        'structureVariable
        Dim newRoom As Wallpaper

        'TryParse
        Double.TryParse(cboLength.SelectedItem.ToString, newRoom.length)
        Double.TryParse(cboWidth.SelectedItem.ToString, newRoom.width)
        Double.TryParse(cboHeight.SelectedItem.ToString, newRoom.height)
        Double.TryParse(cboCoverage.SelectedItem.ToString, newRoom.coverage)

        'Call to Function
        WallpaperCalculation(newRoom)

        'Display Results to User
        txtWallpaperResults.Text = "You'll need " & WallpaperCalculation(newRoom).ToString & " rolls of wallpaper to cover the room"
    End Sub
    Private Function WallpaperCalculation(ByVal room As Wallpaper) As Double

        'Independent Function for Excercise 12 Wallpaper Warehouse Calculation
        Dim totalSquareFt As Double
        Dim rollsNeeded As Double

        'Calculation
        totalSquareFt = ((room.length * 2) * room.height) + ((room.width * 2) * room.height)
        rollsNeeded = totalSquareFt / room.coverage
        rollsNeeded = Math.Ceiling(rollsNeeded)

        'Return Value
        Return rollsNeeded
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Excercise 12 Wallpaper Warehouse Combo Boxes (Length, width, height display from 8 - 30) (coverage displays 30 - 40 step 0.5 increments)
        Dim length As Integer = 8
        While length <= 30
            cboLength.Items.Add(length)
            length += 1
        End While
        cboLength.Text = "10"

        Dim width As Integer = 8
        While width <= 30
            cboWidth.Items.Add(width)
            width += 1
        End While
        cboWidth.Text = "10"

        Dim height As Integer = 8
        While height <= 30
            cboHeight.Items.Add(height)
            height += 1
        End While
        cboHeight.Text = "8"

        Dim coverage As Double = 30
        While coverage <= 40
            cboCoverage.Items.Add(coverage)
            coverage += 0.5
        End While
        cboCoverage.Text = "37"

        'Excercise 11 Cable Direct Listboxes
        Dim pchannels As Integer
        While pchannels <= 10
            lstPremiumChannels.Items.Add(pchannels)
            pchannels += 1
        End While

        Dim connections As Integer
        While connections <= 100
            lstCableConnections.Items.Add(connections)
            connections += 1
        End While

        'Add Info to ListBoxes and ComboBoxes for Cerruti Payroll App
        Dim hours As Double = 1
        While hours <= 168
            lstHoursWorked.Items.Add(hours)
            hours += 0.5
        End While
        lstHoursWorked.SelectedIndex = 39

        Dim payrate As Double = 7.25
        While payrate <= 100
            lstPayRate.Items.Add(payrate)
            payrate += 0.25
        End While
        lstPayRate.SelectedIndex = 31

        Dim allow As Double = 1
        While allow <= 10
            cboAllowances.Items.Add(allow)
            allow += 1
            cboAllowances.SelectedIndex = 0
        End While

        'Add Items to Listboxes ID and COLOR for Excercise 18
        Dim itemColor() As String = {"Blue", "Red", "Blue", "Red", "White", "Red", "Blue", "Black", "White", "Blue"}
        Dim itemId() As String = {"101", "102", "103", "104", "105", "106", "107", "108", "109", "110"}

        For Each element As String In itemId
            lstItemID.Items.Add(element)
        Next element
        lstItemID.SelectedIndex = 0

        For Each element As String In itemColor
            lstItemColor.Items.Add(element)
        Next element
        lstItemColor.SelectedIndex = 0

        'structureVariable Values for Excercise 18
        items(0).id = "101"
        items(0).color = "Blue"
        items(0).price = 4.99
        items(1).id = "102"
        items(1).color = "Red"
        items(1).price = 4.99
        items(2).id = "103"
        items(2).color = "Blue"
        items(2).price = 10.49
        items(3).id = "104"
        items(3).color = "Red"
        items(3).price = 10.49
        items(4).id = "105"
        items(4).color = "White"
        items(4).price = 6.79
        items(5).id = "106"
        items(5).color = "Red"
        items(5).price = 6.79
        items(6).id = "107"
        items(6).color = "Blue"
        items(6).price = 6.79
        items(7).id = "108"
        items(7).color = "Black"
        items(7).price = 21.99
        items(8).id = "109"
        items(8).color = "Wjite"
        items(8).price = 21.99
        items(9).id = "110"
        items(9).color = "Blue"
        items(9).price = 21.99
    End Sub

    Private Sub btnCableCalc_Click(sender As Object, e As EventArgs) Handles btnCableCalc.Click

        'Excercise 11-- Build a Program that displays a Users cable bill based upon the seklected features they choose. use 2 Function Procedures
        'One for RESIDENTIAL customers with pricing as follows: $4.50 processing fee, $30 basic cable fee, $5 for each premium channel
        'The BUSINESS customer pricing is as follows: $16.50 processing fee, $80 for the first 4 connections then $4 for each additional connection
        'premium channels cost a flat rate $50 for any number of connections. Use one function for residential and one for business.

        'structureVariable
        Dim newTv As Cable

        'TryParse
        Integer.TryParse(lstPremiumChannels.SelectedItem.ToString, newTv.premChannels)
        Integer.TryParse(lstCableConnections.SelectedItem.ToString, newTv.premChannels)

        'Call to Functions
        If radResidential.Checked Then
            txtCableResults.Text = ResidentialBilling(newTv)
        ElseIf radBusiness.Checked Then
            txtCableResults.Text = BusinessBilling(newTv)
        End If
    End Sub

    Private Function ResidentialBilling(ByVal tvSetup As Cable) As String

        'Function Procedure to Calculate Residential Cable Billing

        'Variable to store Results
        Dim total As Double

        'Calculation
        If tvSetup.connections = 1 Then
            total = 4.5 + 30 + (5 * tvSetup.premChannels)
        Else
            MessageBox.Show("Residential TV SetUp only allows for one New Connection")
        End If

        'Return Value
        Return "Your bill is " & total.ToString("C2")
    End Function
    Private Function BusinessBilling(ByVal tvSetup As Cable) As String

        'Function Procedure to Calculate Business Cable Billing

        'Variable to store Results
        Dim total As Double

        'Calculation
        If tvSetup.connections <= 10 Then
            total = 16.5 + 80 + (50 * tvSetup.premChannels)
        ElseIf tvSetup.connections > 10 Then
            total = 96.5 + ((tvSetup.connections - 10) * 4) + (50 * tvSetup.premChannels)
        End If

        'Return Value
        Return "Your bill is " & total.ToString("C2")
    End Function

    Private Sub btnArrayDataStructure_Click(sender As Object, e As EventArgs) Handles btnArrayDataStructure.Click

        'Create an App that will create a Structure of Item Information. Then create an Array of structureVariables to search through and find the ID to display the price 

        Dim products(4) As ProductInfo
        Dim userSearch As String = txtArrayDataStucture.Text.ToUpper
        Dim intSub As Integer

        products(0).id = "A45G"
        products(0).price = 8.99
        products(1).id = "J63Y"
        products(1).price = 12.99
        products(2).id = "M93K"
        products(2).price = 5.99
        products(3).id = "C20P"
        products(3).price = 13.5
        products(4).id = "F77T"
        products(4).price = 7.25

        Do Until intSub = products.Length OrElse userSearch = products(intSub).id
            intSub += 1
        Loop

        If intSub < products.Length Then
            txtArrayPrice.Text = products(intSub).price.ToString("c2")
        Else
            MessageBox.Show("Invalid ID")
        End If
    End Sub

    Private Sub btnIncomeCalc_Click(sender As Object, e As EventArgs) Handles btnIncomeCalc.Click

        'Create an App that will create a Structure named "MonthlyIncome" to store monthly pay information(userIncome, userExpense, totalIncome). Then takes user input and finds the total monthly income.

        'declares a structureVariable "netIncome" set to the Structure "MonthlyIncome" which has 3 "memberVariables(userIncome,UserExpenses,totalIncome)"
        Dim netIncome As MonthlyIncome

        'TryParses the datatype strings input into the textbox controls into datatype double(numbers) and stores them in the structureVariables that are reflecting back to the memberVariables of the Structure(MonthlyIncome) by using the (.) dot operator
        Double.TryParse(txtUserIncome.Text, netIncome.userIncome)
        Double.TryParse(txtUserExpenses.Text, netIncome.userExpenses)

        'Call to the Function which passes the values of the structureVariable "netIncome" to the structureVariable created for the function which is "monthlyTotal" the procedure then calculates the totalIncome and then returns a string value to
        'display in the text control to the User
        txtUserNetIncome.Text = Income(netIncome)

        'Displaying the value of the structureVariable netIncome.totalIncome which has been passed back into the variable BY REFERENCE of the function procedure
        txtUserNetIncome.Text += vbCrLf & netIncome.totalIncome.ToString
    End Sub

    Private Function Income(ByRef monthlyTotal As MonthlyIncome) As String

        'Function to Calculate the TotalIncome. The Function creates a "monthlyTotal" variable that is set to the Structure "MonthlyIncome" which has 3 "memberVariables(userIncome, userExpenses, totalIncome)" and the Function is set to return a string value. Information will
        'be passed in ByReference to allow the structure variable netIncome.totalIncome to be assigned the value calculated in the function. When the button btnIncome is clicked the procedure will pass the values of variables (netIncome.userIncome and netIncome.userExpenses)
        'which has been Tryparsed from textboxes into the function variables (monthlytotal.userIncome, monthlyTotal.userExpenses) the Function then assigns the variable (monthlyTotal.totalIncome) to be equal to the difference of
        '(monthlytotal.userIncome - monthlyTotal.userExpenses).Monthlytotal.totalIncome is then set to return as a string value which will be assigned to the to the textbox control. The structureVariable "netIncome.totalIncome" is then also assigned the value of
        'monthlyTotal.totalIncome since the variables are being passed in BY REFERENCE

        monthlyTotal.totalIncome = monthlyTotal.userIncome - monthlyTotal.userExpenses

        Return monthlyTotal.totalIncome.ToString
    End Function

    Private Sub btnPayrollCalc_Click(sender As Object, e As EventArgs) Handles btnPayrollCalc.Click

        'Build the payroll App that will take in a User's Name, Martial Status, Hours Worked, Payrate and Total Allowances allowed. Calculate their Federal Tax Withholdings as well as their FICA tax withholding(7.65%) and display the total results of Gross Pay, FWT, FICA and
        'Net Pay. Use a Function Procedure to calculate for SINGLE and for Married.

        'Creates a structureVariable that is set to the Structure "TAXES"
        Dim userTaxes As Taxes

        'Sets the "structureVariables of userTaxes" to their values which are reflecting back to the "memeberVariables of the Structure Taxes"
        userTaxes.userName = txtPayrollName.Text
        Double.TryParse(lstHoursWorked.SelectedItem.ToString, userTaxes.hours)
        Double.TryParse(lstPayRate.SelectedItem.ToString, userTaxes.payRate)
        Double.TryParse(cboAllowances.Text, userTaxes.userAllowance)

        'declare constants

        Const oneAllowance As Double = 77.9
        Const ficaRate As Double = 0.0765

        'calculate gross pay

        If userTaxes.hours <= 40 Then
            userTaxes.grossPay = userTaxes.payRate * userTaxes.hours
        Else
            userTaxes.grossPay = (userTaxes.payRate * 40) + ((userTaxes.hours - 40) * userTaxes.payRate * 1.5)
        End If

        'calculate FICA tax

        userTaxes.ficaTax = userTaxes.grossPay * ficaRate

        'calculate taxable wages

        userTaxes.taxableIncome = userTaxes.grossPay - (userTaxes.userAllowance * oneAllowance)

        'which Tax bracket is being chosen to calculate FWT

        Select Case True
            Case radSingle.Checked
                Singletax(userTaxes)
            Case Else
                Marriedtaxes(userTaxes)
        End Select

        'round all numbers

        userTaxes.grossPay = Math.Round(userTaxes.grossPay, 2)
        userTaxes.ficaTax = Math.Round(userTaxes.ficaTax, 2)
        userTaxes.fwt = Math.Round(userTaxes.fwt, 2)

        userTaxes.netPay = userTaxes.grossPay - userTaxes.ficaTax - userTaxes.fwt
        Math.Round(userTaxes.netPay, 2)

        'Display Results of Procedures

        txtGrossPay.Text = userTaxes.grossPay.ToString("C2")
        txtFWH.Text = userTaxes.fwt.ToString("C2")
        txtFICA.Text = userTaxes.ficaTax.ToString("C2")
        txtNetPay.Text = userTaxes.netPay.ToString("C2")
    End Sub

    Private Function Singletax(ByRef SingleTaxablePay As Taxes) As Double

        'Excercise is to build the payroll App - This is the Function Procedure for the Single Taxes

        Select Case SingleTaxablePay.taxableIncome
            Case <= 44
                SingleTaxablePay.fwt = 0
            Case <= 224
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 44) * 0.1
            Case <= 774
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 224) * 0.15 + 18
            Case <= 1812
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 774) * 0.25 + 100.5
            Case <= 3730
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 1812) * 0.28 + 360
            Case <= 8058
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 3730) * 0.33 + 897.04
            Case <= 8090
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 8058) * 0.35 + 2325.28
            Case Else
                SingleTaxablePay.fwt = (SingleTaxablePay.taxableIncome - 8090) * 0.396 + 2336.48
        End Select
        Return SingleTaxablePay.fwt
    End Function

    Private Function Marriedtaxes(ByRef MarriedTaxablePay As Taxes) As Double

        'Excercise is to build the payroll App - This is the Function Procedure for the Single Taxes

        Select Case MarriedTaxablePay.taxableIncome
            Case <= 166
                MarriedTaxablePay.fwt = 0
            Case <= 525
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 166) * 0.1
            Case <= 1626
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 525) * 0.15 + 35.9
            Case <= 3111
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 1626) * 0.25 + 201.05
            Case <= 4654
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 3111) * 0.28 + 572.3
            Case <= 8180
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 4654) * 0.33 + 1004.34
            Case <= 9218
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 8180) * 0.35 + 2167.92
            Case Else
                MarriedTaxablePay.fwt = (MarriedTaxablePay.taxableIncome - 9218) * 0.396 + 2531.22
        End Select
        Return MarriedTaxablePay.fwt
    End Function

    Private Sub btn18Excercise_Click(sender As Object, e As EventArgs) Handles btn18Excercise.Click

        'EXCERCISE 18
        'Create an app that will allow a user to select item IDs from a Listbox and display their color and price. When the user selects a color from the listbox all the Item IDs with that color and their price shoudl be displayed
        'when a user searches a price all Item Ids color and Prices with that price and UNDER should be displayed

        'Variables
        Dim userSearch As Double
        Double.TryParse(txt18ExcerciseInput.Text, userSearch)

        lblPriceSearchResults.Text = ""
        For intRow As Integer = 0 To items.GetUpperBound(0)
            If userSearch >= items(intRow).price Then
                lblPriceSearchResults.Text += items(intRow).id & " " & items(intRow).color & " " & items(intRow).price & vbCrLf
            End If
        Next
    End Sub
    Private Sub lstItemID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstItemID.SelectedIndexChanged

        'EXCERCISE 18
        'Create an app that will allow a user to select item IDs from a Listbox and display their color and price. When the user selects a color from the listbox all the Item IDs with that color and their price shoudl be displayed
        'when a user searches a price all Item Ids color and Prices with that price and UNDER should be displayed

        Dim idSearch As String = lstItemID.SelectedItem.ToString

        lbl18ExcerciseResults.Text = ""
        For intRow As Integer = 0 To items.GetUpperBound(0)
            If idSearch = items(intRow).id Then
                lbl18ExcerciseResults.Text += items(intRow).id & " " & items(intRow).color & " " & items(intRow).price
            End If
        Next intRow

    End Sub
    Private Sub lstItemColor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstItemColor.SelectedIndexChanged

        'EXCERCISE 18
        'Create an app that will allow a user to select item IDs from a Listbox and display their color and price. When the user selects a color from the listbox all the Item IDs with that color and their price shoudl be displayed
        'when a user searches a price all Item Ids color and Prices with that price and UNDER should be displayed

        Dim colorSearch As String = lstItemColor.SelectedItem.ToString

        lbl18Results2.Text = ""
        For intRow As Integer = 0 To items.GetUpperBound(0)
            If colorSearch = items(intRow).color Then
                lbl18Results2.Text += items(intRow).id & " " & items(intRow).color & " " & items(intRow).price & vbCrLf
            End If
        Next intRow
    End Sub
End Class
