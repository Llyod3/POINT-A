Public Class clsRectangle
    ' Calculate the area of a floor in a room
    1.1.2 Private _length As Decimal
    Private _width As Decimal

    1.1.3 Public Sub New()
        ' Constructor with default values
        _length = 0
        _width = 0
    End Sub

    1.1.4 Property Length()
        Get
            Return _length
        End Get
        Set(ByVal value)
            _length = value
        End Set
    End Property

    Property Width()
        Get
            Return _width
        End Get
        Set(ByVal value)
            _width = value
        End Set
    End Property

    1.1.5 Public Function CalcArea() As Decimal
        ' Calculate the area of the floor
        Dim Area As Decimal
        Area = _length * _width
        Return Area
    End Function
End Class

Public Class frmMain
    ' Calculate how much it will cost to tile the floor in a room
    1.2.2 Private myRectangle As clsRectangle

    Private Sub btnCalculate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCalculate.Click
        ' Calculate the total cost of tiling the floor
        Dim decTilePrice As Decimal = 0
        Dim decTotalCost As Decimal = 0

        ' Input
        myRectangle.Length = CDec(txtLength.Text)
        myRectangle.Width = CDec(txtWidth.Text)
        decTilePrice = CDec(txtPrice.Text)

        ' Processing
        decTotalCost = myRectangle.CalcArea() * decTilePrice

        ' Output
        lblTotalCost.Text = decTotalCost.ToString("C2")
    End Sub

    1.2.4 Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        ' Clear all the controls on the form
        txtLength.Clear()
        txtWidth.Clear()
        txtPrice.Clear()
        lblTotalCost.Text = ""
        txtLength.Focus()
    End Sub

    1.2.5 Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        ' Close the form and exit the application
        Me.Close()
    End Sub
End Class
—————————————-
Public Class frmCapture
    ' Examination Number

    2.2 Private Sub frmCapture_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetLecturers()
        txtName.Focus()
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        btnSubmit.Enabled = False

        Try
            ' Declare the connection and command objects.
            Dim connection As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand

            ' Initialize the ConnectionString property of the connection and open the database connection.
            connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Databases\CollegeSkyDB.mdb;User Id=;Password=;"
            connection.Open()

            ' Set the Connection and CommandText properties of the command object.
            command.Connection = connection
            command.CommandText = "INSERT INTO LecturerRecruit_Sky (Name, Surname, LecturerNumber, Department, Telephone) VALUES(@Name,@Surname,@LectureNumber,@Department,@Telephone)"

            ' Declare and initialize the parameters.
            2.3.5 Dim parName As New OleDb.OleDbParameter("@Name", txtName.Text)
            Dim parSurname As New OleDb.OleDbParameter("@Surname", txtSurname.Text)
            Dim parLecturerNumber As New OleDb.OleDbParameter("@LectureNumber", txtLecturerNumber.Text)
            Dim parDepartment As New OleDb.OleDbParameter("@Department", txtDepartment.Text)
            Dim parTelephone As New OleDb.OleDbParameter("@Telephone", txtTelephone.Text)

            ' Add the parameters to the command object.
            command.Parameters.Add(parName)
            command.Parameters.Add(parSurname)
            command.Parameters.Add(parLecturerNumber)
            command.Parameters.Add(parDepartment)
            command.Parameters.Add(parTelephone)

            ' Execute the command.
            command.ExecuteNonQuery()

            ' Close the database connection.
            connection.Close()

            ' Fetch the new result set for the Lecturer table.
            GetLecturers()
        Catch ex As Exception
            ' Prompt the user with the exception.
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    2.4 Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtName.Clear()
        txtSurname.Clear()
        txtLecturerNumber.Clear()
        txtDepartment.Clear()
        txtTelephone.Clear()
        txtName.Focus()
        btnSubmit.Enabled = True
    End Sub

    2.5 Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class

3.1 <html>
<head>
    <title>Examination Number</title>
</head>
<body>
    <h3>Briefly tell us about yourself</h3>
    <form>
        Region: <input type='text' size='15' maxlength='15' /><br />
        Gender: <input type='text' size='15' maxlength='6' /><br /><br />

        <b>Your field of study:</b> (1 - Electrical, 2 - Mechanical, 3 - Information Technology) <br />
        <select name='field'>
            <option value='1'>1</option>
            <option value='2'>2</option>
            <option value='3'>3</option>
        </select><br /><br />

        <b>Activities of interest:</b><br />
        <input type='checkbox' checked/> Debate <br />
        <input type='checkbox' /> Music and Drama <br />
        <input type='checkbox' /> Soccer <br />
        <input type='checkbox' /> Netball <br />
        <input type='checkbox' /> Chess <br /><br />

        <b>Are you experiencing any problems with your studies?</b><br />
        <input type='radio' name='coping' value='yes' checked /> Yes <br />
        <input type='radio' name='coping' value='no' /> No <br /><br />

        <b>If your answer is YES to the above question, briefly elaborate:</b><br />
        <textarea rows='5' cols='40'>Enter your comments here.</textarea><br /><br />

        <input type='reset' value='Clear form' />
    </form>
</body>
</html>

4.1 <html>
<head>
    <title>About Us</title>
</head>
<body>
    <h2>About Us</h2>
    <img src='Handshake.jpg' width='180' height='180' />
    <hr />
    <h3>Vision</h3>
    <p>To be the best College in Southern Africa</p>
    <hr />
    <h3>Mission</h3>
    <p>We educate the learner</p>
    <ul>
        <li>Mentally</li>
        <li>Physically</li>
        <li>Spiritually</li>
    </ul>
</body>
</html>

4.2 <html>
<head>
    <title>Enquiries</title>
</head>
<body>
    <h4><b>Enquire about Department:</b></h4>
    <table border='3'>
        <tr>
            <th>Head of Department</th>
            <th>Name of Department</th>
            <th>Telephone</th>
        </tr>
        <tr>
            <td>Mr Photo</td>
            <td>Information Technology</td>
            <td>(012)386 6112</td>
        </tr>
        <tr>
            <td>Mr Tau</td>
            <td>Mechanical Engineering</td>
            <td>(012)386 6113</td>
        </tr>
        <tr>
            <td>Mr Arendse</td>
            <td>Mechatronics</td>
            <td>(012)386 6114</td>
        </tr>
        <tr>
            <td>Mrs Shibambo</td>
            <td>Life Sciences</td>
            <td>(012)386 6115</td>
        </tr>
        <tr>
            <td>Mr Mthambo</td>
            <td>Electrical Engineering</td>
            <td>(012)386 6116</td>
        </tr>
    </table>
</body>
</html>

4.3 <html>
<head>
    <title>Navigate</title>
</head>
<body>
    <h2>Navigate</h2>
    <a href='aboutCollege.html' target='main'>About Us</a><br />
    <a href='enquire.html' target='main'>Department</a>
</body>
</html>

4.4 <html>
<head>
    <title>Linking Frames and HTML</title>
</head>
<frameset cols='150,*'>
    <frame src='navigate.html' name='navigation' />
    <frame src='aboutCollege.html' name='main' />
</frameset>
<noframes>
    <body>
    </body>
    Your browser does not support Frames.
</noframes>
</html>


2010

2.2 ' Super Class source code and subclasses
Public MustInherit Class Recipe
    Public MustOverride Sub NewProcedure()

    Private MilkPrice As Decimal
    Private EggsPrice As Decimal
    Private FlowerPrice As Decimal
    Private SaltPrice As Decimal
    Private SugarPrice As Decimal

    ' Milk Method
    Public Property MilkPrices() As Decimal
        Get
            Return MilkPrice
        End Get
        Set(ByVal MilkValue As Decimal)
            MilkPrice = MilkValue
        End Set
    End Property

    ' Eggs Method
    Public Property EggsPrices() As Decimal
        Get
            Return EggsPrice
        End Get
        Set(ByVal EggsValue As Decimal)
            EggsPrice = EggsValue
        End Set
    End Property

    ' Flower Method
    Public Property FlowerPrices() As Decimal
        Get
            Return FlowerPrice
        End Get
        Set(ByVal FlowerValue As Decimal)
            FlowerPrice = FlowerValue
        End Set
    End Property

    ' Salt Method
    Public Property SaltPrices() As Decimal
        Get
            Return SaltPrice
        End Get
        Set(ByVal SaltValue As Decimal)
            SaltPrice = SaltValue
        End Set
    End Property

    ' Sugar Method
    Public Property SugarPrices() As Decimal
        Get
            Return SugarPrice
        End Get
        Set(ByVal SugarValue As Decimal)
            SugarPrice = SugarValue
        End Set
    End Property

    ' Constructor
    Public Sub New()
        MilkPrice = 0
        EggsPrice = 0
        FlowerPrice = 0
        SaltPrice = 0
        SugarPrice = 0
    End Sub

    ' Overloaded Constructor
    Public Sub New(ByVal FMilkPrice As Decimal, ByVal FEggsPrice As Decimal, ByVal FFlowerPrice As Decimal, ByVal FSaltPrice As Decimal, ByVal FSugarPrice As Decimal)
        MilkPrice = FMilkPrice
        EggsPrice = FEggsPrice
        FlowerPrice = FFlowerPrice
        SaltPrice = FSaltPrice
        SugarPrice = FSugarPrice
    End Sub

    ' Overridable Function
    Public Overridable Function PricePerPerson(ByVal Catered As Integer) As Decimal
        Dim Total As Decimal
        Total = (MilkPrice + EggsPrice + FlowerPrice + SaltPrice + SugarPrice) * Catered
        Return Total
    End Function
End Class

2.3 ' Sub Class Baked
Class Baked
    Inherits Recipe

    Public Overrides Sub NewProcedure()
    End Sub

    Public Property BakingTime As Decimal
    Public Property Temperature As Decimal

    Public Property BakingTimes() As Decimal
        Get
            Return BakingTime
        End Get
        Set(ByVal BakeValue As Decimal)
            BakingTime = BakeValue
        End Set
    End Property

    Public Property Temps() As Decimal
        Get
            Return Temperature
        End Get
        Set(ByVal TempValue As Decimal)
            Temperature = TempValue
        End Set
    End Property

    ' Overloaded Constructor
    Public Sub New(ByVal FMilkPrice As Decimal, ByVal FEggsPrice As Decimal, ByVal FFlowerPrice As Decimal, ByVal FSaltPrice As Decimal, ByVal FSugarPrice As Decimal, ByVal BakeTime As Decimal, ByVal BakeTemperature As Decimal)
        MyBase.New(FMilkPrice, FEggsPrice, FFlowerPrice, FSaltPrice, FSugarPrice)
        BakingTime = BakeTime
        Temperature = BakeTemperature
    End Sub

    Public Overrides Function PricePerPerson(ByVal Catered As Integer) As Decimal
        Return (MyBase.PricePerPerson(Catered) * 1.5 + (BakingTime / 30) + (Temperature / 90))
    End Function
End Class

' Sub Class Freeze
Class Freeze
    Inherits Recipe

    Public Overrides Sub NewProcedure()
    End Sub

    Public Property FridgeTime As Decimal

    Public Property FridgeTimes() As Decimal
        Get
            Return FridgeTime
        End Get
        Set(ByVal FridgeValue As Decimal)
            FridgeTime = FridgeValue
        End Set
    End Property

    ' Overloaded Constructor
    Public Sub New(ByVal FMilkPrice As Decimal, ByVal FEggsPrice As Decimal, ByVal FFlowerPrice As Decimal, ByVal FSaltPrice As Decimal, ByVal FSugarPrice As Decimal, ByVal FreezeTime As Decimal)
        MyBase.New(FMilkPrice, FEggsPrice, FFlowerPrice, FSaltPrice, FSugarPrice)
        FridgeTime = FreezeTime
    End Sub

    Public Overrides Function PricePerPerson(ByVal Catered As Integer) As Decimal
        Return (MyBase.PricePerPerson(Catered) * 1.2 + FridgeTime / 30)
    End Function
End Class

2.4 ' Form Class
Public Class Form1
    Private Sub btnExit_Click(...) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnClear_Click(...) Handles btnClear.Click
        txtMilkPrice.Clear()
        txtEggsPrice.Clear()
        txtFlowerPrice.Clear()
        txtSaltPrice.Clear()
        txtSugarPrice.Clear()
        txtBakeFreezetime.Clear()
        txtBakeTemp.Clear()
        txtCatered.Clear()
        lblResult.Text = ""
    End Sub

    Private Sub btnCalculate_Click(...) Handles btnCalculate.Click
        Dim MyBaked As Baked
        Dim MyFreeze As Freeze
        Dim Amount As Integer
        Dim MilkPrice, EggsPrice, FlowerPrice, SaltPrice, SugarPrice, BakeFreezetime, BakeTemp As Decimal

        MilkPrice = CDec(txtMilkPrice.Text)
        EggsPrice = CDec(txtEggsPrice.Text)
        FlowerPrice = CDec(txtFlowerPrice.Text)
        SaltPrice = CDec(txtSaltPrice.Text)
        SugarPrice = CDec(txtSugarPrice.Text)
        BakeFreezetime = CDec(txtBakeFreezetime.Text)
        BakeTemp = CDec(txtBakeTemp.Text)
        Amount = CDec(txtCatered.Text)

        If radBaked.Checked Then
            MyBaked = New Baked(MilkPrice, EggsPrice, FlowerPrice, SaltPrice, SugarPrice, BakeFreezetime, BakeTemp)
            MyBaked.MilkPrices = MilkPrice
            MyBaked.EggsPrices = EggsPrice
            MyBaked.FlowerPrices = FlowerPrice
            MyBaked.SaltPrices = SaltPrice
            MyBaked.BakingTimes = BakeFreezetime
            MyBaked.Temps = BakeTemp

            lblResult.Text = "The Price for " & txtCatered.Text & " people = " & (MyBaked.PricePerPerson(Amount) + (MyBaked.PricePerPerson(Amount) * 14 / 100)).ToString("C")
        End If

        If radFreeze.Checked Then
            MyFreeze = New Freeze(MilkPrice, EggsPrice, FlowerPrice, SaltPrice, SugarPrice, BakeFreezetime)
            MyFreeze.MilkPrices = MilkPrice
            MyFreeze.EggsPrices = EggsPrice
            MyFreeze.FlowerPrices = FlowerPrice
            MyFreeze.SaltPrices = SaltPrice
            MyFreeze.FridgeTimes = BakeFreezetime

            lblResult.Text = "The Price for " & txtCatered.Text & " people = " & (MyFreeze.PricePerPerson(Amount) + (MyFreeze.PricePerPerson(Amount) * 14 / 100)).ToString("C")
        End If
    End Sub
End Class

3.4.1 Public Class Form1
    EXAM NUMBER: 4645645645 ' (Question 3.4.7)

    Private Sub Form1_Load(…) Handles MyBase.Load ' (Question 3.3.1) and 3.3.2
        ' This code shows that the student could connect to OLEDB
        Me.EmployeeTableAdapter.Fill(Me.EmployeeDataSet.Employee)
    End Sub

    Private Sub btnClear_Click(…) Handles btnClear.Click
        txtEmployeeCode.Clear()
        txtInitials.Clear()
        txtSurname.Clear()
        txtGender.Clear()
        txtContact.Clear() ' (Question 3.4.2)
        txtPosition.Clear()
        txtGrossSalary.Clear()
        txtDeductions.Clear()
        lblSalary.Text = ""
    End Sub

    Private Sub btnExit_Click(…) Handles btnExit.Click
        Me.Close() ' (Question 3.4.3)
    End Sub

    Private Sub btnCalcSalary_Click(…) Handles btnCalcSalary.Click
        Dim Salary As Decimal ' (Question 3.4.1)
        Dim GrossSalary As Decimal
        Dim Deductions As Decimal

        GrossSalary = CDec(txtGrossSalary.Text)
        Deductions = CDec(txtDeductions.Text)
        Salary = GrossSalary - Deductions
        lblSalary.Text = "Employee Salary = " & Salary.ToString("C")
    End Sub

    Private Sub btnSaveChanges_Click(…) Handles btnSaveChanges.Click
        Try ' (Question 3.4.6)
            Me.Validate()
            Me.EmployeeBindingSource.EndEdit()
            Me.EmployeeTableAdapter.Update(Me.EmployeeDataSet.Employee)
            Me.EmployeeDataSet.AcceptChanges()
        Catch ex As Exception
            MsgBox("Update NOT successful")
        End Try
    End Sub
End Class