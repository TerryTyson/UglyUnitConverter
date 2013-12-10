Class MainWindow

    ' Initialize the value objects.
    Dim linearValue As New UglyUnitConversions.ValueLinear
    Dim areaValue As New UglyUnitConversions.ValueArea
    Dim temperatureValue As New UglyUnitConversions.ValueTemperature
    'Dim angularValue As New UglyUnitConversions.ValueAngular

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ShowAbout()
        Dim aboutBox As New AboutBox1
        aboutBox.StartPosition = Forms.FormStartPosition.CenterScreen
        aboutBox.Show()
    End Sub

    Dim isUpdating As Boolean = False

    ''' <summary>
    ''' Update all the linear unit text boxes based on the given magnitude and unit.
    ''' </summary>
    ''' <param name="givenMagnitude">The given numeric magnitude.</param>
    ''' <param name="givenUnit">The given unit.</param>
    ''' <remarks></remarks>
    Private Sub UpdateLinearTextBoxes(ByVal givenMagnitude As Double, ByVal givenUnit As UglyUnitConversions.ValueLinear.units)
        isUpdating = True
        linearValue.ValueMagnitude = givenMagnitude
        linearValue.ValueUnit = givenUnit
        If givenUnit <> UglyUnitConversions.ValueLinear.units.inches Then
            txtInches.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.inches)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.feet Then
            txtFeet.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.feet)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.yards Then
            txtYards.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.yards)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.miles Then
            txtMiles.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.miles)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.millimeters Then
            txtMillimeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.millimeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.centimeters Then
            txtCentimeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.centimeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.decimeters Then
            txtDecimeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.decimeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.meters Then
            txtMeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.meters)
        End If
        If givenUnit <> UglyUnitConversions.ValueLinear.units.kilometers Then
            txtKilometers.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.kilometers)
        End If
        isUpdating = False
    End Sub

    ''' <summary>
    ''' Update all the area unit text boxes based on the given magnitude and unit.
    ''' </summary>
    ''' <param name="givenMagnitude">The given numeric magnitude.</param>
    ''' <param name="givenUnit">The given unit.</param>
    ''' <remarks></remarks>
    Private Sub UpdateAreaTextBoxes(ByVal givenMagnitude As Double, ByVal givenUnit As UglyUnitConversions.ValueArea.units)
        isUpdating = True
        areaValue.ValueMagnitude = givenMagnitude
        areaValue.ValueUnit = givenUnit
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareInches Then
            txtSquareInches.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareInches)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareFeet Then
            txtSquareFeet.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareFeet)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareYards Then
            txtSquareYards.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareYards)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareMiles Then
            txtSquareMiles.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareMiles)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.acres Then
            txtAcres.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.acres)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareMillimeters Then
            txtSquareMillimeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareMillimeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareCentimeters Then
            txtSquareCentimeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareCentimeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareDecimeters Then
            txtSquareDecimeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareDecimeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareMeters Then
            txtSquareMeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareMeters)
        End If
        If givenUnit <> UglyUnitConversions.ValueArea.units.squareKilometers Then
            txtSquareKilometers.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareKilometers)
        End If
        isUpdating = False
    End Sub

    ''' <summary>
    ''' Update all the temperature unit text boxes based on the given magnitude and unit.
    ''' </summary>
    ''' <param name="givenMagnitude">The given numeric magnitude.</param>
    ''' <param name="givenUnit">The given unit.</param>
    ''' <remarks></remarks>
    Private Sub UpdateTemperatureTextBoxes(ByVal givenMagnitude As Double, ByVal givenUnit As UglyUnitConversions.ValueTemperature.units)
        isUpdating = True
        temperatureValue.ValueMagnitude = givenMagnitude
        temperatureValue.ValueUnit = givenUnit
        If givenUnit <> UglyUnitConversions.ValueTemperature.units.fahrenheit Then
            txtFahrenheit.Text = temperatureValue.GetValue(UglyUnitConversions.ValueTemperature.units.fahrenheit)
        End If
        If givenUnit <> UglyUnitConversions.ValueTemperature.units.celsius Then
            txtCelsius.Text = temperatureValue.GetValue(UglyUnitConversions.ValueTemperature.units.celsius)
        End If
        If givenUnit <> UglyUnitConversions.ValueTemperature.units.kelvin Then
            txtKelvin.Text = temperatureValue.GetValue(UglyUnitConversions.ValueTemperature.units.kelvin)
        End If
        isUpdating = False
    End Sub

    ' ########## LINEAR INCHES ##########
    Private Sub txtInches_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtInches.MouseDoubleClick
        txtInches.SelectAll()
    End Sub
    Private Sub txtInches_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtInches.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtInches.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(enteredValue) Then
            UpdateLinearTextBoxes(txtInches.Text, UglyUnitConversions.ValueLinear.units.inches)
        End If
    End Sub

    ' ########## LINEAR FEET ##########
    Private Sub txtFeet_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtFeet.MouseDoubleClick
        txtFeet.SelectAll()
    End Sub
    Private Sub txtFeet_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtFeet.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtFeet.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(enteredValue) Then
            UpdateLinearTextBoxes(txtFeet.Text, UglyUnitConversions.ValueLinear.units.feet)
        End If
    End Sub

    ' ########## LINEAR YARDS ##########
    Private Sub txtYards_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtYards.MouseDoubleClick
        txtYards.SelectAll()
    End Sub
    Private Sub txtYards_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtYards.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtYards.Text) Then
            UpdateLinearTextBoxes(txtYards.Text, UglyUnitConversions.ValueLinear.units.yards)
        End If
    End Sub

    ' ########## LINEAR MILES ##########
    Private Sub txtMiles_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMiles.MouseDoubleClick
        txtMiles.SelectAll()
    End Sub
    Private Sub txtMiles_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtMiles.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtMiles.Text) Then
            UpdateLinearTextBoxes(txtMiles.Text, UglyUnitConversions.ValueLinear.units.miles)
        End If
    End Sub

    ' ########## LINEAR MILLIMETERS ##########
    Private Sub txtMillimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMillimeters.MouseDoubleClick
        txtMillimeters.SelectAll()
    End Sub
    Private Sub txtMillimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtMillimeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtMillimeters.Text) Then
            UpdateLinearTextBoxes(txtMillimeters.Text, UglyUnitConversions.ValueLinear.units.millimeters)
        End If
    End Sub

    ' ########## LINEAR CENTIMETERS ##########
    Private Sub txtCentimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtCentimeters.MouseDoubleClick
        txtCentimeters.SelectAll()
    End Sub
    Private Sub txtCentimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtCentimeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtCentimeters.Text) Then
            UpdateLinearTextBoxes(txtCentimeters.Text, UglyUnitConversions.ValueLinear.units.centimeters)
        End If
    End Sub

    ' ########## LINEAR DECIMETERS ##########
    Private Sub txtDecimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtDecimeters.MouseDoubleClick
        txtDecimeters.SelectAll()
    End Sub
    Private Sub txtDecimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtDecimeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtDecimeters.Text) Then
            UpdateLinearTextBoxes(txtDecimeters.Text, UglyUnitConversions.ValueLinear.units.decimeters)
        End If
    End Sub

    ' ########## LINEAR METERS ##########
    Private Sub txtMeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMeters.MouseDoubleClick
        txtMeters.SelectAll()
    End Sub
    Private Sub txtMeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtMeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtMeters.Text) Then
            UpdateLinearTextBoxes(txtMeters.Text, UglyUnitConversions.ValueLinear.units.meters)
        End If
    End Sub

    ' ########## LINEAR KILOMETERS ##########
    Private Sub txtKilometers_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtKilometers.MouseDoubleClick
        txtKilometers.SelectAll()
    End Sub
    Private Sub txtKilometers_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtKilometers.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtKilometers.Text) Then
            UpdateLinearTextBoxes(txtKilometers.Text, UglyUnitConversions.ValueLinear.units.kilometers)
        End If
    End Sub

    ' ########## SQUARE INCHES ##########
    Private Sub txtSquareInches_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareInches.MouseDoubleClick
        txtSquareInches.SelectAll()
    End Sub
    Private Sub txtSquareInches_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareInches.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareInches.Text) Then
            UpdateAreaTextBoxes(txtSquareInches.Text, UglyUnitConversions.ValueArea.units.squareInches)
        End If
    End Sub

    ' ########## SQUARE FEET ##########
    Private Sub txtSquareFeet_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareFeet.MouseDoubleClick
        txtSquareFeet.SelectAll()
    End Sub
    Private Sub txtSquareFeet_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareFeet.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareFeet.Text) Then
            UpdateAreaTextBoxes(txtSquareFeet.Text, UglyUnitConversions.ValueArea.units.squareFeet)
        End If
    End Sub

    ' ########## SQUARE YARDS ##########
    Private Sub txtSquareYards_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareYards.MouseDoubleClick
        txtSquareYards.SelectAll()
    End Sub
    Private Sub txtSquareYards_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareYards.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareYards.Text) Then
            UpdateAreaTextBoxes(txtSquareYards.Text, UglyUnitConversions.ValueArea.units.squareYards)
        End If
    End Sub

    ' ########## SQUARE MILES ##########
    Private Sub txtSquareMiles_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareMiles.MouseDoubleClick
        txtSquareMiles.SelectAll()
    End Sub
    Private Sub txtSquareMiles_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareMiles.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareMiles.Text) Then
            UpdateAreaTextBoxes(txtSquareMiles.Text, UglyUnitConversions.ValueArea.units.squareMiles)
        End If
    End Sub

    ' ########## ACRES ##########
    Private Sub txtAcres_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtAcres.MouseDoubleClick
        txtAcres.SelectAll()
    End Sub
    Private Sub txtAcres_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtAcres.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtAcres.Text) Then
            UpdateAreaTextBoxes(txtAcres.Text, UglyUnitConversions.ValueArea.units.acres)
        End If
    End Sub

    ' ########## SQUARE MILLIMETERS ##########
    Private Sub txtSquareMillimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareMillimeters.MouseDoubleClick
        txtSquareMillimeters.SelectAll()
    End Sub
    Private Sub txtSquareMillimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareMillimeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareMillimeters.Text) Then
            UpdateAreaTextBoxes(txtSquareMillimeters.Text, UglyUnitConversions.ValueArea.units.squareMillimeters)
        End If
    End Sub

    ' ########## SQUARE CENTIMETERS ##########
    Private Sub txtSquareCentimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareCentimeters.MouseDoubleClick
        txtSquareCentimeters.SelectAll()
    End Sub
    Private Sub txtSquareCentimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareCentimeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareCentimeters.Text) Then
            UpdateAreaTextBoxes(txtSquareCentimeters.Text, UglyUnitConversions.ValueArea.units.squareCentimeters)
        End If
    End Sub

    ' ########## SQUARE DECIMETERS ##########
    Private Sub txtSquareDecimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareDecimeters.MouseDoubleClick
        txtSquareDecimeters.SelectAll()
    End Sub
    Private Sub txtSquareDecimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareDecimeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareDecimeters.Text) Then
            UpdateAreaTextBoxes(txtSquareDecimeters.Text, UglyUnitConversions.ValueArea.units.squareDecimeters)
        End If
    End Sub

    ' ########## SQUARE METERS ##########
    Private Sub txtSquareMeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareMeters.MouseDoubleClick
        txtSquareMeters.SelectAll()
    End Sub
    Private Sub txtSquareMeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareMeters.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareMeters.Text) Then
            UpdateAreaTextBoxes(txtSquareMeters.Text, UglyUnitConversions.ValueArea.units.squareMeters)
        End If
    End Sub

    ' ########## SQUARE KILOMETERS ##########
    Private Sub txtSquareKilometers_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareKilometers.MouseDoubleClick
        txtSquareKilometers.SelectAll()
    End Sub
    Private Sub txtSquareKilometers_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareKilometers.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtSquareKilometers.Text) Then
            UpdateAreaTextBoxes(txtSquareKilometers.Text, UglyUnitConversions.ValueArea.units.squareKilometers)
        End If
    End Sub

    ' ########## FAHRENHEIT ##########
    Private Sub txtFahrenheit_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtFahrenheit.MouseDoubleClick
        txtFahrenheit.SelectAll()
    End Sub
    Private Sub txtFahrenheit_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtFahrenheit.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtFahrenheit.Text) Then
            UpdateTemperatureTextBoxes(txtFahrenheit.Text, UglyUnitConversions.ValueTemperature.units.fahrenheit)
        End If
    End Sub

    ' ########## CELSIUS ##########
    Private Sub txtCelsius_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtCelsius.MouseDoubleClick
        txtCelsius.SelectAll()
    End Sub
    Private Sub txtCelsius_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtCelsius.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtCelsius.Text) Then
            UpdateTemperatureTextBoxes(txtCelsius.Text, UglyUnitConversions.ValueTemperature.units.celsius)
        End If
    End Sub

    ' ########## KELVIN ##########
    Private Sub txtKelvin_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtKelvin.MouseDoubleClick
        txtKelvin.SelectAll()
    End Sub
    Private Sub txtKelvin_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtKelvin.TextChanged
        If isUpdating = True Then Exit Sub
        Dim enteredValue As String = txtYards.Text
        If Strings.Right(enteredValue, 1) = "." Then
            enteredValue = enteredValue & 0
        End If
        If IsNumeric(txtKelvin.Text) Then
            UpdateTemperatureTextBoxes(txtKelvin.Text, UglyUnitConversions.ValueTemperature.units.kelvin)
        End If
    End Sub

End Class
