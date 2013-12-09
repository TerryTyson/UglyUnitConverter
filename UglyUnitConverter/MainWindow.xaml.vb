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
        txtInches.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.inches)
        txtFeet.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.feet)
        txtYards.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.yards)
        txtMiles.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.miles)
        txtMillimeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.millimeters)
        txtCentimeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.centimeters)
        txtDecimeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.decimeters)
        txtMeters.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.meters)
        txtKilometers.Text = linearValue.GetValue(UglyUnitConversions.ValueLinear.units.kilometers)
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
        txtSquareInches.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareInches)
        txtSquareFeet.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareFeet)
        txtSquareYards.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareYards)
        txtSquareMiles.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareMiles)
        txtAcres.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.acres)
        txtSquareMillimeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareMillimeters)
        txtSquareCentimeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareCentimeters)
        txtSquareDecimeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareDecimeters)
        txtSquareMeters.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareMeters)
        txtSquareKilometers.Text = areaValue.GetValue(UglyUnitConversions.ValueArea.units.squareKilometers)
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
        txtFahrenheit.Text = temperatureValue.GetValue(UglyUnitConversions.ValueTemperature.units.fahrenheit)
        txtCelsius.Text = temperatureValue.GetValue(UglyUnitConversions.ValueTemperature.units.celsius)
        txtKelvin.Text = temperatureValue.GetValue(UglyUnitConversions.ValueTemperature.units.kelvin)
        isUpdating = False
    End Sub

    ' ########## LINEAR INCHES ##########
    Private Sub txtInches_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtInches.MouseDoubleClick
        txtInches.SelectAll()
    End Sub
    Private Sub txtInches_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtInches.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtInches.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtInches.Text, UglyUnitConversions.ValueLinear.units.inches)
        End If
    End Sub

    ' ########## LINEAR FEET ##########
    Private Sub txtFeet_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtFeet.MouseDoubleClick
        txtFeet.SelectAll()
    End Sub
    Private Sub txtFeet_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtFeet.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtFeet.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtFeet.Text, UglyUnitConversions.ValueLinear.units.feet)
        End If
    End Sub

    ' ########## LINEAR YARDS ##########
    Private Sub txtYards_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtYards.MouseDoubleClick
        txtYards.SelectAll()
    End Sub
    Private Sub txtYards_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtYards.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtYards.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtYards.Text, UglyUnitConversions.ValueLinear.units.yards)
        End If
    End Sub

    ' ########## LINEAR MILES ##########
    Private Sub txtMiles_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMiles.MouseDoubleClick
        txtMiles.SelectAll()
    End Sub
    Private Sub txtMiles_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtMiles.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtMiles.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtMiles.Text, UglyUnitConversions.ValueLinear.units.miles)
        End If
    End Sub

    ' ########## LINEAR MILLIMETERS ##########
    Private Sub txtMillimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMillimeters.MouseDoubleClick
        txtMillimeters.SelectAll()
    End Sub
    Private Sub txtMillimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtMillimeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtMillimeters.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtMillimeters.Text, UglyUnitConversions.ValueLinear.units.millimeters)
        End If
    End Sub

    ' ########## LINEAR CENTIMETERS ##########
    Private Sub txtCentimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtCentimeters.MouseDoubleClick
        txtCentimeters.SelectAll()
    End Sub
    Private Sub txtCentimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtCentimeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtCentimeters.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtCentimeters.Text, UglyUnitConversions.ValueLinear.units.centimeters)
        End If
    End Sub

    ' ########## LINEAR DECIMETERS ##########
    Private Sub txtDecimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtDecimeters.MouseDoubleClick
        txtDecimeters.SelectAll()
    End Sub
    Private Sub txtDecimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtDecimeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtDecimeters.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtDecimeters.Text, UglyUnitConversions.ValueLinear.units.decimeters)
        End If
    End Sub

    ' ########## LINEAR METERS ##########
    Private Sub txtMeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMeters.MouseDoubleClick
        txtMeters.SelectAll()
    End Sub
    Private Sub txtMeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtMeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtMeters.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtMeters.Text, UglyUnitConversions.ValueLinear.units.meters)
        End If
    End Sub

    ' ########## LINEAR KILOMETERS ##########
    Private Sub txtKilometers_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtKilometers.MouseDoubleClick
        txtKilometers.SelectAll()
    End Sub
    Private Sub txtKilometers_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtKilometers.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtKilometers.Text) Then
            Exit Sub
        Else
            UpdateLinearTextBoxes(txtKilometers.Text, UglyUnitConversions.ValueLinear.units.kilometers)
        End If
    End Sub

    ' ########## SQUARE INCHES ##########
    Private Sub txtSquareInches_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareInches.MouseDoubleClick
        txtSquareInches.SelectAll()
    End Sub
    Private Sub txtSquareInches_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareInches.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareInches.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareInches.Text, UglyUnitConversions.ValueArea.units.squareInches)
        End If
    End Sub

    ' ########## SQUARE FEET ##########
    Private Sub txtSquareFeet_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareFeet.MouseDoubleClick
        txtSquareFeet.SelectAll()
    End Sub
    Private Sub txtSquareFeet_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareFeet.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareFeet.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareFeet.Text, UglyUnitConversions.ValueArea.units.squareFeet)
        End If
    End Sub

    ' ########## SQUARE YARDS ##########
    Private Sub txtSquareYards_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareYards.MouseDoubleClick
        txtSquareYards.SelectAll()
    End Sub
    Private Sub txtSquareYards_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareYards.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareYards.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareYards.Text, UglyUnitConversions.ValueArea.units.squareYards)
        End If
    End Sub

    ' ########## SQUARE MILES ##########
    Private Sub txtSquareMiles_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareMiles.MouseDoubleClick
        txtSquareMiles.SelectAll()
    End Sub
    Private Sub txtSquareMiles_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareMiles.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareMiles.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareMiles.Text, UglyUnitConversions.ValueArea.units.squareMiles)
        End If
    End Sub

    ' ########## ACRES ##########
    Private Sub txtAcres_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtAcres.MouseDoubleClick
        txtAcres.SelectAll()
    End Sub
    Private Sub txtAcres_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtAcres.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtAcres.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtAcres.Text, UglyUnitConversions.ValueArea.units.acres)
        End If
    End Sub

    ' ########## SQUARE MILLIMETERS ##########
    Private Sub txtSquareMillimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareMillimeters.MouseDoubleClick
        txtSquareMillimeters.SelectAll()
    End Sub
    Private Sub txtSquareMillimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareMillimeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareMillimeters.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareMillimeters.Text, UglyUnitConversions.ValueArea.units.squareMillimeters)
        End If
    End Sub

    ' ########## SQUARE CENTIMETERS ##########
    Private Sub txtSquareCentimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareCentimeters.MouseDoubleClick
        txtSquareCentimeters.SelectAll()
    End Sub
    Private Sub txtSquareCentimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareCentimeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareCentimeters.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareCentimeters.Text, UglyUnitConversions.ValueArea.units.squareCentimeters)
        End If
    End Sub

    ' ########## SQUARE DECIMETERS ##########
    Private Sub txtSquareDecimeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareDecimeters.MouseDoubleClick
        txtSquareDecimeters.SelectAll()
    End Sub
    Private Sub txtSquareDecimeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareDecimeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareDecimeters.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareDecimeters.Text, UglyUnitConversions.ValueArea.units.squareDecimeters)
        End If
    End Sub

    ' ########## SQUARE METERS ##########
    Private Sub txtSquareMeters_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareMeters.MouseDoubleClick
        txtSquareMeters.SelectAll()
    End Sub
    Private Sub txtSquareMeters_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareMeters.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareMeters.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareMeters.Text, UglyUnitConversions.ValueArea.units.squareMeters)
        End If
    End Sub

    ' ########## SQUARE KILOMETERS ##########
    Private Sub txtSquareKilometers_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtSquareKilometers.MouseDoubleClick
        txtSquareKilometers.SelectAll()
    End Sub
    Private Sub txtSquareKilometers_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtSquareKilometers.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtSquareKilometers.Text) Then
            Exit Sub
        Else
            UpdateAreaTextBoxes(txtSquareKilometers.Text, UglyUnitConversions.ValueArea.units.squareKilometers)
        End If
    End Sub

    ' ########## FAHRENHEIT ##########
    Private Sub txtFahrenheit_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtFahrenheit.MouseDoubleClick
        txtFahrenheit.SelectAll()
    End Sub
    Private Sub txtFahrenheit_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtFahrenheit.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtFahrenheit.Text) Then
            Exit Sub
        Else
            UpdateTemperatureTextBoxes(txtFahrenheit.Text, UglyUnitConversions.ValueTemperature.units.fahrenheit)
        End If
    End Sub

    ' ########## CELSIUS ##########
    Private Sub txtCelsius_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtCelsius.MouseDoubleClick
        txtCelsius.SelectAll()
    End Sub
    Private Sub txtCelsius_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtCelsius.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtCelsius.Text) Then
            Exit Sub
        Else
            UpdateTemperatureTextBoxes(txtCelsius.Text, UglyUnitConversions.ValueTemperature.units.celsius)
        End If
    End Sub

    ' ########## KELVIN ##########
    Private Sub txtKelvin_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtKelvin.MouseDoubleClick
        txtKelvin.SelectAll()
    End Sub
    Private Sub txtKelvin_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtKelvin.TextChanged
        If isUpdating = True Then Exit Sub
        If Not IsNumeric(txtKelvin.Text) Then
            Exit Sub
        Else
            UpdateTemperatureTextBoxes(txtKelvin.Text, UglyUnitConversions.ValueTemperature.units.kelvin)
        End If
    End Sub

End Class
