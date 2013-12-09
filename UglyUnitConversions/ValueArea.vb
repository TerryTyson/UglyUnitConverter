''' <summary>
''' Class with properties and methods for working with area value objects.
''' </summary>
''' <remarks></remarks>
Public Class ValueArea

    Inherits Value

    ''' <summary>
    ''' The units currently available for output.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum units
        squareInches
        squareFeet
        squareYards
        squareMiles
        acres
        squareMillimeters
        squareCentimeters
        squareDecimeters
        squareMeters
        squareKilometers
    End Enum

    Private _ValueUnit As units
    ''' <summary>
    ''' The unit of measure for the value.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ValueUnit As units
        Get
            Return _ValueUnit
        End Get
        Set(value As units)
            _ValueUnit = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the magintude of the value in the unit given.
    ''' </summary>
    ''' <param name="returnUnit">The unit for the return value.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValue(ByVal returnUnit As units) As Double
        Select Case _ValueUnit
            Case units.squareInches
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude)
                    Case units.squareFeet
                        Return (ValueMagnitude / 144)
                    Case units.squareYards
                        Return (ValueMagnitude / 1296)
                    Case units.squareMiles
                        Return (ValueMagnitude / 4014489600)
                    Case units.acres
                        Return (ValueMagnitude / 6272640)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 645.16)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 6.4516)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 0.064516)
                    Case units.squareMeters
                        Return (ValueMagnitude * 0.00064516)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.00000000064516)
                    Case Else
                        Return 0
                End Select
            Case units.squareFeet
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude * 144)
                    Case units.squareFeet
                        Return (ValueMagnitude)
                    Case units.squareYards
                        Return (ValueMagnitude / 9)
                    Case units.squareMiles
                        Return (ValueMagnitude / 27878400)
                    Case units.acres
                        Return (ValueMagnitude / 43560)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 92.90304)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 9.290304)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 0.9290304)
                    Case units.squareMeters
                        Return (ValueMagnitude * 0.09290304)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.00009290304)
                    Case Else
                        Return 0
                End Select
            Case units.squareYards
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude * 1296)
                    Case units.squareFeet
                        Return (ValueMagnitude * 9)
                    Case units.squareYards
                        Return (ValueMagnitude)
                    Case units.squareMiles
                        Return (ValueMagnitude / 3097600)
                    Case units.acres
                        Return (ValueMagnitude / 4840)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 836127.36)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 8361.2736)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 83.612736)
                    Case units.squareMeters
                        Return (ValueMagnitude * 0.83612736)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.00000083612736)
                    Case Else
                        Return 0
                End Select
            Case units.squareMiles
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude * 4014489600)
                    Case units.squareFeet
                        Return (ValueMagnitude * 27878400)
                    Case units.squareYards
                        Return (ValueMagnitude * 3097600)
                    Case units.squareMiles
                        Return (ValueMagnitude)
                    Case units.acres
                        Return (ValueMagnitude * 640)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 2589988110340)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 25899881103.4)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 258998811.034)
                    Case units.squareMeters
                        Return (ValueMagnitude * 2589988.11034)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 2.58998811034)
                    Case Else
                        Return 0
                End Select
            Case units.acres
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude * 6272640)
                    Case units.squareFeet
                        Return (ValueMagnitude * 43560)
                    Case units.squareYards
                        Return (ValueMagnitude * 4840)
                    Case units.squareMiles
                        Return (ValueMagnitude / 640)
                    Case units.acres
                        Return (ValueMagnitude)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 4046856422.4)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 40468564.224)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 404685.64224)
                    Case units.squareMeters
                        Return (ValueMagnitude * 4046.8564224)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.0040468564224)
                    Case Else
                        Return 0
                End Select
            Case units.squareMillimeters
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude / 645.16)
                    Case units.squareFeet
                        Return (ValueMagnitude / 92903.04)
                    Case units.squareYards
                        Return (ValueMagnitude / 836127.36)
                    Case units.squareMiles
                        Return (ValueMagnitude / 2589988110340)
                    Case units.acres
                        Return (ValueMagnitude / 4046856422.4)
                    Case units.squareMillimeters
                        Return (ValueMagnitude)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 0.01)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 0.0001)
                    Case units.squareMeters
                        Return (ValueMagnitude * 0.000001)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.000000000001)
                    Case Else
                        Return 0
                End Select
            Case units.squareCentimeters
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude / 6.4516)
                    Case units.squareFeet
                        Return (ValueMagnitude / 929.0304)
                    Case units.squareYards
                        Return (ValueMagnitude / 8361.2736)
                    Case units.squareMiles
                        Return (ValueMagnitude / 25899881103.4)
                    Case units.acres
                        Return (ValueMagnitude / 40468564.224)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 100)
                    Case units.squareCentimeters
                        Return (ValueMagnitude)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 0.01)
                    Case units.squareMeters
                        Return (ValueMagnitude * 0.0001)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.0000000001)
                    Case Else
                        Return 0
                End Select
            Case units.squareDecimeters
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude / 0.064516)
                    Case units.squareFeet
                        Return (ValueMagnitude / 9.290304)
                    Case units.squareYards
                        Return (ValueMagnitude / 83.612736)
                    Case units.squareMiles
                        Return (ValueMagnitude / 258998811.034)
                    Case units.acres
                        Return (ValueMagnitude / 404685.64224)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 10000)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 100)
                    Case units.squareDecimeters
                        Return (ValueMagnitude)
                    Case units.squareMeters
                        Return (ValueMagnitude * 0.01)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.00000001)
                    Case Else
                        Return 0
                End Select
            Case units.squareMeters
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude / 0.00064516)
                    Case units.squareFeet
                        Return (ValueMagnitude / 0.09290304)
                    Case units.squareYards
                        Return (ValueMagnitude / 0.83612736)
                    Case units.squareMiles
                        Return (ValueMagnitude / 2589988.11034)
                    Case units.acres
                        Return (ValueMagnitude / 4046.8564224)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 1000000)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 10000)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 100)
                    Case units.squareMeters
                        Return (ValueMagnitude)
                    Case units.squareKilometers
                        Return (ValueMagnitude * 0.000001)
                    Case Else
                        Return 0
                End Select
            Case units.squareKilometers
                Select Case returnUnit
                    Case units.squareInches
                        Return (ValueMagnitude * 1550003100.01)
                    Case units.squareFeet
                        Return (ValueMagnitude * 10763910.4167)
                    Case units.squareYards
                        Return (ValueMagnitude * 1195990.0463)
                    Case units.squareMiles
                        Return (ValueMagnitude / 2.58998811034)
                    Case units.acres
                        Return (ValueMagnitude * 247.105381467)
                    Case units.squareMillimeters
                        Return (ValueMagnitude * 1000000000000)
                    Case units.squareCentimeters
                        Return (ValueMagnitude * 10000000000)
                    Case units.squareDecimeters
                        Return (ValueMagnitude * 100000000)
                    Case units.squareMeters
                        Return (ValueMagnitude * 1000000)
                    Case units.squareKilometers
                        Return (ValueMagnitude)
                    Case Else
                        Return 0
                End Select
            Case Else
                Return 0
        End Select
    End Function

End Class
