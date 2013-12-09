''' <summary>
''' Class with properties and methods for working with temperature value objects.
''' </summary>
''' <remarks></remarks>
Public Class ValueTemperature

    Inherits Value

    ''' <summary>
    ''' The units currently available for output.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum units
        fahrenheit
        celsius
        kelvin
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
            Case units.fahrenheit
                Select Case returnUnit
                    Case units.fahrenheit
                        Return (ValueMagnitude)
                    Case units.celsius
                        Return ((ValueMagnitude - 32) * (5 / 9))
                    Case units.kelvin
                        Return ((ValueMagnitude - 32) * (5 / 9)) + 273.15
                    Case Else
                        Return 0
                End Select
            Case units.celsius
                Select Case returnUnit
                    Case units.fahrenheit
                        Return ((ValueMagnitude * (9 / 5)) + 32)
                    Case units.celsius
                        Return (ValueMagnitude)
                    Case units.kelvin
                        Return (ValueMagnitude - 273.15)
                    Case Else
                        Return 0
                End Select
            Case units.kelvin
                Select Case returnUnit
                    Case units.fahrenheit
                        Return (((ValueMagnitude + 273.15) * (9 / 5)) + 32)
                    Case units.celsius
                        Return (ValueMagnitude + 273.15)
                    Case units.kelvin
                        Return (ValueMagnitude)
                    Case Else
                        Return 0
                End Select
            Case Else
                Return 0
        End Select
    End Function

End Class
