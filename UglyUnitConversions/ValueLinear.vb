''' <summary>
''' Class with properties and methods for working with linear value objects.
''' </summary>
''' <remarks></remarks>
Public Class ValueLinear

    Inherits Value

    ''' <summary>
    ''' The units currently available for output.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum units
        inches
        feet
        yards
        miles
        millimeters
        centimeters
        decimeters
        meters
        kilometers
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
            Case units.inches
                Select Case returnUnit
                    Case units.inches
                        Return ValueMagnitude
                    Case units.feet
                        Return (ValueMagnitude / 12)
                    Case units.yards
                        Return (ValueMagnitude / 36)
                    Case units.miles
                        Return (ValueMagnitude / 63360)
                    Case units.millimeters
                        Return (ValueMagnitude * 25.4)
                    Case units.centimeters
                        Return ((ValueMagnitude * 25.4) / 10)
                    Case units.decimeters
                        Return ((ValueMagnitude * 25.4) / 100)
                    Case units.meters
                        Return ((ValueMagnitude * 25.4) / 1000)
                    Case units.kilometers
                        Return ((ValueMagnitude * 25.4) / 1000000)
                    Case Else
                        Return 0
                End Select
            Case units.feet
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude * 12)
                    Case units.feet
                        Return ValueMagnitude
                    Case units.yards
                        Return (ValueMagnitude / 3)
                    Case units.miles
                        Return (ValueMagnitude / 5280)
                    Case units.millimeters
                        Return (ValueMagnitude * 304.8)
                    Case units.centimeters
                        Return ((ValueMagnitude * 304.8) / 10)
                    Case units.decimeters
                        Return ((ValueMagnitude * 304.8) / 100)
                    Case units.meters
                        Return ((ValueMagnitude * 304.8) / 1000)
                    Case units.kilometers
                        Return ((ValueMagnitude * 304.8) / 1000000)
                    Case Else
                        Return 0
                End Select
            Case units.yards
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude * 36)
                    Case units.feet
                        Return (ValueMagnitude * 3)
                    Case units.yards
                        Return ValueMagnitude
                    Case units.miles
                        Return (ValueMagnitude / 1760)
                    Case units.millimeters
                        Return (ValueMagnitude * 914.4)
                    Case units.centimeters
                        Return ((ValueMagnitude * 914.4) / 10)
                    Case units.decimeters
                        Return ((ValueMagnitude * 914.4) / 100)
                    Case units.meters
                        Return ((ValueMagnitude * 914.4) / 1000)
                    Case units.kilometers
                        Return ((ValueMagnitude * 914.4) / 1000000)
                    Case Else
                        Return 0
                End Select
            Case units.miles
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude * 63360)
                    Case units.feet
                        Return (ValueMagnitude * 5280)
                    Case units.yards
                        Return (ValueMagnitude * 1760)
                    Case units.miles
                        Return ValueMagnitude
                    Case units.millimeters
                        Return (ValueMagnitude * 1609344)
                    Case units.centimeters
                        Return ((ValueMagnitude * 1609344) / 10)
                    Case units.decimeters
                        Return ((ValueMagnitude * 1609344) / 100)
                    Case units.meters
                        Return ((ValueMagnitude * 1609344) / 1000)
                    Case units.kilometers
                        Return ((ValueMagnitude * 1609344) / 1000000)
                    Case Else
                        Return 0
                End Select
            Case units.millimeters
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude / 25.4)
                    Case units.feet
                        Return (ValueMagnitude / 304.8)
                    Case units.yards
                        Return (ValueMagnitude / 914.4)
                    Case units.miles
                        Return (ValueMagnitude / 1609344)
                    Case units.millimeters
                        Return ValueMagnitude
                    Case units.centimeters
                        Return (ValueMagnitude / 10)
                    Case units.decimeters
                        Return (ValueMagnitude / 100)
                    Case units.meters
                        Return (ValueMagnitude / 1000)
                    Case units.kilometers
                        Return (ValueMagnitude / 1000000)
                    Case Else
                        Return 0
                End Select
            Case units.centimeters
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude / 2.54)
                    Case units.feet
                        Return (ValueMagnitude / 30.48)
                    Case units.yards
                        Return (ValueMagnitude / 91.44)
                    Case units.miles
                        Return (ValueMagnitude / 160934.4)
                    Case units.millimeters
                        Return (ValueMagnitude * 10)
                    Case units.centimeters
                        Return ValueMagnitude
                    Case units.decimeters
                        Return (ValueMagnitude / 10)
                    Case units.meters
                        Return (ValueMagnitude / 100)
                    Case units.kilometers
                        Return (ValueMagnitude / 100000)
                    Case Else
                        Return 0
                End Select
            Case units.decimeters
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude / 0.254)
                    Case units.feet
                        Return (ValueMagnitude / 3.048)
                    Case units.yards
                        Return (ValueMagnitude / 9.144)
                    Case units.miles
                        Return (ValueMagnitude / 16093.44)
                    Case units.millimeters
                        Return (ValueMagnitude * 100)
                    Case units.centimeters
                        Return (ValueMagnitude * 10)
                    Case units.decimeters
                        Return ValueMagnitude
                    Case units.meters
                        Return (ValueMagnitude / 10)
                    Case units.kilometers
                        Return (ValueMagnitude / 10000)
                    Case Else
                        Return 0
                End Select
            Case units.meters
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude / 0.0254)
                    Case units.feet
                        Return (ValueMagnitude / 0.3048)
                    Case units.yards
                        Return (ValueMagnitude / 0.9144)
                    Case units.miles
                        Return (ValueMagnitude / 1609.344)
                    Case units.millimeters
                        Return (ValueMagnitude * 1000)
                    Case units.centimeters
                        Return (ValueMagnitude * 100)
                    Case units.decimeters
                        Return (ValueMagnitude * 10)
                    Case units.meters
                        Return ValueMagnitude
                    Case units.kilometers
                        Return (ValueMagnitude / 1000)
                    Case Else
                        Return 0
                End Select
            Case units.kilometers
                Select Case returnUnit
                    Case units.inches
                        Return (ValueMagnitude / 0.0000254)
                    Case units.feet
                        Return (ValueMagnitude / 0.0003048)
                    Case units.yards
                        Return (ValueMagnitude / 0.0009144)
                    Case units.miles
                        Return (ValueMagnitude / 1.609344)
                    Case units.millimeters
                        Return (ValueMagnitude * 1000000)
                    Case units.centimeters
                        Return (ValueMagnitude * 100000)
                    Case units.decimeters
                        Return (ValueMagnitude * 10000)
                    Case units.meters
                        Return (ValueMagnitude * 1000)
                    Case units.kilometers
                        Return ValueMagnitude
                    Case Else
                        Return 0
                End Select
            Case Else
                Return 0
        End Select
    End Function

End Class
