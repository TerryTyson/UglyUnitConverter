''' <summary>
''' Class with properties and methods for working with angular value objects.
''' </summary>
''' <remarks></remarks>
Public Class ValueAngular

    Inherits Value

    ''' <summary>
    ''' The units currently available for output.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum units
        decimalDegrees
        radians
        mils
        degrees
        minutes
        seconds
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
            ' This is so that we can always change Minutes and Seconds and get a proper value.
            If _ValueUnit = units.decimalDegrees Then
                ValueDecimalDegrees = ValueMagnitude
            ElseIf _ValueUnit = units.radians Then
                ValueDecimalDegrees = (ValueMagnitude * 180 / Math.PI)
            ElseIf _ValueUnit = units.mils Then
                ValueDecimalDegrees = (ValueMagnitude * 0.05625)
            End If
        End Set
    End Property

    'Private _DecimalDegrees As Double
    ' ''' <summary>
    ' ''' The angle in Decimal Degrees.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property DecimalDegrees As Double
    '    Get
    '        Return _DecimalDegrees
    '    End Get
    '    Set(value As Double)
    '        _DecimalDegrees = value
    '    End Set
    'End Property

    'Private _Radians As Double
    ' ''' <summary>
    ' ''' The angle in Radians.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property Radians As Double
    '    Get
    '        Return _Radians
    '    End Get
    '    Set(value As Double)
    '        _Radians = value
    '    End Set
    'End Property

    'Private _Mils As Double
    ' ''' <summary>
    ' ''' The angle in Mils.
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Property Mils As Double
    '    Get
    '        Return _Mils
    '    End Get
    '    Set(value As Double)
    '        _Mils = value
    '    End Set
    'End Property

    Private _Degrees As Double
    ''' <summary>
    ''' The Degree portion of the angle when represented in Degrees, Minutes, and Seconds.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Degrees As Double
        Get
            Return _Degrees
        End Get
        Set(value As Double)
            _Degrees = value
        End Set
    End Property

    Private _Minutes As Double
    ''' <summary>
    ''' The Minutes portion of the angle when represented in Degrees, Minutes, and Seconds.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Minutes As Double
        Get
            Return _Minutes
        End Get
        Set(value As Double)
            _Minutes = value
        End Set
    End Property

    Private _Seconds As Double
    ''' <summary>
    ''' The Seconds portion of the angle when represented in Degrees, Minutes, and Seconds.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Seconds As Double
        Get
            Return _Seconds
        End Get
        Set(value As Double)
            _Seconds = value
        End Set
    End Property

    Public ValueDecimalDegrees As Double = 0

    ''' <summary>
    ''' Converts Degrees, Minutes, and Seconds into decimal and stores in ValueDecimalDegrees.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AddDegMinSec()
        ValueDecimalDegrees = Degrees + (Minutes / 60) + ((Seconds / 60) / 60)
    End Sub

    ''' <summary>
    ''' Returns the magintude of the value in the unit given.
    ''' </summary>
    ''' <param name="returnUnit">The unit for the return value.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValue(ByVal returnUnit As units) As Double
        Select Case _ValueUnit
            Case units.decimalDegrees
                Select Case returnUnit
                    Case units.decimalDegrees
                        Return (ValueMagnitude)
                    Case units.radians
                        Return Math.PI * ValueMagnitude / 180
                    Case units.mils
                        Return (ValueMagnitude / 0.05625)
                    Case units.degrees
                        Return Math.Truncate(ValueMagnitude)
                    Case units.minutes
                        Return Math.Truncate((ValueMagnitude - Math.Truncate(ValueMagnitude)) * 60)
                    Case units.seconds
                        Dim minutes As Double = (ValueMagnitude - Math.Truncate(ValueMagnitude)) * 60
                        Return (minutes - Math.Truncate(minutes)) * 60
                    Case Else
                        Return 0
                End Select
            Case units.radians
                ValueDecimalDegrees = (ValueMagnitude * 180 / Math.PI)
                Select Case returnUnit
                    Case units.decimalDegrees
                        Return ValueDecimalDegrees
                    Case units.radians
                        Return (ValueMagnitude)
                    Case units.mils
                        Return (ValueMagnitude / 0.000981748)
                    Case units.degrees
                        Return Math.Truncate(ValueDecimalDegrees)
                    Case units.minutes
                        Return Math.Truncate((ValueDecimalDegrees - Math.Truncate(ValueDecimalDegrees)) * 60)
                    Case units.seconds
                        Dim minutes As Double = (ValueDecimalDegrees - Math.Truncate(ValueDecimalDegrees)) * 60
                        Return (minutes - Math.Truncate(minutes)) * 60
                    Case Else
                        Return 0
                End Select
            Case units.mils
                ValueDecimalDegrees = (ValueMagnitude * 0.05625)
                Select Case returnUnit
                    Case units.decimalDegrees
                        Return ValueDecimalDegrees
                    Case units.radians
                        Return (ValueMagnitude * 0.000981748)
                    Case units.mils
                        Return (ValueMagnitude)
                    Case units.degrees
                        Return Math.Truncate(ValueDecimalDegrees)
                    Case units.minutes
                        Return Math.Truncate((ValueDecimalDegrees - Math.Truncate(ValueDecimalDegrees)) * 60)
                    Case units.seconds
                        Dim minutes As Double = (ValueDecimalDegrees - Math.Truncate(ValueDecimalDegrees)) * 60
                        Return (minutes - Math.Truncate(minutes)) * 60
                    Case Else
                        Return 0
                End Select
            Case units.degrees, units.minutes, units.seconds
                Select Case returnUnit
                    Case units.decimalDegrees
                        Return ValueDecimalDegrees
                    Case units.radians
                        Return (ValueDecimalDegrees * 0.000981748)
                    Case units.mils
                        Return (ValueDecimalDegrees / 0.05625)
                    Case units.degrees
                        'Return Math.Truncate(ValueDecimalDegrees)
                        Return Degrees
                    Case units.minutes
                        'Return Math.Truncate((ValueDecimalDegrees - Math.Truncate(ValueDecimalDegrees)) * 60)
                        Return Minutes
                    Case units.seconds
                        'Dim minutes As Double = (ValueDecimalDegrees - Math.Truncate(ValueDecimalDegrees)) * 60
                        'Return (minutes - Math.Truncate(minutes)) * 60
                        Return Seconds
                    Case Else
                        Return 0
                End Select
            Case Else
                Return 0
        End Select
    End Function

End Class
