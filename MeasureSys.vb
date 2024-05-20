<CLSCompliantAttribute(True)> Public Class MeasureSystem
    Public Shared Event MeasureUnitChanged(ByVal NewUnit As enUniMis)
    Public Enum enUniMis
        micron = 0
        mm = 1
        inches = 2
        dmm = 3
        meters = 4
    End Enum
    Private myUserUnit As enUniMis
    Public Sub New()
        Try
            UserUnit = enUniMis.micron
        Catch ex As Exception
            MsgBox("MeasureSystem: " + ex.Message)
        End Try
    End Sub
    Public Property UserUnit() As enUniMis
        Get
            Return myUserUnit
        End Get
        Set(ByVal value As enUniMis)
            If (value = enUniMis.dmm) Or (value = enUniMis.inches) Or (value = enUniMis.micron) Or (value = enUniMis.mm) Or (value = enUniMis.meters) Then
                If myUserUnit <> value Then
                    myUserUnit = value
                    RaiseEvent MeasureUnitChanged(myUserUnit)
                End If
            End If
        End Set
    End Property
    Public Shared Function MicronToCustomUnit(ByVal Measure_micron As Double, ByVal CustomUnit As enUniMis, Optional ByVal Round As Boolean = False) As Double
        Dim retValue As Double = 0
        Try
            If CustomUnit = enUniMis.micron Then
                Return Measure_micron
            Else
                'Converto da micron a AracneInfo.BaseUnit ...
                Select Case CustomUnit
                    Case enUniMis.inches
                        ' 1 inch = 25400 micron ...
                        retValue = Measure_micron / 25400
                        If Round Then
                            'Risoluzione = inch/100
                            retValue = Int(retValue * 100) / 100
                        End If
                    Case enUniMis.micron
                        retValue = Measure_micron
                    Case enUniMis.mm
                        retValue = Measure_micron / 1000
                        If Round Then
                            'Risoluzione = mm/10
                            retValue = Int(retValue * 10) / 10
                        End If
                    Case enUniMis.meters
                        retValue = Measure_micron / 1000000
                        If Round Then
                            'Risoluzione = m/10
                            retValue = Int(retValue * 10) / 10
                        End If
                    Case enUniMis.dmm
                        retValue = Measure_micron / 100
                        If Round Then
                            'Risoluzione = dmm
                            retValue = Int(retValue)
                        End If
                End Select
                Return retValue
            End If
        Catch ex As Exception
            MsgBox("MeasureSystem: " + ex.Message)
            Return 0
        End Try
    End Function
    Public Shared Function MicronToCustomUnit(ByVal Measure_micron As Integer, ByVal CustomUnit As enUniMis, Optional ByVal Round As Boolean = False) As Integer
        Return CInt(MicronToCustomUnit(CDbl(Measure_micron), CustomUnit, Round))
    End Function
    Public Function MicronToUserUnit(ByVal Measure_micron As Double, Optional ByVal Round As Boolean = False) As Double
        Return MicronToCustomUnit(Measure_micron, UserUnit, Round)
    End Function
    Public Shared Function CustomUnitToMicron(ByVal MeasureValue As Double, ByVal CustomUnit As enUniMis) As Integer
        Try
            Dim retVal As Integer
            'Converto in micron ...
            Select Case CustomUnit
                Case enUniMis.inches
                    ' 1 inch = 25400 micron ...
                    retVal = 25400 * MeasureValue
                Case enUniMis.micron
                    retVal = MeasureValue
                Case enUniMis.meters
                    retVal = 1000000 * MeasureValue
                Case enUniMis.mm
                    retVal = 1000 * MeasureValue
                Case enUniMis.dmm
                    retVal = 100 * MeasureValue
            End Select
            Return retVal
        Catch ex As Exception
            MsgBox("MeasureSystem: " + ex.Message)
            Return 0
        End Try
    End Function
    Public Function UserUnitToMicron(ByVal MeasureValue As Double) As Integer
        Return CustomUnitToMicron(MeasureValue, UserUnit)
    End Function
    Public Function UserUnitDescription() As String
        Return UniMisDescription(UserUnit)
    End Function
    Public Sub FillComboWithAvailableUnits(ByVal cbMeasureUnit As System.Windows.Forms.ComboBox)
        Try
            cbMeasureUnit.Items.Clear()
            Dim myArray() As MeasureSystem.enUniMis = System.Enum.GetValues(GetType(MeasureSystem.enUniMis))
            For iCounter As Integer = 0 To myArray.Length - 1
                cbMeasureUnit.Items.Add(MeasureSystem.UniMisDescription(myArray(iCounter)))
            Next
            cbMeasureUnit.SelectedIndex = Me.UserUnit
        Catch ex As Exception
            MsgBox("MeasureSystem: " + ex.Message)
        End Try
    End Sub
    Public Shared Function UniMisDescription(ByVal UNIT As enUniMis) As String
        Try
            Select Case UNIT
                Case enUniMis.inches
                    Return "inches"
                Case enUniMis.micron
                    Return "micron"
                Case enUniMis.mm
                    Return "mm"
                Case enUniMis.meters
                    Return "m"
                Case enUniMis.dmm
                    Return "dmm"
                Case Else
                    Return ""
            End Select
        Catch ex As Exception
            MsgBox("MeasureSystem: " + ex.Message)
            Return ""
        End Try
    End Function
End Class