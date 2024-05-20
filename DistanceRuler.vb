Imports System.ComponentModel
'Imports Microsoft.VisualBasic

Public Delegate Sub CaptureFishedEventHandler(ByVal sender As Object, ByVal e As CaptureEventArgs)

Public Class CaptureEventArgs
    Inherits EventArgs

    Private ReadOnly _startPoint As Point
    Private ReadOnly _endPoint As Point

    Public Sub New(ByVal startPoint As Point, ByVal endPoint As Point)
        _startPoint = startPoint
        _endPoint = endPoint
    End Sub

    Public ReadOnly Property StartPoint() As Point
        Get
            Return _startPoint
        End Get
    End Property

    Public ReadOnly Property EndPoint() As Point
        Get
            Return _endPoint
        End Get
    End Property
End Class

Public Class DistanceRuler
    Public Event CaptureFinished As CaptureFishedEventHandler
    Private _mouseCaptured As Boolean
    Private _angle As Single
    Private _length As Single
    Private _origin As Point
    Private _last As Point

    Private _lineWidth As Integer = 5
    Private _compArray As Single() = New Single() {0.0, 0.16, 0.33, 0.66, 0.83, 1.0}

    Private myPictureBoxControl As ZRGPictureBoxControl
    Public Property PictureBoxControl() As ZRGPictureBoxControl
        Get
            Return myPictureBoxControl
        End Get
        Private Set(ByVal value As ZRGPictureBoxControl)
            myPictureBoxControl = value
        End Set
    End Property
    Private ReadOnly Property UnitOfMeasure() As MeasureSystem.enUniMis
        Get
            Return PictureBoxControl.UnitOfMeasure
        End Get
    End Property
    Public Sub New(ByVal pictureBox As ZRGPictureBoxControl)
        If pictureBox Is Nothing Then
            Throw New ArgumentNullException("pictureBox", "MouseCapture must be associated with a control.")
        End If

        myPictureBoxControl = pictureBox
    End Sub
    Public Property Backcolor() As Color
        Get
            Return myPictureBoxControl.BackgroundColor
        End Get
        Set(ByVal value As Color)
            myPictureBoxControl.BackgroundColor = value
        End Set
    End Property
    Public Property ForeColor() As Color
        Get
            Return myPictureBoxControl.ForeColor
        End Get
        Set(ByVal value As Color)
            myPictureBoxControl.ForeColor = value
        End Set
    End Property
    Public Property LineWidth() As Integer
        Get
            Return _lineWidth
        End Get
        Set(ByVal value As Integer)
            If _lineWidth < 1 Then
                Throw New ArgumentOutOfRangeException("LineWidth", value, "Line width must greater than or equal to one.")
            End If
            _lineWidth = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets an array values that specify a compound pen. A compound pen draws a compound line made up of parallel lines and spaces.
    ''' </summary>
    ''' <value>An array of single values.</value>
    ''' <remarks>A compound line is made up of alternating parallel lines and spaces of varying widths.
    ''' The values in the array specify the starting points of each component of the compound line 
    ''' relative to the pen's width. The first value in the array specifies where the first component 
    ''' (a line) begins as a fraction of the distance across the width of the pen. The second value in the 
    ''' array specifies the beginning of the next component (a space) as a fraction of the distance across 
    ''' the width of the pen. The final value in the array specifies where the last component ends.
    ''' Suppose you want a pen to draw two parallel lines where the width of the first line is 20 percent of 
    ''' the pen's width, the width of the space that separates the two lines is 50 percent of the pen' s width, 
    ''' and the width of the second line is 30 percent of the pen's width. Start by creating a Pen object and 
    ''' an array of real numbers. 
    ''' Set the compound array by passing the array with the values 0.0, 0.2, 0.7, and 1.0 to this property.
    ''' </remarks>
    Public Property LineCompoundArray() As Single()
        Get
            Return _compArray
        End Get
        Set(ByVal value As Single())
            For Each i As Single In value
                If i < 0 OrElse i > 1 Then
                    Throw New ArgumentOutOfRangeException("LineCompoundArray", i, "All elements in the compound array must be >=0 or <=1.")
                End If
            Next
            _compArray = value
        End Set
    End Property
    Private Function CvRadToDeg(ByVal RadAngle As Double) As Double
        'espresso in gradi, 0-360
        CvRadToDeg = RadAngle * (180 / (System.Math.Atan(1) * 4))
    End Function
    Public Shared Function CutDecimals(ByVal Value As Double, ByVal DesiredDecDigits As Integer) As Double
        Try
            If (Value = Double.NegativeInfinity OrElse Value = Double.PositiveInfinity) Then
                Return Value
            End If
            If DesiredDecDigits > 5 Then
                DesiredDecDigits = 5
            End If
            CutDecimals = CInt(Value * 10 ^ DesiredDecDigits) / (10 ^ DesiredDecDigits)
        Catch ex As Exception
            CutDecimals = Value
        End Try
    End Function
    Public Shared Function strCutDecimals(ByVal Value As Double, ByVal DesiredDecDigits As Integer) As String
        On Error Resume Next
        strCutDecimals = CStr(CutDecimals(Value, DesiredDecDigits))
    End Function
    Friend Sub Painting(ByVal GR As Graphics, Optional ByVal ScaleFactor As Double = 1.0)
        If Not _mouseCaptured Then
            Exit Sub
        End If

        Dim OriginCrossArmLength As Integer = 20


        Dim CurrentAngle As Double = _
           New SEGMENT(_origin.X, _origin.Y, _last.X, _last.Y).SegmentDirection
        CurrentAngle = CvRadToDeg(CurrentAngle)

        If CurrentAngle > 180 Then
            CurrentAngle = CurrentAngle - 360
        End If
        Dim Scale As Single = PictureBoxControl.ScaleFactor * UnitOfMeasureFactor
        Using wallpen As New Pen(myPictureBoxControl.ForeColor, 1)
            Dim midPoint As New Point
            midPoint.X = Math.Min(_origin.X, _last.X) + _
                ((Math.Max(_origin.X, _last.X) - Math.Min(_origin.X, _last.X)) / 2)
            midPoint.Y = Math.Min(_origin.Y, _last.Y) + _
                ((Math.Max(_origin.Y, _last.Y) - Math.Min(_origin.Y, _last.Y)) / 2)



            GR.DrawLine(wallpen, _origin.X - OriginCrossArmLength, _origin.Y, _
                        _origin.X + OriginCrossArmLength, _origin.Y)
            GR.DrawLine(wallpen, _origin.X, _origin.Y - OriginCrossArmLength, _
                        _origin.X, _origin.Y + OriginCrossArmLength)
            GR.DrawArc(wallpen, _origin.X - OriginCrossArmLength, _origin.Y - OriginCrossArmLength, _
                       2 * OriginCrossArmLength, 2 * OriginCrossArmLength, 0, CInt(-CurrentAngle))

            Using mx As New System.Drawing.Drawing2D.Matrix
                Using sf As New StringFormat()
                    Dim ls As String = ""
                    ls = Format(LineLength(_origin, _last, Scale) / ScaleFactor, "#.00") + _
                        "  (" + strCutDecimals(CurrentAngle, 1) + "°)"
                    Dim l As SizeF = GR.MeasureString(ls, myPictureBoxControl.Font, _
                                                      myPictureBoxControl.ClientSize, sf)
                    sf.LineAlignment = StringAlignment.Center
                    sf.Alignment = StringAlignment.Center
                    GR.DrawLine(wallpen, _origin, _last)
                    mx.Translate(midPoint.X, midPoint.Y)
                    mx.Rotate(Angle(_origin, _last))
                    GR.Transform = mx
                    Dim rt As New Rectangle(0, 0, l.Width, l.Height)
                    rt.Inflate(3, 3)
                    rt.Offset(-(l.Width / 2), -(l.Height / 2))
                    Using backBrush As New SolidBrush(myPictureBoxControl.BackgroundColor)
                        GR.FillEllipse(backBrush, rt)
                    End Using
                    Using foreBrush As New SolidBrush(myPictureBoxControl.ForeColor)
                        GR.DrawString(ls, myPictureBoxControl.Font, foreBrush, 0, 0, sf)
                    End Using
                End Using
            End Using
        End Using
    End Sub
    Friend Sub MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        _mouseCaptured = True
        _origin = e.Location
        _last = New Point(-1, -1)
    End Sub

    Friend Sub MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If _mouseCaptured Then
            Dim r As Rectangle = NormalizeRect(_origin, _last)
            r.Inflate(myPictureBoxControl.Font.Height, myPictureBoxControl.Font.Height)
            _last = e.Location
        End If
    End Sub

    Friend Sub MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        _mouseCaptured = False
        myPictureBoxControl.Invalidate()
        RaiseEvent CaptureFinished(Me, New CaptureEventArgs(_origin, e.Location))
    End Sub
    Private ReadOnly Property UnitOfMeasureFactor() As Single
        Get
            Return MeasureSystem.CustomUnitToMicron(1, UnitOfMeasure)
        End Get
    End Property

    Private Function NormalizeRect(ByVal p1 As Point, ByVal p2 As Point) As Rectangle
        Dim r As New Rectangle
        If p1.X < p2.X Then
            r.X = p1.X
            r.Width = p2.X - p1.X
        Else
            r.X = p2.X
            r.Width = p1.X - p2.X
        End If
        If p1.Y < p2.Y Then
            r.Y = p1.Y
            r.Height = p2.Y - p1.Y
        Else
            r.Y = p2.Y
            r.Height = p1.Y - p2.Y
        End If
        Return r
    End Function

    Private Function LineLength(ByVal p1 As Point, ByVal p2 As Point, Optional ByVal ScaleFactor As Single = 1) As Single
        Dim r As Rectangle = NormalizeRect(p1, p2)
        _length = Math.Sqrt(r.Width ^ 2 + r.Height ^ 2) / ScaleFactor
        Return _length
    End Function

    Private Function Angle(ByVal p1 As Point, ByVal p2 As Point) As Single
        _angle = Math.Atan((p1.Y - p2.Y) / (p1.X - p2.X)) * (180 / Math.PI)
        Return _angle
    End Function
    Private Overloads Sub Dispose(ByVal disposing As Boolean)
        Try
            myPictureBoxControl = Nothing
        Catch ex As Exception

        End Try
    End Sub
End Class