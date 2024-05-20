Imports System.Collections.Generic
Imports System.Drawing.Drawing2D

Public Class ConversionInfo
    Public PhysicalWidth As Integer = 640
    Public PhysicalHeight As Integer = 480
    Private myScaleFactor As Single = 1.0F
    Public LogicalOrigin As System.Drawing.Point = RECT.InvalidPoint

#Region "Proprieta'"
    Public Property ScaleFactor() As Single
        Get
            If Single.IsNaN(myScaleFactor) OrElse Single.IsInfinity(myScaleFactor) Then
                myScaleFactor = 1
            End If
            Return myScaleFactor
        End Get
        Set(ByVal value As Single)
            myScaleFactor = Math.Abs(value)
        End Set
    End Property
    Public Property LogicalWidth() As Integer
        Get
            Debug.Assert(ScaleFactor <> 0)
            If ScaleFactor = 0 Then
                ScaleFactor = 1
                Return PhysicalWidth
            Else
                Return PhysicalWidth / ScaleFactor
            End If
        End Get
        Set(ByVal Value As Integer)
            If Value <> 0 Then
                ScaleFactor = PhysicalWidth / Value
            End If
        End Set
    End Property
    Public Property LogicalHeight() As Integer
        Get
            Debug.Assert(ScaleFactor <> 0)
            If ScaleFactor = 0 Then
                ScaleFactor = 1
                Return PhysicalHeight
            Else
                Return PhysicalHeight / ScaleFactor
            End If
        End Get
        Set(ByVal Value As Integer)
            If Value <> 0 Then
                ScaleFactor = PhysicalHeight / Value
            End If
        End Set
    End Property
    Public Property LogicalArea() As RECT
        Get
            Dim isNotValidArea As Boolean = (LogicalOrigin = RECT.InvalidPoint) OrElse (LogicalWidth = Integer.MaxValue) OrElse (LogicalHeight = Integer.MaxValue)
            isNotValidArea = isNotValidArea OrElse (LogicalWidth = 0) OrElse (LogicalHeight = 0)
            If isNotValidArea Then
                Return New RECT()
            End If

            LogicalArea.left = LogicalOrigin.X
            LogicalArea.top = LogicalOrigin.Y
            LogicalArea.bottom = LogicalOrigin.Y + LogicalHeight
            LogicalArea.right = LogicalOrigin.X + LogicalWidth
            LogicalArea.NormalizeRect()
        End Get
        Set(ByVal value As RECT)

            If (LogicalArea = value) Then
                Exit Property
            End If

            LogicalOrigin = New Point(value.left, value.top)
            LogicalWidth = value.Width
            LogicalHeight = value.Height
        End Set
    End Property

#End Region

#Region "Operatori"
    Public Shared Operator =(ByVal C1 As ConversionInfo, ByVal C2 As ConversionInfo) As Boolean
        Return (C1.PhysicalWidth = C2.PhysicalWidth) AndAlso (C1.PhysicalHeight = C2.PhysicalHeight) AndAlso (C1.ScaleFactor = C2.ScaleFactor) AndAlso (C1.LogicalOrigin = C2.LogicalOrigin)
    End Operator
    Public Shared Operator <>(ByVal C1 As ConversionInfo, ByVal C2 As ConversionInfo) As Boolean
        Return Not (C1 = C2)
    End Operator
#End Region

#Region "Conversione da coordinate fisiche a coordinate logiche"

    ''' <summary>
    ''' Trasforma una coordinata X da coordinate fisiche a coordinate logiche
    ''' </summary>
    Public Function ToLogicalCoordX(ByVal PhysicalCoordX As Single) As Single
        Try
            Return PhysicalCoordX / ScaleFactor + LogicalOrigin.X
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Trasforma una coordinata Y da coordinate fisiche a coordinate logiche
    ''' </summary>
    Public Function ToLogicalCoordY(ByVal PhysicalCoordY As Single) As Single
        Try
            Return PhysicalCoordY / ScaleFactor + LogicalOrigin.Y
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Trasforma una dimensione da coordinate fisiche a coordinate logiche
    ''' NOTA: Una dimensione e' intesa come "coordinata1- coordinata2", quindi
    ''' e' invariante rispetto alla posizione dell'origine
    ''' </summary>
    Public Function ToLogicalDimension(ByVal dimension As Single) As Single
        Try
            Return dimension / ScaleFactor
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Trasforma un punto da coordinate fisiche a coordinate logiche
    ''' </summary>
    Public Function ToLogicalPoint(ByVal PhysicalPoint As System.Drawing.Point) As System.Drawing.Point
        Try
            Return New System.Drawing.Point(PhysicalPoint.X / ScaleFactor + LogicalOrigin.X, PhysicalPoint.Y / ScaleFactor + LogicalOrigin.Y)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Trasforma un punto da coordinate fisiche a coordinate logiche
    ''' </summary>
    Public Function ToLogicalPoint(ByVal X As Integer, ByVal Y As Integer) As Point
        Try
            Return New Point(X / ScaleFactor + LogicalOrigin.X, Y / ScaleFactor + LogicalOrigin.Y)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

#End Region

#Region "Conversione da coordinate logiche a coordinate fisiche"

    ''' <summary>
    ''' Trasforma una coordinata X da coordinate logiche a coordinate fisiche
    ''' </summary>
    Public Function ToPhysicalCoordX(ByVal LogicalCoordX As Single) As Single
        Try
            Return (LogicalCoordX - LogicalOrigin.X) * ScaleFactor
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Trasforma una coordinata Y da coordinate logiche a coordinate fisiche
    ''' </summary>
    Public Function ToPhysicalCoordY(ByVal LogicalCoordY As Single) As Single
        Try
            Return (LogicalCoordY - LogicalOrigin.Y) * ScaleFactor
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Trasforma una dimensione da coordinate logiche a coordinate fisiche
    ''' NOTA: Una dimensione e' intesa come "coordinata1- coordinata2", quindi
    ''' e' invariante rispetto alla posizione dell'origine
    ''' </summary>
    Public Function ToPhysicalDimension(ByVal dimension As Single) As Single
        Try
            Return dimension * ScaleFactor
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Trasforma un punto da coordinate logiche a coordinate fisiche
    ''' </summary>
    Public Function ToPhysicalPoint(ByVal LogicalPoint As System.Drawing.Point) As System.Drawing.Point
        Try
            Return New System.Drawing.Point((LogicalPoint.X - LogicalOrigin.X) * ScaleFactor, (LogicalPoint.Y - LogicalOrigin.Y) * ScaleFactor)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Trasforma un rettangolo da coordinate logiche a coordinate fisiche
    ''' </summary>
    Public Function ToPhysicalRect(ByVal LogicalRect As RECT) As RECT
        Try
            Return New RECT(ToPhysicalCoordX(LogicalRect.left), ToPhysicalCoordY(LogicalRect.top), ToPhysicalCoordX(LogicalRect.right), ToPhysicalCoordY(LogicalRect.bottom))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

#End Region

#Region "Conversione da Dot a Micron"
    Public Shared Function DotToMicron(ByVal BitmapDPI As Integer) As Single
        Try
            Return 1 / ((BitmapDPI / 25.4) / 1000)
        Catch ex As Exception
            Return -1.0F
        End Try
    End Function
#End Region

#Region "Funzioni pubbliche"
    Public Function Clone() As Object
        Dim retVal As New ConversionInfo
        retVal.CopyParamsFrom(Me)
        Return retVal
    End Function
    Public Overridable Sub CopyParamsFrom(ByVal info As ConversionInfo)
        Me.PhysicalWidth = info.PhysicalWidth
        Me.PhysicalHeight = info.PhysicalHeight
        Me.ScaleFactor = info.ScaleFactor
        Me.LogicalOrigin = info.LogicalOrigin
    End Sub
#End Region

End Class

