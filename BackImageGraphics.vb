
Public Enum enBitmapOriginPosition
    TopLeft = 0
    Custom = 4
End Enum

Public Class cBackImageGraphics

#Region "Variabili private"
    Friend Origin As Point
    Private BitmapImage As Bitmap
    Private BitmapOrigin As enBitmapOriginPosition
    Private PixelWidth As Double = 100.0, PixelHeight As Double = 100.0
#End Region

#Region "Costruttori"
    Private Sub New()
        Return
    End Sub
    Public Sub New(ByVal BitmapImg As Bitmap, ByVal OriginX As Integer, ByVal OriginY As Integer, ByVal OriginPosition As enBitmapOriginPosition, ByVal Pixel_Width As Double, ByVal Pixel_Height As Double)
        BitmapImage = BitmapImg
        Origin.X = OriginX
        Origin.Y = OriginY
        BitmapOrigin = OriginPosition
        PixelWidth = Pixel_Width
        PixelHeight = Pixel_Height
        If PixelWidth < 10 Then
            PixelHeight = 10
        End If
        If PixelHeight < 10 Then
            PixelHeight = 10
        End If
    End Sub
#End Region

#Region "Funzioni pubbliche"
    Public Sub Draw(ByVal GR As Graphics)
        If BitmapImage Is Nothing Then
            Return
        End If
        GR.DrawImage(BitmapImage, New Rectangle(Origin.X, Origin.Y, BitmapImage.Width * PixelWidth, BitmapImage.Height * PixelWidth), 0, 0, BitmapImage.Width, BitmapImage.Height, GraphicsUnit.Pixel)
    End Sub
#End Region

#Region "Inizializzazione e finalizzazione"
    Public Sub Dispose()
        Try
            Me.BitmapImage.Dispose()
            Me.BitmapImage = Nothing
        Catch ex As Exception
            Return
        End Try
    End Sub
#End Region

End Class
