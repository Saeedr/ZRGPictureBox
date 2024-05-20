Imports System
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Collections.Generic

Partial Public Class ZRGPictureBoxControl
    Inherits System.Windows.Forms.UserControl

#Region "Pan & zoom"
    Friend Const ZoomMultiplier As Single = 1.25
    Friend Const PanFactorNoShift As Single = 100.0! / 3.0!
    Friend Const PanFactorWithShift As Single = 10.0!
#End Region

#Region "Size"
    Public Shared ReadOnly DefaultMinLogicalWindowSize As Size = New Size(2000, 2000)
    Public Shared ReadOnly DefaultMaxLogicalWindowSize As Size = New Size(100000000, 100000000)
#End Region

#Region "Costanti"
    ''' <summary>
    ''' Unita' di misura di deafult utilizzata per la visualizzazione delle coordinate e per i righelli
    ''' </summary>
    Public Const DefaultUnitOfMeasure As MeasureSystem.enUniMis = MeasureSystem.enUniMis.mm

#End Region

#Region "Eventi"
    Public Shadows Event MouseClick(ByVal sender As ZRGPictureBoxControl, ByVal e As System.Windows.Forms.MouseEventArgs, ByVal LogicalCoord As Point, ByVal CurrentClickAction As enClickAction)
    Public Shadows Event MouseMove(ByVal sender As ZRGPictureBoxControl, ByVal e As System.Windows.Forms.MouseEventArgs, ByVal LogicalCoord As Point, ByVal CurrentClickAction As enClickAction)
    Public Shadows Event MouseDown(ByVal sender As ZRGPictureBoxControl, ByVal e As System.Windows.Forms.MouseEventArgs, ByVal LogicalCoord As Point, ByVal CurrentClickAction As enClickAction)
    Public Shadows Event MouseUp(ByVal sender As ZRGPictureBoxControl, ByVal e As System.Windows.Forms.MouseEventArgs, ByVal LogicalCoord As Point, ByVal CurrentClickAction As enClickAction)
    Public Shadows Event MouseEnter(ByVal sender As ZRGPictureBoxControl, ByVal e As System.EventArgs)
    Public Shadows Event MouseLeave(ByVal sender As ZRGPictureBoxControl, ByVal e As System.EventArgs)
    Public Shadows Event Paint(ByVal sender As ZRGPictureBoxControl, ByVal e As System.Windows.Forms.PaintEventArgs)

    Public Event OnMeasureCompleted(ByVal sender As ZRGPictureBoxControl, ByVal StartPoint As Point, ByVal EndPoint As Point)
    Public Event OnRedrawCompleted(ByVal sender As ZRGPictureBoxControl, ByVal CacheRebuilded As Boolean)
    Public Event OnPictureBoxDoubleClick(ByVal sender As ZRGPictureBoxControl, ByVal e As System.Windows.Forms.MouseEventArgs, ByVal LogicalCoord As Point)
    Public Event OnMinimumZoomLevelReached(ByVal sender As ZRGPictureBoxControl)
    Public Event OnMaximumZoomLevelReached(ByVal sender As ZRGPictureBoxControl)

    Public Event OnMeasureUnitChanged(ByVal unit As MeasureSystem.enUniMis)
    Public Event OnClickActionChanged(ByVal oldClickAction As enClickAction, ByVal newClickAction As enClickAction)
#End Region

#Region "Variabili private"
    Private myGraphicInfo As New ConversionInfo()
    Private myClickAction As enClickAction = enClickAction.Zoom
    Private myCoordinatesBox As CoordinatesBox = New CoordinatesBox(Me)
    Private myDefaultBackgroundColor As Color = Color.WhiteSmoke
    Private myZoomSelectionBoxColor As Color = Color.Black

    Private myShowGrid As Boolean = True
    Private myGridView As GridKind = GridKind.Crosses
    Private myGridStep As Integer = 10000
    Private mySmartGridAdjust As Boolean = True

#Region "Gestione del double buffering video"

    '''' <summary>
    '''' Bitmap che costituisce il buffer video di primo livello (persistente fino al Refresh())
    '''' </summary>
    Private myRefreshBackBuffer As Bitmap

    ''' <summary>
    ''' Bitmap che costituisce il buffer video di secondo livello (persistente fino al Redraw())
    ''' </summary>
    Private myRedrawBackBuffer As Bitmap

#End Region

#Region "Varie"
    Private myIsDragging As Boolean = False
    Private myIsLoaded As Boolean = False
    Private myRulers As New Rulers(Me)
    Private myLastMouseDownPoint As Point
    Public ReadOnly Property LastMouseDownLogicalCoord() As Point
        Get
            Return myLastMouseDownPoint
        End Get
    End Property
    Private WithEvents myDistanceRuler As DistanceRuler = New DistanceRuler(Me)
    Private myIsLayoutSuspended As Boolean = True
#End Region

#Region "Variabili per il Resize()"

    Private myLastVisibleAreaRequested As RECT = DefaultRect
    Private myResizeBeginEndPreviewArea As RECT = DefaultRect
    Private myResizeMode As ResizeMode = ResizeMode.Stretch
    Private myIsBetweenResizeBeginEnd As Boolean = False
    Dim myBeginResizeClientArea As Rectangle

#End Region

#Region "Immagine di sfondo della PictureBox"
    Private myPictureBoxImagePosition As enBitmapOriginPosition = enBitmapOriginPosition.TopLeft
    Private myPictureBoxImageCustomOrigin As Point
    Private myShowPictureBoxImage As Boolean = True
    Private myPictureBoxImageGR As cBackImageGraphics
    Private myPictureBoxImage As System.Drawing.Image
    Private myPictureBoxImagePixelSize_micron As Integer = 100
    Public Property BackgroundImagePixelSize_Mic() As Integer
        Get
            Return myPictureBoxImagePixelSize_micron
        End Get
        Set(ByVal value As Integer)
            If myPictureBoxImagePixelSize_micron <> value Then
                myPictureBoxImagePixelSize_micron = value
                myPictureBoxImageGR = New cBackImageGraphics(myPictureBoxImage, ImageCustomOrigin.X, ImageCustomOrigin.Y, enBitmapOriginPosition.TopLeft, myPictureBoxImagePixelSize_micron, myPictureBoxImagePixelSize_micron)
            End If
        End Set
    End Property
    Public Property ImageCustomOrigin() As Point
        Get
            Return myPictureBoxImageCustomOrigin
        End Get
        Set(ByVal value As Point)
            myPictureBoxImageCustomOrigin = value
        End Set
    End Property
    Public Property ImagePosition() As enBitmapOriginPosition
        Get
            Return myPictureBoxImagePosition
        End Get
        Set(ByVal value As enBitmapOriginPosition)
            myPictureBoxImagePosition = value
            If myPictureBoxImage IsNot Nothing Then
                myPictureBoxImageGR = New cBackImageGraphics(myPictureBoxImage, ImageCustomOrigin.X, ImageCustomOrigin.Y, enBitmapOriginPosition.TopLeft, myPictureBoxImagePixelSize_micron, myPictureBoxImagePixelSize_micron)
            End If
        End Set
    End Property
    Public Shadows Property Image As System.Drawing.Image
        Get
            Return myPictureBoxImage
        End Get
        Set(value As System.Drawing.Image)
            myPictureBoxImage = value
            If value IsNot Nothing Then
                myPictureBoxImageGR = New cBackImageGraphics(myPictureBoxImage, ImageCustomOrigin.X, ImageCustomOrigin.Y, enBitmapOriginPosition.TopLeft, myPictureBoxImagePixelSize_micron, myPictureBoxImagePixelSize_micron)
            Else
                myPictureBoxImageGR = Nothing
            End If
        End Set
    End Property
#End Region

#Region "Dimensioni e unita' di misura"
    Private myBorderStyle As BorderStyle = Windows.Forms.BorderStyle.FixedSingle
    Private myUnitOfMeasure As MeasureSystem.enUniMis = DefaultUnitOfMeasure
    Private myMinLogicalWindowSize As Size = DefaultMinLogicalWindowSize
    Private myMaxLogicalWindowSize As Size = DefaultMaxLogicalWindowSize
#End Region

#Region "Flag per la visualizzazione"
    Private myShowMouseCoordinates As Boolean = True
    Private myShowRulers As Boolean = True
    Private myIsChangingAutoScroll As Boolean = False
#End Region

#Region "Box di selezione/zoom"
    Private mySelectionBox As New SelectionBoxElement(Me)
#End Region

#End Region

#Region "Proprieta'"

#Region "Origine, coordinate, dimensioni, ecc. "
    Protected Overrides ReadOnly Property DefaultSize() As Size
        Get
            Return New Size(560, 400)
        End Get
    End Property
    <Browsable(False)> _
    Public Property ScaleFactor() As Single
        Get
            Return GraphicInfo.ScaleFactor
        End Get
        Set(ByVal Value As Single)
            ' Se il nuovo valore mi porterebbe fuori dai limiti minimi o massimi, lo ignoro
            If (Width / Value) > MaxLogicalWindowSize.Width AndAlso Value < GraphicInfo.ScaleFactor Then Exit Property
            If (Height / Value) > MaxLogicalWindowSize.Height AndAlso Value < GraphicInfo.ScaleFactor Then Exit Property
            If (Width / Value) < MinLogicalWindowSize.Width AndAlso Value > GraphicInfo.ScaleFactor Then Exit Property
            If (Height / Value) < MinLogicalWindowSize.Height AndAlso Value > GraphicInfo.ScaleFactor Then Exit Property
            ' Aggiorno i dati interni
            GraphicInfo.ScaleFactor = Value
        End Set
    End Property
    <Browsable(False)> _
    Public Property LogicalOrigin() As System.Drawing.Point
        Get
            Return GraphicInfo.LogicalOrigin
        End Get
        Set(ByVal Value As System.Drawing.Point)
            GraphicInfo.LogicalOrigin = Value
        End Set
    End Property
    <Browsable(False)> _
    Public ReadOnly Property LogicalCenter() As Point
        Get
            Return New Point(LogicalOrigin.X + LogicalWidth / 2, LogicalOrigin.Y + LogicalHeight / 2)
        End Get
    End Property
    <Browsable(False)> _
    Public Property LogicalWidth() As Integer
        Get
            Return GraphicInfo.LogicalWidth
        End Get
        Set(ByVal Value As Integer)
            GraphicInfo.LogicalWidth = Value
        End Set
    End Property
    <Browsable(False)> _
    Public Property LogicalHeight() As Integer
        Get
            Return GraphicInfo.LogicalHeight
        End Get
        Set(ByVal Value As Integer)
            GraphicInfo.LogicalHeight = Value
        End Set
    End Property
    <Browsable(False)> _
    Public Property LogicalArea() As RECT
        Get
            Return GraphicInfo.LogicalArea
        End Get
        Private Set(ByVal value As RECT)
            GraphicInfo.LogicalArea = value
        End Set
    End Property
    Public Property MinLogicalWindowSize() As Size
        Get
            Return myMinLogicalWindowSize
        End Get
        Set(ByVal Value As Size)
            myMinLogicalWindowSize = Value
        End Set
    End Property
    Public Property MaxLogicalWindowSize() As Size
        Get
            Return myMaxLogicalWindowSize
        End Get
        Set(ByVal Value As Size)
            myMaxLogicalWindowSize = Value
            ' Aggiorno i dati delle scrollbar
            If ShowScrollbars Then
                UpdateScrollbars()
            End If
        End Set
    End Property
    Public Property GraphicInfo() As ConversionInfo
        Get
            Return myGraphicInfo
        End Get
        Private Set(ByVal Value As ConversionInfo)
            myGraphicInfo = Value
        End Set
    End Property

#Region "Funzioni che impediscono la serializzazione di queste proprieta'"
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeScaleFactor() As Boolean
        Return False
    End Function
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeLogicalOrigin() As Boolean
        Return False
    End Function
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeLogicalWidth() As Boolean
        Return False
    End Function
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeLogicalHeight() As Boolean
        Return False
    End Function
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeLogicalArea() As Boolean
        Return False
    End Function
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeMinLogicalWindowSize() As Boolean
        Return MinLogicalWindowSize <> DefaultMinLogicalWindowSize
    End Function
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeMaxLogicalWindowSize() As Boolean
        Return MaxLogicalWindowSize <> DefaultMaxLogicalWindowSize
    End Function

#End Region

#End Region

#Region "Flag per la visualizzazione"
    ''' <summary>
    ''' Visualizza l'immagine di sfondo della picturebox
    ''' </summary>
    <Description("Visualizza l'immagine di sfondo della picturebox."), Category("Opzioni di visualizzazione"), _
     DefaultValue(True)> _
    Public Property ShowPictureBoxBackgroundImage() As Boolean
        Get
            Return myShowPictureBoxImage
        End Get
        Set(ByVal value As Boolean)
            ' Check se il valore e' gia' quello desiderato
            If myShowPictureBoxImage = value Then
                Exit Property
            End If
            myShowPictureBoxImage = value
        End Set
    End Property

    ''' <summary>
    ''' Permette di visualizzare le coordinate a cui si trova il mouse
    ''' </summary>
    <Description("Permette di visualizzare le coordinate a cui si trova il mouse"), _
     DefaultValue(True)> _
    Public Property ShowMouseCoordinates() As Boolean
        Get
            Return myShowMouseCoordinates
        End Get
        Set(ByVal value As Boolean)
            myShowMouseCoordinates = value
        End Set
    End Property

    <Description("Imposta se visualizzare la griglia"), Category("Opzioni di visualizzazione"), _
     DefaultValue(True)> _
    Public Property ShowGrid() As Boolean
        Get
            Return myShowGrid
        End Get
        Set(ByVal Value As Boolean)
            myShowGrid = Value
        End Set
    End Property
    ''' <summary>
    ''' Permette di visualizzare i righelli
    ''' </summary>
    <Description("Permette di visualizzare i righelli"), Category("Opzioni di visualizzazione"), _
     DefaultValue(True)> _
    Public Property ShowRulers() As Boolean
        Get
            Return myShowRulers
        End Get
        Set(ByVal Value As Boolean)
            myShowRulers = Value
        End Set
    End Property
#End Region

#Region "Gestione del double buffering video"

    '''' <summary>
    '''' Bitmap che costituisce il buffer video di primo livello (persistente fino al Refresh())
    '''' </summary>
    Protected ReadOnly Property RefreshBackBuffer() As Bitmap
        Get
            Return myRefreshBackBuffer
        End Get
    End Property

    ''' <summary>
    ''' Bitmap che costituisce il buffer video di secondo livello (persistente fino al Redraw())
    ''' </summary>
    Protected ReadOnly Property RedrawBackBuffer() As Bitmap
        Get
            Return myRedrawBackBuffer
        End Get
    End Property

#End Region

#Region "Colori"

#Region "Elementi della PictureBox"
    Public Shared ReadOnly Property AxesColor() As Color
        Get
            Return Color.Navy
        End Get
    End Property
    Public Shared ReadOnly Property RulerColor() As Color
        Get
            Return Color.White
        End Get
    End Property
    Public Property BackgroundColor() As Color
        Get
            Return myDefaultBackgroundColor
        End Get
        Set(ByVal Value As Color)
            myDefaultBackgroundColor = Value
        End Set
    End Property
    ''' <summary>
    ''' Colore del cursore a croce
    ''' </summary>
    <Description("Imposta il colore del cursore a croce"), Category("Colors"), _
    DefaultValue(GetType(Color), "Black")> _
    Public Property CrossCursorColor() As Color
        Get
            Return FullCrossCursor.Color
        End Get
        Set(ByVal Value As Color)
            FullCrossCursor.Color = Value
        End Set
    End Property
    Public ReadOnly Property GridColor() As Color
        Get
            Return Color.LightSteelBlue
        End Get
    End Property
    Public ReadOnly Property SnapGridColor() As Color
        Get
            Return Color.Gray
        End Get
    End Property
    ''' <summary>
    ''' Colore del box di zoom/selezione
    ''' </summary>
    <Description("Imposta il colore del box di zoom/selezione"), Category("Colors"), _
    DefaultValue(GetType(Color), "Black")> _
    Public Property ZoomSelectionBoxColor() As Color
        Get
            Return myZoomSelectionBoxColor
        End Get
        Set(ByVal Value As Color)
            Try
                myZoomSelectionBoxColor = Value
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Set
    End Property
#End Region

#End Region

#Region "Griglia e Snap"
    <Description("Imposta la modalità di visualizzazione delle griglie"), _
     DefaultValue(GetType(GridKind), "Crosses")> _
    Public Property GridView() As GridKind
        Get
            Return myGridView
        End Get
        Set(ByVal Value As GridKind)
            myGridView = Value
        End Set
    End Property
    ''' <summary>
    ''' Imposta il passo della griglia
    ''' </summary>
    <Description("Imposta il passo della griglia"), _
     DefaultValue(10000)> _
    Public Property GridStep() As Integer
        Get
            Return myGridStep
        End Get
        Set(ByVal Value As Integer)
            myGridStep = Value
        End Set
    End Property
    Public Property SmartGridAdjust() As Boolean
        Get
            Return mySmartGridAdjust
        End Get
        Set(ByVal Value As Boolean)
            mySmartGridAdjust = Value
        End Set
    End Property
#End Region

#Region "Tasti premuti"

    ''' <summary>
    ''' Ritorna true se e' premuto il tasto SHIFT
    ''' </summary>
    Public Shared ReadOnly Property IsShiftKeyPressed() As Boolean
        Get
            Return (Control.ModifierKeys And Keys.Shift) <> 0
        End Get
    End Property
    Public Shared ReadOnly Property IsAltKeyPressed() As Boolean
        Get
            Return (Control.ModifierKeys And Keys.Alt) <> 0
        End Get
    End Property
    Public Shared ReadOnly Property IsCtrlKeyPressed() As Boolean
        Get
            Return (Control.ModifierKeys And Keys.Control) <> 0
        End Get
    End Property
#End Region

#Region "Misura delle distanze"
    Private Sub myDistanceRuler_CaptureFinished(ByVal sender As Object, ByVal e As CaptureEventArgs) Handles myDistanceRuler.CaptureFinished
        RaiseEvent OnMeasureCompleted(Me, GraphicInfo.ToLogicalPoint(e.StartPoint), GraphicInfo.ToLogicalPoint(e.EndPoint))
    End Sub
#End Region

#Region "Menu di contesto"
    Public Event ShowContextMenuRequired(ByVal sender As ZRGPictureBoxControl, ByVal X As Single, ByVal Y As Single)
    Public Sub RaiseContextMenuRequest(ByVal sender As ZRGPictureBoxControl, ByVal X As Single, ByVal Y As Single)
        Try
            RaiseEvent ShowContextMenuRequired(sender, X, Y)
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "Cursori"
    ''' <summary>
    ''' Cursore correntemente visualizzato
    ''' </summary>
    Protected Property CurrentCursor() As Cursor
        Get
            Return MyBase.Cursor
        End Get
        Set(ByVal value As Cursor)
            ' Se mi e' stato richiesto il cursore di waiting, non permetto cambiamenti
            If UseWaitCursor Then
                Exit Property
            End If
            ' Altrimenti aggiorno il cursore visualizzato
            ' NOTA: Devo modificare MyBase.Cursor, non Me.Cursor, altrimenti entro in un ciclo infinito
            '       (ed inoltre perdo la reference al cursore di default)
            If MyBase.Cursor IsNot value Then
                MyBase.Cursor = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Cursore di default
    ''' </summary>
    Protected Overrides ReadOnly Property DefaultCursor() As Cursor
        Get
            If Me.ClickAction = enClickAction.Zoom Then
                Return cCommonCursors.ZoomCursor
            Else
                Return cCommonCursors.EditCursor
            End If
        End Get
    End Property
#End Region

#Region "Scrollbar"

    ''' <summary>
    ''' Ottiene o imposta un valore che indica se il contenitore consentirà all'utente di scorrere i controlli posizionati all'esterno dei limiti visibili.
    ''' NOTA: E' stata ridichiarata per impedire la modifica da parte di applicativi esterni
    ''' </summary>
    <Browsable(False)> _
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Public Overrides Property AutoScroll() As Boolean
        Get
            Return MyBase.AutoScroll
        End Get
        Set(ByVal value As Boolean)
            MyBase.AutoScroll = value
        End Set
    End Property

    ''' <summary>
    ''' Ottiene o imposta la dimensione minima dello scorrimento automatico.
    ''' NOTA: E' stata ridichiarata per impedire la modifica da parte di applicativi esterni
    ''' </summary>
    <Browsable(False)> _
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Public Shadows Property AutoScrollMinSize() As System.Drawing.Size
        Get
            Return MyBase.AutoScrollMinSize
        End Get
        Private Set(ByVal value As System.Drawing.Size)
            MyBase.AutoScrollMinSize = value
        End Set
    End Property

    ''' <summary>
    ''' Ottiene o imposta la dimensione del margine di scorrimento automatico.
    ''' NOTA: E' stata ridichiarata per impedire la modifica da parte di applicativi esterni
    ''' </summary>
    <Browsable(False)> _
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Public Shadows Property AutoScrollMargin() As System.Drawing.Size
        Get
            Return MyBase.AutoScrollMargin
        End Get
        Private Set(ByVal value As System.Drawing.Size)
            MyBase.AutoScrollMargin = value
        End Set
    End Property

    ''' <summary>
    ''' Permette di visualizzare le scrollbar
    ''' </summary>
    <Description("Permette di visualizzare le scrollbar"), _
    DefaultValue(False)> _
    Public Property ShowScrollbars() As Boolean
        Get
            Return AutoScroll
        End Get
        Set(ByVal Value As Boolean)
            ' Check se il valore e' gia' quello desiderato
            If AutoScroll = Value Then
                Exit Property
            End If
            ' Aggiorno i dati delle scrollbar, lo faccio prima di visualizzarle, cosi' mi evito un ridisegno con le scrollbar sbagliate
            If Value Then
                UpdateScrollbars()
            End If
            ' NOTA: Devo impostare il flag che mi fara' saltare l'evento Resize() che si genera quando si cambia il valore di Autoscroll
            myIsChangingAutoScroll = True
            AutoScroll = Value
            myIsChangingAutoScroll = False
        End Set
    End Property

#End Region

#Region "Stato attuale"
    ''' <summary>
    ''' Imposta il tipo di azione da eseguire sul click del mouse
    ''' </summary>
    <Description("Imposta il tipo di azione da eseguire sul click del mouse"), _
    DefaultValue(GetType(enClickAction), "SelectObjects")> _
    Public Property ClickAction() As enClickAction
        Get
            Return myClickAction
        End Get
        Set(ByVal Value As enClickAction)
            ' Check se ho gia' il valore desiderato
            'If myClickAction = Value Then
            '    Exit Property
            'End If
            ' Aggiorno i dati interni
            Dim oldClickAction As enClickAction = myClickAction
            myClickAction = Value

            ' Se sono in modalita' di progettazione, non devo fare altro
            If DesignMode Then
                Exit Property
            End If

            ' Genero l'evento "modificato il tipo di azione da eseguire sul click del mouse"
            RaiseEvent OnClickActionChanged(oldClickAction, myClickAction)

            ' NOTA: Se imposto la modalita' a Zoom, nel Designer sparisce il cursore.
            '       Penso che questo accada perche' il Designer non sa gestire la classe PictureBoxCursors. -> INVESTIGARE
            Select Case myClickAction
                Case enClickAction.Zoom
                    ' L'eventuale box di selezione deve mantenere l'aspetto della finestra
                    SelectionBox.KeepAspectRatio = True
                    Cursor = cCommonCursors.ZoomCursor
                Case enClickAction.MeasureDistance
                    ' L'eventuale box di selezione deve mantenere l'aspetto della finestra
                    SelectionBox.KeepAspectRatio = True
                    Cursor = cCommonCursors.EditCursor
            End Select
        End Set
    End Property
#End Region

#Region "Box di selezione/zoom"
    Public ReadOnly Property SelectionBox() As SelectionBoxElement
        Get
            Return mySelectionBox
        End Get
    End Property

#End Region


#Region "Variabili per il Resize()"
    <DefaultValue(GetType(ResizeMode), "Stretch")> _
    Public Property ResizeMode() As ResizeMode
        Get
            Return myResizeMode
        End Get
        Set(ByVal value As ResizeMode)
            myResizeMode = value
        End Set
    End Property
#End Region

    <Browsable(False)> _
    Public ReadOnly Property IsLayoutSuspended() As Boolean
        Get
            ' Evito le Redraw() a DesignTime
            Return myIsLayoutSuspended 'OrElse Me.DesignMode
        End Get
    End Property

    Protected ReadOnly Property IsDragging() As Boolean
        Get
            Return myIsDragging
        End Get
    End Property

    Protected Property IsLoaded() As Boolean
        Get
            Return myIsLoaded
        End Get
        Private Set(ByVal value As Boolean)
            myIsLoaded = value
        End Set
    End Property

    Public ReadOnly Property ContainsMousePosition() As Boolean
        Get
            '' NOTA: Non posso usare il ClientRectangle percvhe' se le scrollbar sono attive,
            ''       l'area occupata dalle scrollbar viene esclusa dall'area client 
            'Dim p As Point = PointToClient(MousePosition)
            'If (p.X < 0) OrElse (p.X > Me.Size.Width) Then Return False
            'If (p.Y < 0) OrElse (p.Y > Me.Size.Height) Then Return False
            'Return True
            Return Me.ClientRectangle.Contains(PointToClient(MousePosition))
        End Get
    End Property
    Private myFullCrossCursor As CrossCursor
    Private ReadOnly Property FullCrossCursor As CrossCursor
        Get
            If myFullCrossCursor Is Nothing Then
                myFullCrossCursor = New CrossCursor(Me)
            End If
            Return myFullCrossCursor
        End Get
    End Property
    Private ReadOnly Property FullPictureBoxCross() As Boolean
        Get
            Return myClickAction = enClickAction.MeasureDistance
        End Get
    End Property
    ''' <summary>
    ''' Imposta il bordo della PictureBox
    ''' </summary>
    <Description("Imposta il bordo della PictureBox"), _
    DefaultValue(GetType(BorderStyle), "FixedSingle")> _
    Public Overloads Property BorderStyle() As BorderStyle
        Get
            Return myBorderStyle
        End Get
        Set(ByVal Value As BorderStyle)
            myBorderStyle = Value
        End Set
    End Property

    ''' <summary>
    ''' Unità di misura del sistema di coordinate logico della PictureBox
    ''' </summary>
    <Description("Imposta l'unità di misura del sistema di coordinate logico della PictureBox"), _
    DefaultValue(GetType(MeasureSystem.enUniMis), "Millimeter")> _
    Public Property UnitOfMeasure() As MeasureSystem.enUniMis
        Get
            Return myUnitOfMeasure
        End Get
        Set(ByVal Value As MeasureSystem.enUniMis)
            If myUnitOfMeasure = Value Then
                Exit Property
            End If
            myUnitOfMeasure = Value
            RaiseEvent OnMeasureUnitChanged(Value)
            ' Se sto disegnando qualcosa che richiede l'unita' di misura, devo fare un ridisegno
            If ShowMouseCoordinates OrElse ShowRulers Then
                Invalidate()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Ritorna l'area correntemente visualizzata dalla PictureBox [coordinate logiche]
    ''' NOTA BENE: L'area varia al variare dello zoom
    ''' </summary>
    <Browsable(False)> _
    Public ReadOnly Property VisibleRect() As RECT
        Get
            Return New RECT(LogicalOrigin.X, LogicalOrigin.Y, LogicalOrigin.X + LogicalWidth, LogicalOrigin.Y + LogicalHeight)
        End Get
    End Property


#Region "Funzioni che impediscono la serializzazione di queste proprieta'"
    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeAutoScroll() As Boolean
        Return False
    End Function

    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeAutoScrollMargin() As Boolean
        Return False
    End Function

    <EditorBrowsable(EditorBrowsableState.Never)> _
    Private Function ShouldSerializeAutoScrollMinSize() As Boolean
        Return False
    End Function
#End Region

#End Region

#Region "Costruttori"

    Public Sub New(ByVal visible As Boolean)
        'This call is required by the Windows Form Designer.
        MyBase.New()

        Try
            Me.Visible = visible
            InitializeComponent()
            'Add any initialization after the InitializeComponent() call
            Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub New()
        Me.New(True)
    End Sub

#End Region

#Region "Funzioni di supporto per gli eventi"
    ''' <summary>
    ''' Funzione chiamata durante la MouseMove(), permette di personalizzare il cursore nelle classi derivate.
    ''' Viene passata la posizione attuale del cursore in coordinate logiche.
    ''' Deve ritornare true se il cursore e' stato personalizzato, false altrimenti.
    ''' </summary>
    Protected Overridable Function OnCustomCursorRequest(ByVal logicalCoord As Point) As Boolean
        Return False
    End Function

    ''' <summary>
    ''' Cancella gli eventuali dati temporanei usati tra un MouseDown e un MouseUp
    ''' </summary>
    Private Sub ResetTemporaryData()
        ' Cancello il rettangolo di selezione, in modo che non venga piu' ridisegnato
        SelectionBox.Reset()
    End Sub

    ''' <summary>
    ''' Aggiorna lo stato del flag che indica se si sta facendo un drag and drop
    ''' </summary>
    Protected Sub UpdateDraggingState()
        ' Si puo' avere un drag and drop solo se il tasto sinistro e' abbassato
        ' e se esiste un punto di partenza da cui si e' cominciato il drag
        ' e se sono in standardView
        myIsDragging = (MouseButtons = Windows.Forms.MouseButtons.Left) AndAlso (Not SelectionBox.TopLeftCorner = RECT.InvalidPoint)
        If myIsDragging Then
            ' Si puo' avere un drag and drop solo se mi sono spostato dal punto in cui avevo premuto il tasto
            Dim physicalMousePos As Point = PointToClient(MousePosition)
            Dim distanceX As Integer = physicalMousePos.X - myLastMouseDownPoint.X
            Dim distanceY As Integer = physicalMousePos.Y - myLastMouseDownPoint.Y
            myIsDragging = (Math.Abs(distanceX) >= 3 OrElse Math.Abs(distanceY) >= 3)
        End If
    End Sub

    ''' <summary>
    ''' Aggiorna l'area selezionata con il box di selezione.
    ''' Genera l'evento "Richiesta una selezione di oggetti"
    ''' Utilizzata nella gestione dell'evento MouseUp
    ''' </summary>
    Private Sub UpdateSelectionBox()
        ' Se il primo punto dell'area e' non valido, ignoro il click
        ' Puo' succedere se ho premuto il bottone del mouse fuori dal mio applicativo, poi ho trascinato il mouse
        ' sopra la PictureBox e poi ho rilasciato il bottone, oppure se ho premuto OK su una dialog che era
        ' sopra la PictureBox e che poi e' scomparsa
        If SelectionBox.TopLeftCorner = RECT.InvalidPoint Then
            Exit Sub
        End If

        ' Rettangolo corrispondente al box di selezione.
        ' Vale sia nel caso di singolo click che di selezione tramite trascinamento di un'area
        Dim selectedBox As RECT = SelectionBox

        ' Check se il rettangolo di selezione e' valido
        If selectedBox.IsZeroSized Then
            Exit Sub
        End If

    End Sub

    ''' <summary>
    ''' Aggiorna le dimensioni degli elementi che hanno bisogno di essere ridimensionati
    ''' Crea delle nuove bitmap di primo e secondo livello in funzione delle nuove dimensioni.
    ''' Ritorna true se le bitmap sono state aggiornate, false altrimenti.
    ''' </summary>
    Private Function UpdateDimensions() As Boolean
        ' Se sono in fase di creazione del controllo o se ho sospeso il layout del controllo, posso uscire direttamente
        Try
            'If (Not Me.Created) OrElse IsLayoutSuspended Then
            '    Return False
            'End If
            ' Check se la finestra e' minimizzata
            If (Width < 1) OrElse (Height < 1) Then
                Return False
            End If

            ' Per evitare di allocare e deallocare continuamente delle bitmap quando l'utente ridimensiona la finestra
            ' trascinando con il mouse, le bitmap vengono allocate arrotondando ai 100 pixel superiori
            Dim newWidth As Integer = Math.Ceiling(CDbl(Me.Width) / 100.0) * 100
            Dim newHeight As Integer = Math.Ceiling(CDbl(Me.Height) / 100.0) * 100

            ' Flag che indica se le bitmap vanno ricreate oppure no
            Dim bitmapCreationNeeded As Boolean = (myRedrawBackBuffer Is Nothing) OrElse (myRedrawBackBuffer.Width < newWidth) OrElse (myRedrawBackBuffer.Height < newHeight)
            If bitmapCreationNeeded Then
                ' Pulisco eventuali bitmap precedenti
                If (myRedrawBackBuffer IsNot Nothing) Then myRedrawBackBuffer.Dispose()
                If (myRefreshBackBuffer IsNot Nothing) Then myRefreshBackBuffer.Dispose()
                ' Aggiorno le dimensioni delle bitmap su cui andro' a tracciare
                myRedrawBackBuffer = New Bitmap(newWidth, newHeight)
                myRefreshBackBuffer = New Bitmap(newWidth, newHeight)
            End If

            ' Ritorno true se le bitmap sono state aggiornate, false altrimenti.
            Return bitmapCreationNeeded
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

#Region "Gestione degli eventi"

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            ' Aggiorno lo stato del flag che indica se si sta facendo un drag and drop
            myIsDragging = False

            ' Posizione del mouse in vari tipi di coordinate
            Dim logicalMousePos As Point = GraphicInfo.ToLogicalPoint(e.X, e.Y)
            myLastMouseDownPoint = logicalMousePos

            If e.Button = Windows.Forms.MouseButtons.Right Then
                RaiseEvent MouseDown(Me, e, logicalMousePos, myClickAction)
                MyBase.OnMouseDown(e)
                RaiseContextMenuRequest(Me, e.X, e.Y)
                Return
            End If

            ' Imposto il primo punto del rettangolo di selezione/zoom
            If (e.Button = Windows.Forms.MouseButtons.Left) Then
                SelectionBox.TopLeftCorner = logicalMousePos
                SelectionBox.BottomRightCorner = logicalMousePos
            End If

            ' Il tipo di azione da eseguire dipende dalla modalita' di click attuale
            Select Case ClickAction
                Case enClickAction.None
                    Exit Select
                Case enClickAction.MeasureDistance
                    ' Imposto il primo punto del righello usato per misurare
                    myDistanceRuler.MouseDown(Me, e)
                Case enClickAction.Zoom
                    ' L'eventuale box di selezione mantiene l'aspetto della finestra
                    SelectionBox.KeepAspectRatio = True
            End Select

            RaiseEvent MouseDown(Me, e, logicalMousePos, myClickAction)
            MyBase.OnMouseDown(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Protected Overrides Sub OnMouseClick(ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            If IsDragging Then
                Return
            End If
            ' Posizione del mouse in vari tipi di coordinate
            Dim physicalMousePos As Point = New Point(e.X, e.Y)
            Dim logicalMousePos As Point = GraphicInfo.ToLogicalPoint(physicalMousePos)
            ' Genero l'evento 
            RaiseEvent MouseClick(Me, e, logicalMousePos, myClickAction)
            MyBase.OnMouseClick(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ListDragSource_GiveFeedback(ByVal sender As Object, ByVal e As GiveFeedbackEventArgs) Handles Me.GiveFeedback

        ' Set the custom cursor based upon the effect.
        If ((e.Effect And DragDropEffects.Move) = DragDropEffects.Move) Then
            Cursor.Current = Cursors.SizeAll
        Else
            Cursor.Current = Cursors.No
        End If

    End Sub


    Private DisableHighlightWhileSelecting As Boolean = False
    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim needsRepaint As Boolean = False
        Try
            ' Posizione del mouse in vari tipi di coordinate
            Dim physicalMousePos As Point = New Point(e.X, e.Y)
            Dim logicalMousePos As Point = GraphicInfo.ToLogicalPoint(physicalMousePos)

            ' Aggiorno lo stato del flag che indica se si sta facendo un drag and drop
            Me.UpdateDraggingState()

            ' Il tipo di azione da eseguire dipende dalla modalita' di click attuale
            Select Case myClickAction
                Case enClickAction.None
                    Exit Select
                Case enClickAction.MeasureDistance
                    ' Se mi sto muovendo con il tasto sinistro premuto, aggiorno il righello usato per misurare
                    If e.Button = System.Windows.Forms.MouseButtons.Left Then
                        myDistanceRuler.MouseMove(Me, e)
                        needsRepaint = True
                    End If
                Case enClickAction.Zoom
                    ' Se mi sto muovendo con il tasto sinistro premuto,
                    ' aggiorno il secondo punto del rettangolo di selezione
                    If IsDragging Then
                        SelectionBox.BottomRightCorner = logicalMousePos
                        needsRepaint = True
                    End If
            End Select

            ' Genero l'evento 
            RaiseEvent MouseMove(Me, e, logicalMousePos, myClickAction)
            MyBase.OnMouseMove(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            ' Se serve, ridisegno la finestra
            If needsRepaint OrElse myShowMouseCoordinates Then
                Invalidate()
            End If
        End Try
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            ' Posizione del mouse in vari tipi di coordinate
            Dim logicalMousePos As Point = GraphicInfo.ToLogicalPoint(e.X, e.Y)

            ' Aggiorno il secondo punto del rettangolo di selezione
            If (e.Button = Windows.Forms.MouseButtons.Left) Then
                If (SelectionBox.BottomRightCorner <> logicalMousePos) Then
                    SelectionBox.BottomRightCorner = logicalMousePos
                End If
            End If

            ' Il tipo di azione da eseguire dipende dalla modalita' di click attuale
            Select Case myClickAction
                Case enClickAction.None
                    Exit Select
                Case enClickAction.MeasureDistance
                    myDistanceRuler.MouseUp(Me, e)
                Case enClickAction.Zoom
                    If Not IsDragging Then
                        If e.Button = System.Windows.Forms.MouseButtons.Left Then ZoomForward(logicalMousePos)
                        If e.Button = System.Windows.Forms.MouseButtons.Right Then ZoomBack(logicalMousePos)
                    Else
                        If e.Button = System.Windows.Forms.MouseButtons.Left Then
                            ShowLogicalWindow(SelectionBox)
                        End If
                    End If
            End Select

            RaiseEvent MouseUp(Me, e, logicalMousePos, myClickAction)
            MyBase.OnMouseUp(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            ' Cancello gli eventuali dati temporanei usati tra un MouseDown e un MouseUp
            ResetTemporaryData()
        End Try
    End Sub
    Protected Overrides Sub OnMouseDoubleClick(ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            MyBase.OnMouseDoubleClick(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub PictureBoxEx_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Inizializzo i dati interni
        Initialize()
        ' Notifico che questo controllo e' stato caricato nel suo contenitore
        IsLoaded = True
        ' NOTA: Non serve fare un ridisegno, tanto nella Initialize() faccio uno zoom al rettangolo di default
        UpdateDimensions()
        If DesignMode Then
            Redraw()
        End If
    End Sub

#Region "Eventi della finestra di selezione degli oggetti"
    ''' <summary>
    ''' Ritorna true se l'oggetto passato e' una stringa.
    ''' </summary>
    Private Function IsString(ByVal obj As Object) As Boolean
        Debug.Assert(obj IsNot Nothing, "Trovato oggetto nullo")
        If (TypeOf (obj) Is String) Then
            Return True
        End If
    End Function
    Public Overloads Sub Invalidate()
        MyBase.Invalidate()
        ' se la seguente istruzione viene decommentata
        ' il tooltip che evidenzia l'angolo di cui si ruota un blocco funziona
        ' ma il refresh della picturebox si comporta in modo strano su Windows Vista
        'Update()
    End Sub
#End Region

#End Region

#Region "Gestione delle scrollbar"

    ''' <summary>
    ''' Imposta il valore delle scrollbar (dipende dal livello di zoom della finestra logica)
    ''' </summary>
    Private Sub UpdateScrollbars()
        ' NOTA: La thumb di una scrollbar e' la parte che posso cliccare e trascinare qua e la

        ' Aggiorno la dimensione che fa apparire le scrollbar, in pratica la modifico in modo
        ' che la thumb delle scrollbar diventi piu' stretta o piu' larga in funzione dell'area logica visualizzata
        Dim newValueX, newValueY As Integer
        newValueX = Math.BigMul(Me.Size.Width, MaxLogicalWindowSize.Width) / LogicalWidth
        newValueY = Math.BigMul(Me.Size.Height, MaxLogicalWindowSize.Height) / LogicalHeight
        ' Check se ho dei valori validi
        newValueX = Math.Max(newValueX, 0)
        newValueX = Math.Max(newValueY, 0)
        Me.AutoScrollMinSize = New Size(newValueX, newValueY)

        ' Aggiorno la posizione delle thumb all'interno delle scrollbar
        ' NOTA: Devono essere sempre maggiori o uguali a zero
        newValueX = Math.BigMul(Me.Size.Width, LogicalCenter.X + MaxLogicalWindowSize.Width / 2) / LogicalWidth
        newValueY = Math.BigMul(Me.Size.Height, LogicalCenter.Y + MaxLogicalWindowSize.Height / 2) / LogicalHeight
        ' Check se ho dei valori validi
        newValueX = Math.Max(newValueX, 0)
        newValueX = Math.Max(newValueY, 0)
        Me.AutoScrollPosition = New Point(newValueX, newValueY)

        ' Aggiorno gli spostamenti relativi alle scrollbar
        Me.HorizontalScroll.SmallChange = CSng(LogicalWidth) * PanFactorWithShift / 100.0!
        Me.HorizontalScroll.LargeChange = CSng(LogicalWidth) * PanFactorNoShift / 100.0!
        Me.VerticalScroll.SmallChange = CSng(LogicalHeight) * PanFactorWithShift / 100.0!
        Me.VerticalScroll.LargeChange = CSng(LogicalHeight) * PanFactorNoShift / 100.0!
    End Sub

    Protected Overrides Sub OnScroll(ByVal se As ScrollEventArgs)
        'MsgBox(String.Format("Oldvalue: {0}", se.OldValue))
        'MsgBox(String.Format("Newvalue: {0}", se.NewValue))
        'MsgBox(String.Format("ScrollValues BEFORE: {0}/{1}  {2}/{3}", Me.HorizontalScroll.Value, MaxLogicalWindowSize.Width, Me.VerticalScroll.Value, MaxLogicalWindowSize.Height))

        ' In funzione del tipo di scroll, ho uno spostamento diverso 
        If se.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
            'MsgBox("OnScroll: Horizontal")
            Select Case se.Type
                Case ScrollEventType.SmallIncrement
                    PanRight(PanFactorWithShift)
                Case ScrollEventType.SmallDecrement
                    PanLeft(PanFactorWithShift)
                Case ScrollEventType.LargeIncrement
                    PanRight(PanFactorNoShift)
                Case ScrollEventType.LargeDecrement
                    PanLeft(PanFactorNoShift)
                Case Else
                    ' Arrivo qui se:
                    '  - l'utente sta trascinando la scrollbar (ScrollEventType.ThumbTrack)
                    '  - l'utente ha finito il trascinamento della scrollbar (ScrollEventType.ThumbPosition)
                    '  - la scrollbar e' stata portata al suo valore minimo o massimo (ScrollEventType.First, ScrollEventType.Last)
                    '  - (???) (ScrollEventType.EndScroll)
                    Dim newValueX As Integer = Math.BigMul(Me.HorizontalScroll.Value, LogicalWidth) / Me.Size.Width
                    newValueX -= MaxLogicalWindowSize.Width / 2
                    LogicalOrigin = New Point(newValueX, LogicalOrigin.Y)
                    Redraw()
                    'MsgBox(se.Type.ToString())
            End Select
        Else
            'MsgBox("OnScroll: Vertical")
            Select Case se.Type
                Case ScrollEventType.SmallIncrement
                    PanDown(PanFactorWithShift)
                Case ScrollEventType.SmallDecrement
                    PanUp(PanFactorWithShift)
                Case ScrollEventType.LargeIncrement
                    PanDown(PanFactorNoShift)
                Case ScrollEventType.LargeDecrement
                    PanUp(PanFactorNoShift)
                Case Else
                    ' Arrivo qui se:
                    '  - l'utente sta trascinando la scrollbar (ScrollEventType.ThumbTrack)
                    '  - l'utente ha finito il trascinamento della scrollbar (ScrollEventType.ThumbPosition)
                    '  - la scrollbar e' stata portata al suo valore minimo o massimo (ScrollEventType.First, ScrollEventType.Last)
                    '  - (???) (ScrollEventType.EndScroll)
                    Dim newValueY As Integer = Math.BigMul(Me.VerticalScroll.Value, LogicalHeight) / Me.Size.Height
                    newValueY -= MaxLogicalWindowSize.Height / 2
                    LogicalOrigin = New Point(LogicalOrigin.X, newValueY)
                    Redraw()
                    'MsgBox(se.Type.ToString())
            End Select
        End If
    End Sub

#End Region

#Region "Routine per la Redraw()"

    Public Sub Redraw()
        Redraw(False)
    End Sub

    Public Overridable Sub Redraw(ByVal forceGraphicCacheRebuild As Boolean)
        Dim GR As Graphics = Nothing
        Try
            ' Se sono in fase di creazione del controllo o se il controllo e' gia' stato distrutto, posso uscire direttamente
            If (Me.IsDisposed) Then
                Exit Sub
            End If

            ' Se la PictureBox non e' visibile o se il suo layout e' sospeso, salto il ridisegno
            'If Not IsLoaded OrElse Not Visible Then
            If Not Visible OrElse IsLayoutSuspended Then
                'Exit Sub
            End If


            'If Not IsLayoutSuspended Then
            UpdateDimensions()
            'End If

            ' Check se il fattore di scala e' valido. 
            ' Serve se chiamo la Redraw() quando la finestra e' minimizzata
            If (ScaleFactor = 0.0) OrElse (LogicalWidth = 0) OrElse (LogicalHeight = 0) Then
                Exit Sub
            End If

            ' Check se ho un'area valida da visualizzare 
            Debug.Assert(LogicalArea.IsNonZeroSized)


            ' Questa routine va a disegnare sul buffer di secondo livello
            GR = GetScaledGraphicObject(myRedrawBackBuffer)
            If GR Is Nothing Then Exit Sub

            ' Pulisco con il colore di sfondo
            GR.Clear(BackgroundColor)

            If myShowPictureBoxImage Then
                DrawPictureBoxImage(GR)
            End If

            ' Disegno le griglie
            DrawGrids(GR)


            ' Chiamo la Refresh()
            Refresh()

            ' Aggiorno i dati delle scrollbar
            ' NOTA: Devo farlo qui e non nella ShowLogicalWindow() perche' le routine che fanno lo zoom e il pan
            '       chiamano direttamente la Redraw() senza passare per la ShowLogicalWindow() 
            If ShowScrollbars Then
                UpdateScrollbars()
            End If

            RaiseEvent OnRedrawCompleted(Me, forceGraphicCacheRebuild)
        Catch ex As InvalidOperationException
            If ex.Message.ToUpper.Contains("CROSS-THREAD") Then
                Return
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            ' Libero la memoria allocata per l'oggetto Graphics
            If GR IsNot Nothing Then
                GR.Dispose()
            End If
        End Try
    End Sub
    Public Function GetScreenShot() As Image
        Try
            Dim OutSize As New System.Drawing.Size(Me.Width, Me.Height)
            Dim retValue As New System.Drawing.Bitmap(OutSize.Width, OutSize.Height, Imaging.PixelFormat.Format32bppPArgb)
            Dim gr As Graphics = Graphics.FromImage(retValue)
            Me.Redraw(True)
            gr.DrawImageUnscaled(myRedrawBackBuffer, 0, 0)
            Return retValue
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function SaveAScreenShot(ByVal strDestFileName As String, ByVal _Format As System.Drawing.Imaging.ImageFormat) As Boolean
        Try
            If myRedrawBackBuffer IsNot Nothing Then
                myRedrawBackBuffer.Save(strDestFileName, _Format)
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Disegna l'immagine di sfondo della picturebox
    ''' </summary>
    Private Sub DrawPictureBoxImage(ByVal GR As Graphics)
        Try
            If myPictureBoxImageGR Is Nothing Then Exit Sub
            Select Case myPictureBoxImagePosition
                Case enBitmapOriginPosition.TopLeft
                    myPictureBoxImageGR.Origin = Point.Empty
                Case enBitmapOriginPosition.Custom
                    myPictureBoxImageGR.Origin = myPictureBoxImageCustomOrigin
            End Select
            myPictureBoxImageGR.Draw(GR)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Disegna le griglie sul Graphics passatogli
    ''' </summary>
    Private Sub DrawGrids(ByVal GR As Graphics)
        ' Disegno la griglia
        If myShowGrid Then
            DrawGrid(GR, 0, GridStep, GridView, GridColor, SmartGridAdjust)
        End If
        If myShowGrid Then
            DrawAxes(GR)
        End If
    End Sub

    ''' <summary>
    ''' Disegna una griglia
    ''' </summary>
    Private Sub DrawGrid(ByVal GR As Graphics, ByVal GridInitialOffset As Integer, ByVal GridStep As Integer, ByVal GridMode As GridKind, ByVal GridColor As Color, ByVal SmartAdjust As Boolean)
        Try
            ' Check se il passo della griglia e' valido
            If GridStep = 0 Then Exit Sub

            ' Check se il fattore di scala e' valido
            If (ScaleFactor <= 0.0) Then
                Exit Sub
            End If

            'Controllo la risoluzione della griglia per evitare sovraffollamenti ...
            If (GridStep < myRulers.GetRulerStep) AndAlso SmartAdjust Then
                GridStep = myRulers.GetRulerStep
            End If

            Dim InitialX As Integer = Math.Ceiling(LogicalOrigin.X / GridStep) * GridStep
            Dim InitialY As Integer = Math.Ceiling(LogicalOrigin.Y / GridStep) * GridStep
            Dim myPen As New Pen(GridColor)
            Dim FinalX As Integer = Math.Floor((LogicalOrigin.X + LogicalWidth) / GridStep) * GridStep
            Dim FinalY As Integer = Math.Floor((LogicalOrigin.Y + LogicalHeight) / GridStep) * GridStep
            Dim iIterX As Integer
            Dim iIterY As Integer
            Select Case GridMode
                Case GridKind.Crosses
                    For iIterY = InitialY To FinalY Step GridStep
                        For iIterX = InitialX To FinalX Step GridStep
                            GR.DrawLine(myPen, iIterX + GridInitialOffset - 10 / ScaleFactor, iIterY + GridInitialOffset, iIterX + GridInitialOffset + 10 / ScaleFactor, iIterY + GridInitialOffset)
                            GR.DrawLine(myPen, iIterX + GridInitialOffset, iIterY + GridInitialOffset - 10 / ScaleFactor, iIterX + GridInitialOffset, iIterY + GridInitialOffset + 10 / ScaleFactor)
                        Next
                    Next
                Case GridKind.FullLines
                    For iIterY = InitialY To FinalY Step GridStep
                        GR.DrawLine(myPen, LogicalOrigin.X, iIterY + GridInitialOffset, LogicalWidth + LogicalOrigin.X, iIterY + GridInitialOffset)
                    Next
                    For iIterX = InitialX To FinalX Step GridStep
                        GR.DrawLine(myPen, iIterX + GridInitialOffset, LogicalOrigin.Y, iIterX + GridInitialOffset, LogicalHeight + LogicalOrigin.Y)
                    Next
                Case GridKind.Points
                    For iIterY = InitialY To FinalY Step GridStep
                        For iIterX = InitialX To FinalX Step GridStep
                            GR.DrawLine(myPen, iIterX + GridInitialOffset - 1 / ScaleFactor, iIterY + GridInitialOffset, iIterX + GridInitialOffset + 1 / ScaleFactor, iIterY + GridInitialOffset)
                            GR.DrawLine(myPen, iIterX + GridInitialOffset, iIterY + GridInitialOffset - 1 / ScaleFactor, iIterX + GridInitialOffset, iIterY + GridInitialOffset + 1 / ScaleFactor)
                        Next
                    Next
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' Disegna gli assi cartesiani
    ''' </summary>
    Private Sub DrawAxes(ByVal GR As Graphics)
        Try
            ' Check se il fattore di scala e' valido
            If (ScaleFactor <= 0.0) Then
                Exit Sub
            End If

            GR.DrawLine(New Pen(AxesColor, -1), 0, LogicalOrigin.Y, 0, LogicalOrigin.Y + LogicalHeight)
            GR.DrawLine(New Pen(AxesColor, -1), LogicalOrigin.X, 0, LogicalOrigin.X + LogicalWidth, 0)
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
#End Region

#Region "Routine per la Refresh()"

    ''' <summary>
    ''' Copia myBackbufferBitmap sulla image della PictureBox (non tiene conto di cambiamenti di scala)
    ''' Fa anche l'aggiornamento della posizione delle maniglie della finestra di selezione
    ''' </summary>
    Public Shadows Sub Refresh(Optional ByVal _Invalidate As Boolean = True)
        Dim GR As Graphics = Nothing
        Try
            ' Se sono in fase di creazione del controllo o se il controllo e' gia' stato distrutto, posso uscire direttamente
            If (Not Me.Created) OrElse (Me.IsDisposed) Then
                Exit Sub
            End If

            ' Check se il fattore di scala e' valido. 
            ' Serve se chiamo la Refresh() quando la finestra e' minimizzata
            If ScaleFactor = 0.0 Then
                Return
            End If

            ' Oggetto graphics su cui andro' a tracciare
            GR = Graphics.FromImage(myRefreshBackBuffer)

            ' Copio il buffer video di terzo livello su quello di secondo livello
            ' NOTA: Inizialmente usavo la GR.DrawImageUnscaled(), ma e' inutile ricopiare la parte di bitmap che cade fuori
            '       dall'area client del controllo. Quindi uso la DrawImageUnscaledAndClipped()

            'LUCADN: ho rimosso l'impostazione a None di SmoothingMode
            'GR.SmoothingMode = Drawing2D.SmoothingMode.None

            GR.DrawImageUnscaledAndClipped(myRedrawBackBuffer, Me.ClientRectangle)
            'MsgBox(String.Format("Redraw(): ClientRectangle = {0}", ClientRectangle))

            ' Disegno i righelli
            If myShowRulers Then
                myRulers.Draw(GR)
            End If

            ' Adesso posso scalare il Graphics per disegnare i vari DesignObject
            ScaleGraphicObject(GR)

            If _Invalidate Then Invalidate()

#If DEBUG_TIMING_REFRESH Then
            globalTimer.Trace("Refresh: tempo totale = ")
#End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            ' Libero la memoria allocata per l'oggetto Graphics
            If GR IsNot Nothing Then
                GR.Dispose()
            End If
        End Try
    End Sub

#End Region

#Region "Routine per la OnPaint()"

    ''' <summary>
    ''' Disegna una preview di cio' che si vedra' in modalita' Stretch o Normal dopo aver terminato di ridimensionare la PictureBox.
    ''' </summary>
    Private Sub DrawResizePreview(ByVal GR As Graphics)
        ' Pulisco lo sfondo
        GR.Clear(BackgroundColor)

        ' Disegno le griglie
        ScaleGraphicObject(GR)
        DrawGrids(GR)
        GR.ResetTransform()

        ' Check se la bitmap di partenza e' corretta, serve nel caso in cui Stilista parta minimizzato o con ClientArea nulla
        If (myRedrawBackBuffer Is Nothing) OrElse (myRedrawBackBuffer.Width = 0) OrElse (myRedrawBackBuffer.Width = 0) Then
            Exit Sub
        End If

        ' Se sono in modalita' Normal non ho bisogno di disegnare la bitmap scalata, mi basta:
        '  - fare una preview utilizzando l'ultima bitmap di refresh
        '  - aggiornare i righelli (sarebbe un "in piu'" ma lo faccio lo stesso, cosi' e' piu' completo)
        If (ResizeMode = ResizeMode.Normal) Then
            ' Faccio la preview utilizzando l'ultima bitmap di redraw
            GR.DrawImage(myRedrawBackBuffer, Point.Empty)
            ' Aggiorno i righelli
            If myShowRulers Then
                myRulers.Draw(GR)
            End If
            Exit Sub
        End If

        ' NOTA: Se arrivo qui, vuol dire che sono in modalita' Stretch e che devo disegnare la bitmap scalata

        ' Fattore di scala necessario per convertire la dimensione delle vecchia bitmap, creata dalla Redraw()
        ' con un ClientArea diversa, in una bitmap adatta alla ClientArea attuale
        Dim actualPreviewArea As RECT = GraphicInfo.ToPhysicalRect(myLastVisibleAreaRequested)
        Dim tmpScaleFactor As Single = CSng(Math.Max(actualPreviewArea.Width, 1)) / CSng(Math.Max(myResizeBeginEndPreviewArea.Width, 1))

        ' Rettangolo di output su cui disegnare la bitmap [coordinate logiche]
        ' Questo e' il rettangolo su cui verra' fatto lo shrink dell'ultima bitmap creata dalla Redraw()
        Dim bitmapOutputRect As New RECT

        ' Il mio unico punto fermo tra la vecchia bitmap (che avevo creato con una ClientArea completamente diversa) 
        ' e la nuova ClientArea e' il punto di origine dell'ultima area visibile richiesta.
        ' Quindi devo partire dalla "coordinata fisica attuale dell'ultima area visibile" e sottrarre la
        ' "coordinata che aveva con la ClientArea iniziale", moltiplicata per il fattore di scala
        bitmapOutputRect.top = actualPreviewArea.top - (myResizeBeginEndPreviewArea.top * tmpScaleFactor)
        bitmapOutputRect.left = actualPreviewArea.left - (myResizeBeginEndPreviewArea.left * tmpScaleFactor)

        ' Le dimensioni del rettangolo non richiedono spostamenti, basta adattarle con il fattore di scala
        bitmapOutputRect.Width = myRedrawBackBuffer.Width * tmpScaleFactor
        bitmapOutputRect.Height = myRedrawBackBuffer.Height * tmpScaleFactor

        ' Disegno la bitmap 
        GR.DrawImage(myRedrawBackBuffer, bitmapOutputRect)

        ' Aggiorno i righelli
        If myShowRulers Then
            myRulers.Draw(GR)
        End If
    End Sub

#End Region

#Region "Override della classe base"
    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
        ' Don't allow the background to paint
        'MsgBox("OnPaintBackground")
    End Sub
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Try
            If IsLayoutSuspended OrElse (Not Visible) OrElse (Not Created) OrElse (IsDisposed) Then
                Exit Sub
            End If

            Dim gr As Graphics = e.Graphics

            ' Se sto ridimensionando, devo disegnare la preview e basta  
            If myIsBetweenResizeBeginEnd Then
                DrawResizePreview(gr)
                Exit Sub
            End If

            ' Copio il backbuffer sullo schermo
            gr.DrawImage(myRefreshBackBuffer, e.ClipRectangle, e.ClipRectangle, GraphicsUnit.Pixel)
            'MsgBox(String.Format("OnPaint(): ClipRectangle={0}", e.ClipRectangle))

            ' Se sono in modalita' designer, non serve fare altro
            If DesignMode Then
                Exit Sub
            End If

            ' Posizione del mouse in vari tipi di coordinate
            Dim physicalMousePos As Point = Me.PointToClient(MousePosition)
            Dim logicalMousePos As Point = GraphicInfo.ToLogicalPoint(physicalMousePos)

            ' Adesso posso scalare il Graphics in modo definitivo
            ScaleGraphicObject(gr)

            ' Flag che indica se si puo' disegnare il box di zoom/selezione oggetti
            Dim drawSelectionBox As Boolean = False

            ' Il tipo di azione da eseguire dipende dalla modalita' di click attuale
            Select Case myClickAction
                Case enClickAction.None
                    Exit Select
                Case enClickAction.MeasureDistance
                    ' NOTA: Le successive chiamate a funzioni di disegno richiedono un Graphics NON scalato
                    gr.ResetTransform()
                    ' Disegno il righello usato per misurare
                    If IsCtrlKeyPressed Then
                        Dim ScaleFactor As Double = MeasureSystem.MicronToCustomUnit(CDbl(BackgroundImagePixelSize_Mic), myUnitOfMeasure, False)
                        myDistanceRuler.Painting(gr, ScaleFactor)
                    Else
                        myDistanceRuler.Painting(gr)
                    End If

                Case enClickAction.Zoom
                    ' Disegno l'eventuale box di zoom/selezione oggetti
                    drawSelectionBox = IsDragging
                    Exit Select
            End Select

            ' NOTA: Le successive chiamate a funzioni di disegno richiedono un Graphics NON scalato
            gr.ResetTransform()

            ' Disegno l'eventuale box di zoom/selezione oggetti
            If drawSelectionBox AndAlso (Not SelectionBox.IsInvalid) Then
                SelectionBox.Draw(gr)
            End If

            If FullPictureBoxCross Then

            End If

            If Me.FullPictureBoxCross AndAlso ContainsMousePosition Then
                FullCrossCursor.DrawCross(gr, logicalMousePos)
            End If

            If myShowMouseCoordinates Then
                If IsCtrlKeyPressed Then
                    Dim BitmapCoord As System.Drawing.Point
                    BitmapCoord.X = logicalMousePos.X / BackgroundImagePixelSize_Mic
                    BitmapCoord.Y = logicalMousePos.Y / BackgroundImagePixelSize_Mic
                    myCoordinatesBox.DrawCoordinateInfo(gr, BitmapCoord, True)
                Else
                    myCoordinatesBox.DrawCoordinateInfo(gr, logicalMousePos)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Evento chiamato quando comincio un ciclo di resize multipli.
    ''' NOTA: Viene generato sia quando l'utente ridimensiona la finestra (sia con il mouse che da tastiera) che quando la sposta.
    ''' </summary>
    Private Sub OnResizeBegin(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Notifico che sto gestendo una serie di eventi di resize
        myIsBetweenResizeBeginEnd = True
        ' Salvo la dimensione che la PictureBox ha all'inizio del ciclo di resize
        myBeginResizeClientArea = Me.ClientRectangle
        ' Converto l'area di preview in coordinate fisiche della bitmap
        myResizeBeginEndPreviewArea = GraphicInfo.ToPhysicalRect(myLastVisibleAreaRequested)
    End Sub

    ''' <summary>
    ''' Evento chiamato quando e' finto un ciclo di resize multipli.
    ''' NOTA: Viene generato sia quando l'utente ridimensiona la finestra (sia con il mouse che da tastiera) che quando la sposta.
    ''' </summary>
    Private Sub OnResizeEnd(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Notifico che ho finito di gestire una serie di eventi di resize
        myIsBetweenResizeBeginEnd = False
        ' Check se le dimensioni della PictureBox sono effettivamente cambiate
        ' NOTA: Serve perche' questa routine viene chiamata durante la gestione dell''evento OnResizeEnd(),
        '       che viene chiamato anche quando sposto la form contenente la PictureBox.
        If (myBeginResizeClientArea <> Me.ClientRectangle) Then
            ' Genero un evento di cambio dimensione, in modo da forzare il ridisegno
            OnSizeChanged(e)
        End If
    End Sub

    Protected Overrides Sub OnSizeChanged(ByVal e As EventArgs)
        Try
            ' MsgBox("OnSizeChanged()")
            ' NOTA: Quando cambio il valore di Me.AutoScroll, viene generato un evento Resize (e/o SizeChanged)
            '       anche se la dimensione della PictureBox non viene effettivamente modificata
            If myIsChangingAutoScroll Then Exit Sub

            ' Aggiorno i valori interni
            ' NOTA BENE: Questo mi va a modificare l'area logica che sto visualizzando!
            '            E' per questo che mi serve la variabile myLastVisibleAreaRequested, perche' altrimenti
            '            ad ogni Resize() perderei completamente l'area che stavo visualizzando
            GraphicInfo.PhysicalWidth = Me.Width
            GraphicInfo.PhysicalHeight = Me.Height

            ' Se la finestra e' minimizzata esce subito
            If (Width < 1) OrElse (Height < 1) Then Exit Sub
            'MsgBox("Resize to (" + picPictureBox.Width.ToString() + "," + picPictureBox.Height.ToString() + ")")

            If Not IsLoaded And ResizeMode <> ResizeMode.Stretch Then Exit Sub

            ' Check se l'utente sta ridimensionando la finestra (e quindi mi stanno arrivando tutta una serie di Resize() di fila)
            If myIsBetweenResizeBeginEnd Then
                ' Se sono in modalita' Stretch, imposto l'area logica in modo che sia uguale a quella 
                ' che andrei a visualizzare se il resize terminasse con queste dimensioni.
                ' Se non lo facessi, la bitmap di preview verrebbe disegnata in una posizione completamente sbagliata
                ' NOTA: Questo e' uno dei pochi casi in cui la LogicalArea va impostata direttamente.
                ' NOTA: Non va fatto in modalita' Normal, altrimenti righelli e icone vengono disegnati nel posto sbagliato
                If (ResizeMode = ResizeMode.Stretch) Then
                    LogicalArea = VisibleAreaToLogicalArea(myLastVisibleAreaRequested)
                End If

                ' Devo invalidare "a mano" la finestra, altrimenti non mi lascia disegnare sopra la parte di PictureBox che era gia' visibile 
                Invalidate()
                Exit Sub
            End If

            ' Se necessario, aggiorno le bitmap del buffer video
            ' NOTA: La UpdateDimensions() ritorna true se le bitmap sono state aggiornate, false altrimenti.
            UpdateDimensions()

            ' Check se la area logica deve variare in modo da mantenere l'ultima visualizzazione 
            If (ResizeMode = ResizeMode.Stretch) Then
                ' Faccio uno zoom all'ultima area visibile richiesta, senza salvarlo nella history.
                ' NOTA: L'area visibile richiesta verrebbe sovrascritta durante l'esecuzione della ZoomToLogicalWindow(),
                '       ma siccome io passo lo stesso valore che aveva prima, in pratica e' come se fosse invariante.
                ShowLogicalWindow(myLastVisibleAreaRequested, False)
            Else
                ' Devo sempre fare una Redraw(), perche' se faccio una Refresh() posso avere errori nel ridisegno nel caso in cui:
                '  - diminuisco la dimensione della finestra 
                '  - cambio la View attuale 
                '  - reingrandisco la finestra, con dimensioni che non obbligano a ricreare le bitmap
                Redraw()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            ' Devo chiamare la routine della classe base, altrimenti gli eventi OnSizeChanged() e OnResize() non vengono mai generati
            MyBase.OnSizeChanged(e)
        End Try
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal e As EventArgs)
        'MsgBox("OnMouseEnter()")
        ' Aggiorno il cursore al valore di default
        CurrentCursor = DefaultCursor
        ' Chiamo la corrispondente funzione della classe base
        MyBase.OnMouseEnter(e)
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
        Try
            'MsgBox("OnMouseLeave()")
            ' Se avevo il cursore a croce a pieno schermo, devo cancellarlo, altrimenti mi ritrovo 
            ' una riga che attraversa tutta la pictureBox e rimane li' fino a qunado non rientro con il mouse
            If Me.FullPictureBoxCross Then
                ' NOTA: Non serve chiamare la myCrossCursor.ResetCrossPosition(), tanto la posizione da disegnare
                '       viene passata al cursore direttamente nella OnPaint()
                Invalidate()
            End If
            ' Chiamo la corrispondente funzione della classe base
            MyBase.OnMouseLeave(e)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            'MsgBox("OnMouseWheel()")
            ' NOTA: Non utilizzare la variabile MousePosition, perche' ritorna la posizione 
            '       del mouse in coordinate DELLO SCHERMO (rispetto all'angolo superiore sinistro)
            Dim MouseLogicalPosition As Point
            MouseLogicalPosition.X = e.X / ScaleFactor + LogicalOrigin.X
            MouseLogicalPosition.Y = e.Y / ScaleFactor + LogicalOrigin.Y

            ' NOTA: Con la rotellina del mouse mi sposto come con il PAN con shift premuto,
            '       cioe' il pan con spostamento MINORE
            If e.Delta > 0 Then
                If IsCtrlKeyPressed Then
                    PanLeft(PanFactorWithShift)
                    Return
                End If
                If IsShiftKeyPressed Then
                    PanUp(PanFactorWithShift)
                    Return
                End If
                ' Fa uno Zoom In rispetto alla posizione del mouse
                ZoomForward(MouseLogicalPosition)
            Else
                If IsCtrlKeyPressed Then
                    PanRight(PanFactorWithShift)
                    Return
                End If
                If IsShiftKeyPressed Then
                    PanDown(PanFactorWithShift)
                    Return
                End If

                ' Fa uno Zoom Out rispetto alla posizione del mouse
                ZoomBack(MouseLogicalPosition)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Override dei parametri di creazione del controllo
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            ' Imposto il bordo del controllo
            Const WS_BORDER As Integer = &H800000
            Const WS_EX_STATICEDGE As Integer = &H20000
            Dim cp As CreateParams = MyBase.CreateParams
            Select Case myBorderStyle
                Case BorderStyle.FixedSingle
                    cp.Style = cp.Style Or WS_BORDER
                Case BorderStyle.Fixed3D
                    cp.ExStyle = cp.ExStyle Or WS_EX_STATICEDGE
            End Select

            Return cp
        End Get
    End Property

    Private Sub AddResizeHandlers()
        Dim parentForm As Form = FindForm()
        ' Check se il controllo e' nullo
        If (parentForm Is Nothing) Then
            Exit Sub
        End If
        If (parentForm.IsMdiChild) Then
            parentForm = parentForm.MdiParent
            If (parentForm Is Nothing) Then
                Exit Sub
            End If
        End If
        AddHandler parentForm.ResizeBegin, AddressOf Me.OnResizeBegin
        AddHandler parentForm.ResizeEnd, AddressOf Me.OnResizeEnd
    End Sub

    Private Sub RemoveResizeHandlers()
        Dim parentForm As Form = FindForm()
        ' Check se il controllo e' nullo
        If (parentForm Is Nothing) Then
            Exit Sub
        End If
        RemoveHandler parentForm.ResizeBegin, AddressOf Me.OnResizeBegin
        RemoveHandler parentForm.ResizeEnd, AddressOf Me.OnResizeEnd
    End Sub

#End Region

#Region "Gestione dello Zoom"

#Region "Varie"

    ''' <summary>
    ''' Ritorna true se sono al minimo livello di Zoom Out ammissibile
    ''' </summary>
    <Browsable(False)> _
    Public ReadOnly Property MinimumZoomReached() As Boolean
        Get
            ' Ritorno true se IL PROSSIMO zoom mi porterebbe fuori dai limiti
            Dim nextHeight As Integer = LogicalHeight * ZoomMultiplier
            Dim nextWidth As Integer = LogicalWidth * ZoomMultiplier
            Return nextHeight > MaxLogicalWindowSize.Height OrElse nextWidth > MaxLogicalWindowSize.Width
        End Get
    End Property

    ''' <summary>
    ''' Ritorna true se sono al massimo livello di Zoom In ammissibile
    ''' </summary>
    <Browsable(False)> _
    Public ReadOnly Property MaximumZoomReached() As Boolean
        Get
            ' Ritorno true se IL PROSSIMO zoom mi porterebbe fuori dai limiti
            Dim nextHeight As Integer = LogicalHeight / ZoomMultiplier
            Dim nextWidth As Integer = LogicalWidth / ZoomMultiplier
            Return nextHeight < myMinLogicalWindowSize.Height OrElse nextWidth < myMinLogicalWindowSize.Width
        End Get
    End Property

#End Region

#Region "Zoom in"

    ''' <summary>
    ''' Effettua uno "Zoom In" mettendo al centro dello schermo il punto passatogli
    ''' </summary>
    Public Sub ZoomForwardUsingCenter(ByVal ZoomCenter As Point)
        Try
            ' Se ho raggiunto lo zoom massimo, esco direttamente
            If MaximumZoomReached Then
                RaiseEvent OnMaximumZoomLevelReached(Me)
                Exit Sub
            End If

            ' Faccio lo zoom dell'area logica
            ' NOTA: Quando faccio uno zoom in, l'area logica visibile viene ridotta
            Dim tmpArea As RECT = LogicalArea.ExpandFromFixedPoint(1.0! / ZoomMultiplier, LogicalCenter)
            ' Riporto il centro della nuova area logica sul punto voluto
            tmpArea.Offset(ZoomCenter - LogicalCenter)

            ' Devo espandere e spostare della stessa quantita' anche l'ultima area visualizzata 
            myLastVisibleAreaRequested = myLastVisibleAreaRequested.ExpandFromFixedPoint(1.0! / ZoomMultiplier, LogicalCenter)
            myLastVisibleAreaRequested.Offset(ZoomCenter - LogicalCenter)

            ' Aggiorno l'area logica visualizzata
            LogicalArea = tmpArea

            ' Ridisegno la finestra
            Redraw()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Effettua uno "Zoom In" rispetto al centro logico della finestra
    ''' </summary>
    Public Sub ZoomForwardOnLogicalCenter()
        ZoomForwardUsingCenter(LogicalCenter)
    End Sub

    ''' <summary>
    ''' Effettua uno "Zoom In" mantenendo "fisso" il punto che gli viene passato
    ''' </summary>
    Public Sub ZoomForward(ByRef LogicalPosition As Point)
        ' Se ho raggiunto lo zoom massimo, esco direttamente
        If MaximumZoomReached Then
            RaiseEvent OnMaximumZoomLevelReached(Me)
            Exit Sub
        End If

        ' Trovo la distanza del punto passatomi dal centro dello schermo
        Dim distance As Point = LogicalPosition - LogicalCenter

        ' Trovo la distanza dal centro dello schermo che avrei con la prossima visualizzazione e calcolo 
        ' il centro dello schermo della prossima visualizzazione (rispetto al punto passatomi)
        distance.X /= ZoomMultiplier
        distance.Y /= ZoomMultiplier
        Dim newZoomCenter As System.Drawing.Point = LogicalPosition - distance

        ' Sposto lo zoom sul punto calcolato
        ZoomForwardUsingCenter(newZoomCenter)
    End Sub

#End Region

#Region "Zoom out"

    ''' <summary>
    ''' Effettua uno "Zoom Out" mettendo al centro dello schermo il punto passatogli
    ''' </summary>
    Public Sub ZoomBackUsingCenter(ByVal ZoomCenter As Point)
        Try
            ' Se ho raggiunto lo zoom minimo, esco direttamente
            If MinimumZoomReached Then
                RaiseEvent OnMinimumZoomLevelReached(Me)
                Exit Sub
            End If

            ' Faccio lo zoom dell'area logica
            ' NOTA: Quando faccio uno zoom out, l'area logica visibile viene espansa
            Dim tmpArea As RECT = LogicalArea.ExpandFromFixedPoint(ZoomMultiplier, LogicalCenter)
            ' Riporto il centro della nuova area logica sul punto voluto
            tmpArea.Offset(ZoomCenter - LogicalCenter)

            ' Devo espandere e spostare della stessa quantita' anche l'ultima area visualizzata 
            myLastVisibleAreaRequested = myLastVisibleAreaRequested.ExpandFromFixedPoint(ZoomMultiplier, LogicalCenter)
            myLastVisibleAreaRequested.Offset(ZoomCenter - LogicalCenter)


            ' Aggiorno l'area logica visualizzata
            LogicalArea = tmpArea

            ' Ridisegno la finestra
            Redraw()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Effettua uno "Zoom Out" rispetto al centro logico della finestra
    ''' </summary>
    Public Sub ZoomBackOnLogicalCenter()
        ZoomBackUsingCenter(LogicalCenter)
    End Sub

    ''' <summary>
    ''' Effettua uno "Zoom Out" mantenendo "fisso" il punto che gli viene passato
    ''' </summary>
    Public Sub ZoomBack(ByRef LogicalPosition As Point)

        ' Se ho raggiunto lo zoom minimo, esco direttamente
        If MinimumZoomReached Then
            RaiseEvent OnMinimumZoomLevelReached(Me)
            Exit Sub
        End If

        ' Trovo la distanza del punto passatomi dal centro dello schermo
        Dim distance As Point = LogicalPosition - LogicalCenter

        ' Trovo la distanza dal centro dello schermo che avrei con la prossima visualizzazione e calcolo 
        ' il centro dello schermo della prossima visualizzazione (rispetto al punto passatomi)
        distance.X *= ZoomMultiplier
        distance.Y *= ZoomMultiplier
        Dim newZoomCenter As System.Drawing.Point = LogicalPosition - distance

        ' Sposto lo zoom sul punto calcolato
        ZoomBackUsingCenter(newZoomCenter)
    End Sub

#End Region

#Region "Zoom"

    Public Sub ZoomToDefaultRect()
        Try
            If Image IsNot Nothing Then
                ZoomToFit()
            Else
                ' Zoom al rettangolo di default (senza centrarlo nella finestra)
                ShowLogicalWindow(DefaultRect, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub ZoomToFit()
        Try
            If Image IsNot Nothing Then
                Dim ImgR As New RECT(ImageCustomOrigin.X, ImageCustomOrigin.Y, ImageCustomOrigin.X + Image.Width, ImageCustomOrigin.Y + Image.Height)
                Dim PhysicalR As RECT
                PhysicalR.left = ImgR.left * BackgroundImagePixelSize_Mic
                PhysicalR.top = ImgR.top * BackgroundImagePixelSize_Mic
                PhysicalR.Width = ImgR.Width * BackgroundImagePixelSize_Mic
                PhysicalR.Height = ImgR.Height * BackgroundImagePixelSize_Mic
                ShowLogicalWindow(PhysicalR, True)

                'Dim LogicRect As RECT
                'LogicRect.left = GraphicInfo.ToLogicalCoordX(PhysicalR.left)
                'LogicRect.top = GraphicInfo.ToLogicalCoordX(PhysicalR.top)
                'LogicRect.Width = GraphicInfo.ToLogicalCoordX(PhysicalR.Width)
                'LogicRect.Height = GraphicInfo.ToLogicalCoordX(PhysicalR.Height)
                'ShowLogicalWindow(LogicRect, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

    <Browsable(False)> _
    Public Shared ReadOnly Property DefaultRect() As RECT
        Get
            ' Un rettangolo di default per me vale circa 10 x 10 cm, tengo lo zero in alto a sx
            ' e faccio in modo che l'incrocio degli assi sia visibile (tengo circa 3mm in piu') 
            Dim _Rect As RECT
            _Rect.left = -3000 ' -3mm 
            _Rect.top = -3000 ' -3mm 
            _Rect.right = 100000 ' 10cm 
            _Rect.bottom = 100000 ' 10cm 
            Return _Rect
        End Get
    End Property

#End Region

#Region "Gestione del Pan"
    Public Overridable Sub PanHome()
    End Sub
    Public Overridable Sub PanEnd()
    End Sub
    Public Overridable Sub PageUp()
    End Sub
    Public Overridable Sub PageDown()
    End Sub

    ''' <summary>
    ''' Effettua uno spostamento in basso di un valore pari alla "percentuale della finestra logica" passatagli
    ''' </summary>
    Public Overridable Sub PanDown(ByVal Percent As Single)
        Try
            ' Check se ho raggiunto il limite visualizzabile
            If LogicalOrigin.Y + LogicalHeight > MaxLogicalWindowSize.Height / 2 Then Exit Sub
            ' Calcolo l'offset di cui spostarmi 
            Dim offsetY As Integer = CInt(CSng(LogicalHeight) * (Percent / 100.0!))
            ' Sposto l'origine
            LogicalOrigin = New Point(LogicalOrigin.X, LogicalOrigin.Y + offsetY)
            ' Devo spostare della stessa quantita' anche l'ultima area visualizzata 
            myLastVisibleAreaRequested.Offset(0, offsetY)
            ' Faccio un ridisegno
            Redraw()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Effettua uno spostamento in alto di un valore pari alla "percentuale della finestra logica" passatagli
    ''' </summary>
    Public Overridable Sub PanUp(ByVal Percent As Single)
        ' Check se ho raggiunto il limite visualizzabile
        If LogicalOrigin.Y < -MaxLogicalWindowSize.Height / 2 Then Exit Sub
        ' Calcolo l'offset di cui spostarmi 
        Dim offsetY As Integer = CInt(CSng(LogicalHeight) * (Percent / 100.0!))
        ' Sposto l'origine
        LogicalOrigin = New Point(LogicalOrigin.X, LogicalOrigin.Y - offsetY)
        ' Devo spostare della stessa quantita' anche l'ultima area visualizzata 
        myLastVisibleAreaRequested.Offset(0, -offsetY)
        ' Faccio un ridisegno
        Redraw()
    End Sub

    ''' <summary>
    ''' Effettua uno spostamento a sinistra di un valore pari alla "percentuale della finestra logica" passatagli
    ''' </summary>
    Public Overridable Sub PanLeft(ByVal Percent As Single)
        ' Check se ho raggiunto il limite visualizzabile
        If LogicalOrigin.X < -MaxLogicalWindowSize.Width / 2 Then Exit Sub
        ' Calcolo l'offset di cui spostarmi
        Dim offsetX As Integer = CInt(CSng(LogicalWidth) * (Percent / 100.0!))
        ' Sposto l'origine
        LogicalOrigin = New Point(LogicalOrigin.X - offsetX, LogicalOrigin.Y)
        ' Devo spostare della stessa quantita' anche l'ultima area visualizzata 
        myLastVisibleAreaRequested.Offset(-offsetX, 0)
        ' Faccio un ridisegno
        Redraw()
    End Sub

    ''' <summary>
    ''' Effettua uno spostamento a destra di un valore pari alla "percentuale della finestra logica" passatagli
    ''' </summary>
    Public Overridable Sub PanRight(ByVal Percent As Single)
        ' Check se ho raggiunto il limite visualizzabile
        If LogicalOrigin.X + LogicalWidth > MaxLogicalWindowSize.Width / 2 Then Exit Sub
        ' Calcolo l'offset di cui spostarmi
        Dim offsetX As Integer = CInt(CSng(LogicalWidth) * (Percent / 100.0!))
        ' Sposto l'origine
        LogicalOrigin = New Point(LogicalOrigin.X + offsetX, LogicalOrigin.Y)
        ' Devo spostare della stessa quantita' anche l'ultima area visualizzata 
        myLastVisibleAreaRequested.Offset(offsetX, 0)
        ' Faccio un ridisegno
        Redraw()
    End Sub

#End Region

#Region "Gestione input da tastiera"

    ''' <summary>
    ''' Flag che indica se la ProcessCmdKey sta processando un carattere che non arriva da tastiera
    ''' </summary>
    Private myAlreadyInProcessCmdKey As Boolean = False

    ''' <summary>
    ''' Funzione per il processamento dei messaggi da tastiera che arrivano da altri controlli
    ''' </summary>
    Public Function ProcessKeyboardKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        Return ProcessCmdKey(msg, keyData)
    End Function

    ''' <summary>
    ''' Override della funzione che recupera i messaggi da windows
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        ' Check se sto tentando una ricorsione infinita
        If myAlreadyInProcessCmdKey Then
            Return False
        End If

        Try
            myAlreadyInProcessCmdKey = True

            ' Se e' attivo il tracciamento punto-punto, ha la precedenza sul processamento dei comandi da tastiera
            ' NOTA: La ProcessKeyboardKey() ritorna true se il messaggio e' stato gestito

            ' Flag che indica se e' premuto SHIFT
            Dim shiftIsPressed As Boolean = (keyData And Keys.Shift) = Keys.Shift
            Dim ctrlIsPressed As Boolean = (keyData And Keys.Control) = Keys.Control

            Dim msgKey As Keys = CType(msg.WParam.ToInt32, Keys)

            ' Output di debug
            If (msgKey <> Keys.Control) AndAlso (msgKey <> Keys.Shift) AndAlso (msgKey <> Keys.Alt) AndAlso _
               (msgKey <> Keys.ControlKey) AndAlso (msgKey <> Keys.ShiftKey) Then

                '??
            End If

            Select Case msgKey
                Case Keys.Left ' -> Pan Left
                    If shiftIsPressed Then
                        PanLeft(PanFactorWithShift)
                    Else
                        PanLeft(PanFactorNoShift)
                    End If
                    Return True
                Case Keys.Right ' -> Pan Right
                    If shiftIsPressed Then
                        PanRight(PanFactorWithShift)
                    Else
                        PanRight(PanFactorNoShift)
                    End If
                    Return True
                Case Keys.PageDown
                    PageDown()
                Case Keys.PageUp
                    PageUp()
                Case Keys.Up ' -> Pan Up
                    If shiftIsPressed Then
                        PanUp(PanFactorWithShift)
                    Else
                        PanUp(PanFactorNoShift)
                    End If
                    Return True
                Case Keys.Down ' -> Pan Down
                    If shiftIsPressed Then
                        PanDown(PanFactorWithShift)
                    Else
                        PanDown(PanFactorNoShift)
                    End If
                    Return True
                Case Keys.End
                    PanEnd()
                Case Keys.Home
                    PanHome()
                Case Keys.Add ' -> Zoom In
                    ZoomForwardOnLogicalCenter()
                    Return True
                Case Keys.Subtract ' -> Zoom Out
                    ZoomBackOnLogicalCenter()
                    Return True
                Case Keys.Escape ' -> Annullamento azione (o deselezione oggetti)
                    ' Se ho il tasto sinistro del mouse premuto, annullo la trasformazione che stavo facendo
                    ' Se non ho pulsanti premuti, deseleziono eventuali funzioni selezionate
                    Select Case MouseButtons
                        Case Windows.Forms.MouseButtons.Left
                            ' Cancello gli eventuali dati temporanei usati tra un MouseDown e un MouseUp e ridisegno la finestra
                            ResetTemporaryData()
                            Invalidate()
                        Case Windows.Forms.MouseButtons.None
                            ' Deseleziono le funzioni selezionate 
                            ' NOTA: La finestra viene ridisegnata automaticamente

                    End Select
                    Return True
            End Select

            ' Chiamo la routine della classe base
            Return MyBase.ProcessCmdKey(msg, keyData)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myAlreadyInProcessCmdKey = False
        End Try
    End Function


#End Region

#Region "Routine per calcolare/mostrare finestre logiche"

    ''' <summary>
    ''' Mostra la finestra logica desiderata.
    ''' Se SaveInZoomHistory e' true, salva il nuovo zoom nella history degli zoom.
    ''' Se CenterWindow e' true, la finestra viene centrata nella picturebox, altrimenti viene lasciata allineata a sinistra.
    ''' Se AddEmptyBorder e' true lo zoom diminuisce lievemente, in modo da avere una cornice vuota attorno all'area desiderata.
    ''' Se ExcludeRulersArea e' true (e se i righelli sono visibili) la finestra passata viene mappata nell'area della picturebox.
    ''' non coperta dai righelli
    ''' </summary>
    Public Sub ShowLogicalWindow(ByVal LogicalWindow As RECT, _
                                 Optional ByVal CenterWindow As Boolean = True, Optional ByVal AddEmptyBorder As Boolean = True, _
                                 Optional ByVal ExcludeRulersArea As Boolean = True)
        Try
            ' Check se il rettangolo passato e' valido. Se non lo e' faccio lo zoom al rettangolo di default.
            ' NOTA: Se ho una serie di funzioni macchina selezionate, otterro' un ingombro NON VALIDO,
            '       quindi faro' lo zoom al rettangolo di default
            If LogicalWindow.IsZeroSized Then
                ZoomToDefaultRect()
                Exit Sub
            End If

            ' Imposto l'area visualizzabile 
            LogicalArea = VisibleAreaToLogicalArea(LogicalWindow, CenterWindow, AddEmptyBorder, ExcludeRulersArea)

            ' Ridisegno la finestra
            Redraw()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Ritorna l'area logica a cui e' necessario effettuare lo zoom per garantire la visibilita' dell'area desiderata.
    ''' Le due aree non coincidono in quanto l'area ritornata rispetta la proporzione tra altezza e larghezza
    ''' e se necessario, tiene conto dell'area occupata dai righelli e della centratura dell'area desiderata.
    ''' Se CenterWindow e' true, l'area desiderata viene centrata nella picturebox, altrimenti viene lasciata allineata a sinistra
    ''' Se AddEmptyBorder e' true lo zoom diminuisce lievemente, in modo da avere una cornice vuota attorno all'area desiderata
    ''' Se ExcludeRulersArea e' true (e se i righelli sono visibili) l'area desiderata
    ''' viene mappata nella parte della picturebox non coperta dai righelli
    ''' </summary>
    Private Function VisibleAreaToLogicalArea(ByVal visibleArea As RECT, Optional ByVal CenterWindow As Boolean = True, _
                                                Optional ByVal AddEmptyBorder As Boolean = True, Optional ByVal ExcludeRulersArea As Boolean = True) As RECT
        ' Check di sicurezza
        If visibleArea.IsZeroSized Then
            Return New RECT()
        End If

        ' Mi assicuro che il rettangolo sia normalizzato
        visibleArea.NormalizeRect()

        ' Aggiorno l'ultima area visibile che mi e' stata richiesta
        myLastVisibleAreaRequested = visibleArea

        ' Dimensioni orizzontali e verticali del controllo.
        ' NOTA: Uso il ClientRectangle.Width al posto di Width perche' cosi' tengo conto delle eventuali scrollbar.
        Dim clientWidth As Single = Me.ClientRectangle.Width
        Dim clientHeight As Single = Me.ClientRectangle.Height

        ' Spazio occupato dai righelli, e' diverso da zero solo se devo disegnare i righelli 
        ' e se e' attiva l'opzione ExcludeRulersArea
        Dim rulersPhysicalSize As Single = 0

        ' Calcolo lo spazio occupato dai righelli
        ' Serve per impedire che la parte superiore sinistra dell'area richiesta venga coperta dai righelli
        If myShowRulers AndAlso ExcludeRulersArea Then
            rulersPhysicalSize = myRulers.Size
        End If

        ' Se e' stata richiesta una cornice vuota devo aumentare lievemente 
        ' lo spazio richiesto dalla finestra logica sulla coordinata in uso.
        ' Modifico entrambe le coordinate, tanto usero' solo quella che mi serve
        If AddEmptyBorder Then
            Dim widthBorder As Integer = CInt(visibleArea.Width / 18)
            Dim heightBorder As Integer = CInt(visibleArea.Height / 18)
            visibleArea.top -= heightBorder
            visibleArea.bottom += heightBorder
            visibleArea.left -= widthBorder
            visibleArea.right += widthBorder
            Debug.Assert(visibleArea.top = myLastVisibleAreaRequested.top - heightBorder)
            Debug.Assert(visibleArea.bottom = myLastVisibleAreaRequested.bottom + heightBorder)
            Debug.Assert(visibleArea.left = myLastVisibleAreaRequested.left - widthBorder)
            Debug.Assert(visibleArea.right = myLastVisibleAreaRequested.right + widthBorder)
        End If

        ' Spazio disponibile per tracciare.
        ' Se la client area diventa talmente piccola che ci stanno solo i righelli,
        ' faccio comunque in modo da tenermi un pixel per tracciare
        Dim availableWidth As Single = Math.Max(clientWidth - rulersPhysicalSize, 1)
        Dim availableHeight As Single = Math.Max(clientHeight - rulersPhysicalSize, 1)

        ' Fattori di scala corrispondenti alla piu' piccola e alla piu' grande finestra visualizzabile
        Dim minScaleFactor As Single = Math.Min(availableWidth / MinLogicalWindowSize.Width, availableHeight / MinLogicalWindowSize.Height)
        Dim maxScaleFactor As Single = Math.Min(availableWidth / MaxLogicalWindowSize.Width, availableHeight / MaxLogicalWindowSize.Height)
        If availableWidth = 1 Then
            availableWidth = 1
        End If
        ' Trovo i due fattori di scala che mi portano ad avere la finestra desiderata 
        ' a piena dimensione verticale o orizzontale
        Dim horzScaleFactor As Single = availableWidth / visibleArea.Width
        Dim vertScaleFactor As Single = availableHeight / visibleArea.Height

        ' Check se i fattori di scala sono validi
        ' NOTA: Possono diventare nulli quando rimpicciolisco la finestra fino ad arrivare
        '       ad una dimensione pari o minore di quella dei righelli
        ' TODO: Questo e' solo un workaround, in questo caso la visualizzazione diventa sbagliata
        If (horzScaleFactor <= 0) Then horzScaleFactor = maxScaleFactor
        If (vertScaleFactor <= 0) Then vertScaleFactor = maxScaleFactor

        ' Nuovo fattore di scala da usare
        ' Dei due fattori di scala devo prendere quello piu' piccolo, in modo che 
        ' visibleArea.Width e visibleArea.Height staranno dentro all'area finale.
        ' Questo sara' il fattore di scala che la PictureBox avrebbe se visualizzasse la visibleArea finale.
        Dim newScaleFactor As Single = Math.Min(horzScaleFactor, vertScaleFactor)
        Debug.Assert(newScaleFactor <> 0, "Trovato fattore di scala non valido")

        ' Check se il fattore di scala mi porterebbe ad una finestra troppo grande o troppo piccola per essere visualizzata
        If (newScaleFactor > minScaleFactor) Then newScaleFactor = minScaleFactor
        If (newScaleFactor < maxScaleFactor) Then newScaleFactor = maxScaleFactor

        ' Dimensioni dell'area logica che la PictureBox visualizzerebbe con il nuovo fattore di scala
        ' NOTA: Queste dimensioni comprendono l'ingombro degli eventuali righelli
        Dim newLogicalHeight As Single = clientHeight / newScaleFactor
        Dim newLogicalWidth As Single = clientWidth / newScaleFactor

        ' Dimensione logica che i righelli avrebbero con il nuovo fattore di scala
        Dim rulersLogicalSize As Integer = rulersPhysicalSize / newScaleFactor

        ' Offset orizzontale e verticale da sommare all'area visibile
        Dim horizontalOffset As Single = 0
        Dim verticalOffset As Single = 0

        ' Se richesto, calcolo gli offset necessari per il centraggio
        ' NOTA: Faccio entrambi i centraggi, tanto nel caso standard uno dei due offset rimane a zero.
        '       Invece mi servono entrambi i centraggi nel caso in cui tento uno zoom ad un oggetto piccolissimo, cioe' quando supero minScaleFactor
        If CenterWindow Then
            ' NOTA: Questi offset NON comprendono l'ingombro degli eventuali righelli
            verticalOffset = Math.Abs((newLogicalHeight - rulersLogicalSize - visibleArea.Height) / 2)
            horizontalOffset = Math.Abs((newLogicalWidth - rulersLogicalSize - visibleArea.Width) / 2)
        End If

        ' Aggiorno la posizione dell'area da visualizzare 
        Dim logicalAreaToShow As New RECT()
        logicalAreaToShow.left = visibleArea.left - rulersLogicalSize - horizontalOffset
        logicalAreaToShow.top = visibleArea.top - rulersLogicalSize - verticalOffset

        ' Aggiorno entrambe le dimensioni dell'area da visualizzare in modo da riflettere il nuovo fattore di scala
        ' NOTA: Le dimensioni dell'area logica che visualizzo comprendono i righelli,
        '       quindi non devo sottrarre l'ingombro dei righelli
        ' NOTA: Qui va usata Width al posto di ClientRectangle.Width perche' altrimenti il centraggio
        '       non avviene correttamente quando le scrollbar sono visualizzate
        logicalAreaToShow.Width = Me.Width / newScaleFactor
        logicalAreaToShow.Height = Me.Height / newScaleFactor
        logicalAreaToShow.NormalizeRect()

        ' Ritorno l'area logica da visualizzare 
        Return logicalAreaToShow
    End Function

#End Region

#Region "Preview delle trasformazioni"

    ''' <summary>
    ''' Ritorna i fattori di scala da utilizzare in X e in Y per la preview della trasformazione
    ''' </summary>
    Public Function CalculateScaleFactors(ByVal ScalingCenter As Point, ByVal FirstScalePoint As Point, ByVal SecondScalePoint As Point, Optional ByVal mantainAspectRatio As Boolean = False) As System.Drawing.PointF
        Try
            ' Distanza in X e Y
            Dim DeltaX As Double = SecondScalePoint.X - FirstScalePoint.X
            Dim DeltaY As Double = SecondScalePoint.Y - FirstScalePoint.Y

            ' Facio gli eventuali aggiustamenti per mantenere il rapporto tra X e Y
            If mantainAspectRatio Then
                If Math.Abs(DeltaX) > Math.Abs(DeltaY) Then
                    DeltaY = DeltaX * ((FirstScalePoint.Y - ScalingCenter.Y) / (FirstScalePoint.X - ScalingCenter.X))
                Else
                    DeltaX = DeltaY * ((FirstScalePoint.X - ScalingCenter.X) / (FirstScalePoint.Y - ScalingCenter.Y))
                End If
            End If

            ' Valori comuni per la riscalatura in X e in Y, il valore di riscalatura in X e':
            '    DeltaX * ((actualPoint.X - ScalingCenter.X) / (FirstScalePoint.X - ScalingCenter.X))
            ' Siccome actualPoint.X varia da punto a punto, la parte comune a tutti i punti e': 
            '    DeltaX * 1.0 / (FirstScalePoint.X - ScalingCenter.X)
            ' che viene conglobata nella CommonFactorX
            ' Un analogo discorso vale per la CommonFactorY
            Dim commonFactorX As Double = 0.0
            If FirstScalePoint.X <> ScalingCenter.X Then
                commonFactorX = DeltaX / (FirstScalePoint.X - ScalingCenter.X)
            End If
            Dim commonFactorY As Double = 0.0
            If FirstScalePoint.Y <> ScalingCenter.Y Then
                commonFactorY = DeltaY / (FirstScalePoint.Y - ScalingCenter.Y)
            End If

            ' La scrittura originale era:
            '    ScaledXRatio = (FirstScalePoint.X + DeltaX - ScalingCenter.X) / (FirstScalePoint.X - ScalingCenter.X)
            ' Da cui si ottiene:
            '    ScaledXRatio = (FirstScalePoint.X - ScalingCenter.X + DeltaX) / (FirstScalePoint.X - ScalingCenter.X)
            '                 = (FirstScalePoint.X - ScalingCenter.X) / (FirstScalePoint.X - ScalingCenter.X) + DeltaX / (FirstScalePoint.X - ScalingCenter.X)
            '                 = 1.0 + DeltaX / (FirstScalePoint.X - ScalingCenter.X)
            '                 = 1.0 + commonFactorX 
            ' Un analogo discorso vale per la preview.ScaledYRatio
            Return New System.Drawing.PointF(1.0 + commonFactorX, 1.0 + commonFactorY)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return New System.Drawing.PointF(1.0, 1.0)
        End Try
    End Function
#End Region

#Region "Inizializzazione e finalizzazione"

    ''' <summary>
    ''' Routine di inizializzazione
    ''' </summary>
    Private Sub Initialize()
        Try
            ' Cursore a croce utilizzato per visualizzare lo spostamento del mouse
            FullCrossCursor.CoordinatesBox = myCoordinatesBox
            ' Cursore a croce generico
            FullCrossCursor.CoordinatesBox = myCoordinatesBox

            ZoomToDefaultRect()

            ' Aggiungo gli handler per la gestione degli eventi di resize della finestra
            AddResizeHandlers()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Ripulisce la memoria
    ''' </summary>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        ' NOTA: Non e' necessario implementare una Dispose() senza parametri e una Finalize(),
        '       perche' attraverso la derivazione da UserControl, il controllo implementa gia' le funzioni:
        '
        '       Public Overloads Sub Dispose() implements IDisposable.Dispose()
        '          Dispose(True)
        '          GC.SuppressFinalize(Me)
        '       End Sub
        '
        '       Protected Overrides Sub Finalize()
        '          Try
        '             Me.Dispose(False)
        '          Finally
        '             MyBase.Finalize()
        '          End Try
        '       End Sub
        '
        '       Inoltre la chiamata a MyBase.Dispose() garantisce che tutti i controlli figli ecc. vengano disallocati

        ' Check se la Dispose() era gia' stata chiamata.
        ' NOTA: Le linee guide del .NET stabiliscono che non si devono generare errori
        '       in caso di chiamate multiple,semplicemente le chiamate successive vanno ignorate.
        If Me.IsDisposed Then
            Return
        End If

        Try
            ' If disposing equals true, the method has been called directly
            ' or indirectly by a user's code. Managed and unmanaged resources
            ' can be disposed.
            If disposing Then
                ' Pulizia menu di contesto e simili
                If (components IsNot Nothing) Then
                    components.Dispose()
                End If
            End If

            ' If disposing equals false, the method has been called by the 
            ' runtime from inside the finalizer and you should not reference 
            ' other objects. Only unmanaged resources can be disposed.

            ' Rimuovo gli handler per la gestione degli eventi di resize della finestra
            RemoveResizeHandlers()
        Catch ex As Exception
            Debug.Print("WARNING: Exception occurred while disposing object of type: " + Me.GetType.ToString())
        End Try

        MyBase.Dispose(disposing)
    End Sub

#Region "Funzioni per la gestione della logica di layout"

    ''' <summary>
    ''' Consente di sospendere temporaneamente la logica di layout per il controllo
    ''' </summary>
    Private Overloads Sub SuspendLayout()
        MyBase.SuspendLayout()
        Me.myIsLayoutSuspended = True
    End Sub

    ''' <summary>
    ''' Consente di riprendere la consueta logica di layout
    ''' </summary>
    Private Overloads Sub ResumeLayout()
        MyBase.ResumeLayout()
        Me.myIsLayoutSuspended = False
        ' Faccio una Redraw() quando ho concluso l'aggiornamento del layout
        Redraw()
    End Sub

    ''' <summary>
    ''' Consente di riprendere la consueta logica di layout, imponendo, eventualmente, l'esecuzione di un layout immediato delle richieste di layout in sospeso.
    ''' </summary>
    Private Overloads Sub ResumeLayout(ByVal performLayout As Boolean)
        MyBase.ResumeLayout(performLayout)
        Me.myIsLayoutSuspended = False
        ' Faccio una Redraw() quando ho concluso l'aggiornamento del layout
        Redraw()
    End Sub

#End Region

#End Region

#Region "Routine per la scalatura dei Graphics"

    ''' <summary>
    ''' Ritorna un oggetto graphics derivato dalla bitmap passatagli
    ''' L'oggetto ritornato ha le matrici di scalatura e traslazione impostate sul fattore di scala e sulla LogicalOrigin attualmente in uso.
    ''' NOTA BENE: E' necessario fare un Dispose() del Graphics ritornato appena si e' finito di utilizzarlo
    ''' </summary>
    Protected Friend Function GetScaledGraphicObject(ByVal Src As Bitmap) As Graphics
        ' Check se la sorgente passatami e' valida
        If Src Is Nothing Then Return Nothing

        ' Creo il Graphics, lo scalo e lo ritorno
        Return ScaleGraphicObject(Graphics.FromImage(Src))
    End Function

    ''' <summary>
    ''' Ritorna un oggetto graphics con le matrici di scalatura e traslazione impostate 
    ''' sul fattore di scala e sulla LogicalOrigin attualmente in uso.
    ''' </summary>
    Protected Friend Function ScaleGraphicObject(ByRef GR As Graphics) As Graphics
        Try
            ' Check se il Graphics passatomi e' valido
            If GR Is Nothing Then
                Return Nothing
            End If

            ' Check che il fattore di scala sia valido
            If (ScaleFactor <= 0.0) Then
                MsgBox("Fattore di scala non valido in ScaleGraphicObject()")
                Return GR
            End If

            ' Cancello eventuali trasformazioni precedenti
            GR.ResetTransform()

            ' Per ottenere una traslazione di coordinate logiche va impostata prima la matrice di scala e poi quella di traslazione 
            GR.ScaleTransform(ScaleFactor, ScaleFactor)
            GR.TranslateTransform(-LogicalOrigin.X, -LogicalOrigin.Y)
            Return GR
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Ritorna un oggetto graphics con le matrici di scalatura e traslazione impostate 
    ''' sul fattore di scala e sulla LogicalOrigin attualmente in uso.
    ''' </summary>
    Public Function GetGraphics() As Graphics
        Try
            Dim GR As Graphics = Graphics.FromImage(myRefreshBackBuffer)

            ' Check che il fattore di scala sia valido
            If (ScaleFactor <= 0.0) Then
                MsgBox("Fattore di scala non valido in ScaleGraphicObject()")
                Return GR
            End If

            ' Cancello eventuali trasformazioni precedenti
            GR.ResetTransform()

            ' Per ottenere una traslazione di coordinate logiche va impostata prima la matrice di scala e poi quella di traslazione 
            GR.ScaleTransform(ScaleFactor, ScaleFactor)
            GR.TranslateTransform(-LogicalOrigin.X, -LogicalOrigin.Y)
            Return GR
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region


#Region "Metodi"
    Private Function MyDoubleClick() As Boolean
        Static myTimer As Long
        If (Now.Ticks - myTimer) < 5000000 Then
            Return True
        End If
        myTimer = Now.Ticks
        Return False
    End Function

#End Region

End Class
