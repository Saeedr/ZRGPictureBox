Imports System.Drawing.Drawing2D
Imports System.Collections.Generic
Partial Public Class ZRGPictureBoxControl

    Private Class Rulers

#Region "Costanti"
        Const RulerSize As Integer = 20
        Private Const FreeSpaceFactor As Double = 1.75
        Private Const RulerColorAlpha As Integer = 130
#End Region

#Region "Tipi di dati privati"
        Private Class DrawNumberBitmap

#Region "Costanti"

            ''' <summary>
            ''' Larghezza di una singola cifra [pixel]
            ''' </summary>
            Private Const DigitWidth As Integer = 6

#End Region

#Region "Variabili shared"

            ''' <summary>
            ''' Tabella contenente gli array di coordinate relativi alle varie cifre (e segni)
            ''' </summary>
            Shared signsTable As New Dictionary(Of Char, Point())

#End Region

#Region "Funzioni private"

            ''' <summary>
            ''' Inizializza gli array di pixel corrispondenti alle varie cifre
            ''' </summary>
            Private Shared Sub Initialize()

                ' Array delle coordinate associate alle cifre (o ai segni) da tracciare
                ' Ogni coppia di coordinate indica uno degli estremi dei segmenti 
                ' che devo tracciare per ottenere la cifra (o il segno) voluto

                ' Cifra "1"
                Dim one As Point() = {New Point(0, 2), New Point(2, 0), New Point(2, 7)}

                ' Cifra "2"
                Dim two As Point() = {New Point(0, 1), New Point(1, 0), New Point(3, 0), _
                                      New Point(4, 1), New Point(4, 3), New Point(0, 7), _
                                      New Point(4, 7)}

                ' Cifra "3"
                Dim three As Point() = {New Point(0, 1), New Point(1, 0), New Point(3, 0), _
                                        New Point(4, 1), New Point(4, 2), New Point(3, 3), _
                                        New Point(2, 3), New Point(3, 3), New Point(4, 4), _
                                        New Point(4, 6), New Point(3, 7), New Point(1, 7), _
                                        New Point(0, 6)}

                ' Cifra "4"
                Dim four As Point() = {New Point(4, 5), New Point(0, 5), New Point(0, 4), _
                                       New Point(2, 1), New Point(3, 0), New Point(3, 7)}

                ' Cifra "5"
                Dim five As Point() = {New Point(4, 0), New Point(1, 0), New Point(1, 1), _
                                       New Point(0, 2), New Point(0, 3), New Point(3, 3), _
                                       New Point(4, 4), New Point(4, 6), New Point(3, 7), _
                                       New Point(1, 7), New Point(0, 6)}

                ' Cifra "6"
                Dim six As Point() = {New Point(3, 0), New Point(1, 0), New Point(0, 1), _
                                      New Point(0, 6), New Point(1, 7), New Point(3, 7), _
                                      New Point(4, 6), New Point(4, 4), New Point(3, 3), _
                                      New Point(0, 3)}

                ' Cifra "7"
                Dim seven As Point() = {New Point(0, 0), New Point(4, 0), New Point(1, 7)}

                ' Cifra "8"
                Dim eight As Point() = {New Point(3, 0), New Point(1, 0), New Point(0, 1), _
                                        New Point(0, 2), New Point(1, 3), New Point(3, 3), _
                                        New Point(4, 4), New Point(4, 6), New Point(3, 7), _
                                        New Point(1, 7), New Point(0, 6), New Point(0, 4), _
                                        New Point(1, 3), New Point(3, 3), New Point(4, 2), _
                                        New Point(4, 1), New Point(3, 0)}

                ' Cifra "9"
                Dim nine As Point() = {New Point(0, 6), New Point(1, 7), New Point(3, 7), _
                                       New Point(4, 6), New Point(4, 1), New Point(3, 0), _
                                       New Point(1, 0), New Point(0, 1), New Point(0, 3), _
                                       New Point(1, 4), New Point(4, 4)}

                ' Cifra "0"
                Dim zero As Point() = {New Point(1, 0), New Point(3, 0), New Point(4, 1), _
                                       New Point(4, 6), New Point(3, 7), New Point(1, 7), _
                                       New Point(0, 6), New Point(0, 1), New Point(1, 0)}

                ' Segno "-"
                Dim minus As Point() = {New Point(1, 3), New Point(4, 3)}

                ' Segno "." e segno ","
                Dim dot As Point() = {New Point(2, 6), New Point(3, 6), New Point(2, 7), New Point(3, 7)}

                ' Aggiungo i vari array alla tabella
                signsTable.Add("1", one)
                signsTable.Add("2", two)
                signsTable.Add("3", three)
                signsTable.Add("4", four)
                signsTable.Add("5", five)
                signsTable.Add("6", six)
                signsTable.Add("7", seven)
                signsTable.Add("8", eight)
                signsTable.Add("9", nine)
                signsTable.Add("0", zero)
                signsTable.Add("-", minus)
                signsTable.Add(".", dot)
                signsTable.Add(",", dot)
            End Sub

            ''' <summary>
            ''' Ritorna una serie di coordinate dei segmenti da tracciare per ottenere una rappresentazione grafica del numero passatogli.
            ''' </summary>
            Private Sub CreateSegmentsList(ByVal Value As Double, ByRef pointList As List(Of System.Drawing.Point), ByVal Horizontal As Boolean, Optional ByVal HideSign As Boolean = False)
                Try
                    ' Check se devo ancora inizializzare gli array di pixel delle singole cifre
                    If signsTable.Count = 0 Then
                        Initialize()
                    End If

                    ' Cancello la lista di punti
                    pointList.Clear()

                    ' Converto il valore in stringa 
                    Dim strValue As String = ValueString(Value)

                    ' Offset necessario per allineare la scritta al centro 
                    Dim alignmentOffset As Integer = -MaskWidth(Value) / 2

                    ' Array che conterra' i dati del carattere attuale
                    Dim actualSign As Point() = Nothing
                    Dim actualChar As Char

                    ' Scandisco tutti i carateri costituenti il numero da stampare
                    For actualIndex As Integer = 0 To strValue.Length - 1
                        actualChar = strValue(actualIndex)

                        ' Se richiesto, salto il segno meno
                        If HideSign AndAlso actualChar = "-" Then
                            Continue For
                        End If

                        ' Recupero la tabella di coordinate associata al carattere attuale
                        actualSign = signsTable(actualChar)

                        Dim newPoint As System.Drawing.Point
                        Dim xCoord, yCoord As Integer
                        pointList.Capacity = pointList.Count + actualSign.Length + 1
                        For i As Integer = 0 To actualSign.Length - 1
                            ' Calcolo le coordinate del nuovo punto basandomi sul template
                            xCoord = (DigitWidth * actualIndex) + actualSign(i).X + alignmentOffset
                            yCoord = actualSign(i).Y
                            ' Se la scritta e' verticale, scambio le coordinate
                            If Horizontal Then
                                newPoint = New System.Drawing.Point(xCoord, yCoord)
                            Else
                                newPoint = New System.Drawing.Point(yCoord, -xCoord)
                            End If
                            ' Aggiungo il punto alla lista
                            pointList.Add(newPoint)
                        Next
                        ' Il punto con X e Y a maxvalues verra' ignorato 
                        pointList.Add(New System.Drawing.Point(Integer.MaxValue, Integer.MaxValue))
                    Next

                Catch ex As Exception
                    MsgBox(ex.Message + vbCr + ex.StackTrace)
                End Try
            End Sub

#End Region

#Region "Proprieta'"

            ''' <summary>
            ''' Larghezza che avra' la maschera finale del valore passato [pixel]
            ''' </summary>
            Public ReadOnly Property MaskWidth(ByVal aValue As Double) As Integer
                Get
                    ' Larghezza di una cifra [pixel] * "numero di cifre nella rappresentazione come stringa"
                    Return DigitWidth * ValueString(aValue).Length
                End Get
            End Property

            ''' <summary>
            ''' Ritorna la stringa da stampare per il valore passatogli
            ''' </summary>
            Private ReadOnly Property ValueString(ByVal aValue As Double) As String
                Get
                    ' Converto in stringa in modo che mantenga almeno uno zero prima dei decimali
                    ' NOTA: Le 3 cifre dopo la virgola servono solo quando uso le inches e vado in zoom molto alti.
                    '       In realta' per il livello di zoom massimo implementato nella PictureBox avrei 4 cifre
                    '       dopo la virgola, ma la quarta cifra non e' molto precisa, quindi non la stampo
                    Return aValue.ToString("0.###")
                End Get
            End Property

#End Region


            Public Sub DrawScaledNumber(ByVal GR As Graphics, ByVal value As Double, ByVal xCoord As Single, ByVal yCoord As Single, ByVal ScaleFactor As Single, ByVal Horizontal As Boolean)
                Try
                    ' Calcolo le coordinate dei segmenti da tracciare
                    Static pixelList As New List(Of System.Drawing.Point)
                    CreateSegmentsList(value, pixelList, Horizontal)

                    ' Lista temporanea dei segmenti in coordinate logiche
                    Static logicCoordList As New List(Of Point)
                    logicCoordList.Clear()

                    Dim tmpLogicPoint As Point
                    For iIter As Integer = 0 To pixelList.Count - 1
                        ' Il tag a zero indica che sto processando i segmenti relativi ad una cifra
                        ' Il tag messo a 1 segnala che e' finita una cifra (o un segno)
                        If pixelList(iIter).X <> Integer.MaxValue Then
                            ' Tutti i segmenti costituenti una cifra vanno convertiti in coordinate logiche
                            tmpLogicPoint.X = pixelList(iIter).X / ScaleFactor + xCoord
                            tmpLogicPoint.Y = pixelList(iIter).Y / ScaleFactor + yCoord
                            ' Poi li salvo nella lista
                            logicCoordList.Add(tmpLogicPoint)
                        Else
                            ' Quando e' finita una cifra (o un segno), disegno l'array di segmenti corrispondente
                            GR.DrawLines(RulerPen, logicCoordList.ToArray())
                            ' Poi resetto la lista per la prossima cifra
                            logicCoordList.Clear()
                        End If
                    Next
                Catch ex As Exception
                    MsgBox(ex.Message + vbCr + ex.StackTrace)
                End Try
            End Sub

        End Class

#End Region

#Region "Variabili private"

        Private WithEvents myPictureBoxControl As ZRGPictureBoxControl

#Region "Variabili shared"

        Private Shared RulerPen As New Pen(Color.Navy)

        ''' <summary>
        ''' Penna che uso per disegnare le righe di drag and drop dei righelli
        ''' </summary>
        Private Shared myDragPen As Pen = Nothing

        ''' <summary>
        ''' Bitmap che mi serve per l'origine
        ''' </summary>
        Private Shared myOriginBmp As Image = Nothing

        ''' <summary>
        ''' Bitmap che mi serve per l'origine
        ''' </summary>
        Private Shared myOriginBmpSnapped As Bitmap = Nothing

        ''' <summary>
        ''' Variabile utilizzata per la creazione delle maschere di pixel.
        ''' NOTA: Ogni cifra che verra' disegnata nei righelli richiede la creazione di una maschera
        ''' </summary>
        Private Shared digitMaskCreator As New DrawNumberBitmap()

#End Region

#Region "Penne, pennelli e bitmap"

        ''' <summary>
        ''' Bitmap contenente il righello orizzontale
        ''' </summary>
        Private myHRulerBmp As Bitmap = Nothing

        ''' <summary>
        ''' Bitmap contenente il righello verticale
        ''' </summary>
        Private myVRulerBmp As Bitmap = Nothing

#End Region

#Region "Colori"

        ''' <summary>
        ''' Colore della penna che uso per disegnare le righe di drag and drop dei righelli
        ''' </summary>
        Dim myDragLineColor As Color = Color.Black

#End Region

#Region "Dimensioni e unita' di misura"

        ''' <summary>
        ''' Ampiezza dei righelli [pixel]
        ''' </summary>
        Private mySize As Integer = RulerSize
#End Region

#Region "Gestione del ridisegno"

        ''' <summary>
        ''' Ultimi dati grafici (origine, dimensioni, fattore di scala, ecc.) con cui ho ridisegnato la bitmap dei righelli.
        ''' </summary>
        Private myLastGraphicInfo As ConversionInfo = Nothing


        ''' <summary>
        ''' Flag usato per indicare quando c'e' bisogno di un ridisegno del righello orizzontale
        ''' </summary>
        Private NeedsHorizontalRedraw As Boolean = True

        ''' <summary>
        ''' Flag usato per indicare quando c'e' bisogno di un ridisegno del righello verticale
        ''' </summary>
        Private NeedsVerticalRedraw As Boolean = True

#End Region

#End Region

#Region "Proprieta'"

        ''' <summary>
        ''' PictureBox associata a questa istanza delle classe
        ''' </summary>
        Public Property PictureBoxControl() As ZRGPictureBoxControl
            Get
                Return myPictureBoxControl
            End Get
            Private Set(ByVal value As ZRGPictureBoxControl)
                myPictureBoxControl = value
            End Set
        End Property
        ''' <summary>
        ''' Unita' di misura usata nel righello
        ''' </summary>
        Public ReadOnly Property UnitOfMeasure() As MeasureSystem.enUniMis
            Get
                Return PictureBoxControl.UnitOfMeasure
            End Get
        End Property

        ''' <summary>
        ''' Larghezza del righello orizzontale [pixel]
        ''' </summary>
        Public ReadOnly Property Width() As Integer
            Get
                Return PictureBoxControl.Width
            End Get
        End Property

        ''' <summary>
        ''' Altezza del righello verticale [pixel]
        ''' </summary>
        Public ReadOnly Property Height() As Integer
            Get
                Return PictureBoxControl.Height
            End Get
        End Property

        ''' <summary>
        ''' Ampiezza dei righelli [pixel]
        ''' </summary>
        Public Property Size() As Integer
            Get
                Return mySize
            End Get
            Set(ByVal value As Integer)
                If mySize = value Then
                    Exit Property
                End If
                mySize = value
                ' Ho cambiato una dimensione di entrambi i righelli, quindi devo ricrearne le bitmap da zero
                myHRulerBmp = Nothing
                myVRulerBmp = Nothing
            End Set
        End Property

        ''' <summary>
        ''' Larghezza del righello orizzontale [dimensioni logiche]
        ''' </summary>
        Public ReadOnly Property LogicalWidth() As Integer
            Get
                If ScaleFactor <> 0 Then
                    Return Width / ScaleFactor
                End If
                Return 0
            End Get
        End Property

        ''' <summary>
        ''' Altezza del righello verticale [dimensioni logiche]
        ''' </summary>
        Public ReadOnly Property LogicalHeight() As Integer
            Get
                If ScaleFactor <> 0 Then
                    Return Height / ScaleFactor
                End If
                Return 0
            End Get
        End Property

        ''' <summary>
        ''' Ampiezza dei righelli [dimensioni logiche]
        ''' </summary>
        Public ReadOnly Property LogicalSize() As Integer
            Get
                If ScaleFactor <> 0 Then
                    Return Size / ScaleFactor
                End If
                Return 0
            End Get
        End Property

        ''' <summary>
        ''' Ritorna il fattore di scala con cui va disegnato questo oggetto
        ''' </summary>
        Public ReadOnly Property ScaleFactor() As Single
            Get
                Return PictureBoxControl.ScaleFactor
            End Get
        End Property

        ''' <summary>
        ''' Punto disegnato in alto a sinistra dello schermo [coordinate logiche]
        ''' </summary>
        Public ReadOnly Property LogicalOrigin() As Point
            Get
                Return PictureBoxControl.LogicalOrigin
            End Get
        End Property

        ''' <summary>
        ''' Colore del righello
        ''' </summary>
        Public ReadOnly Property RulerColor() As Color
            Get
                Return Color.FromArgb(RulerColorAlpha, Color.LightYellow)
            End Get
        End Property
        ''' <summary>
        ''' Ritorna la larghezza che la bitmap deve avere per disegnare correttamente il righello orizzontale sulla PictureBox
        ''' </summary>
        Private ReadOnly Property NeededBitmapWidth() As Integer
            Get
                ' Devo avere una bitmap di dimensione pari a quello che devo disegnare
                Dim newWidth As Integer = PictureBoxControl.Width
                ' Per evitare di allocare e deallocare continuamente delle bitmap quando l'utente ridimensiona la finestra
                ' trascinando con il mouse, le bitmap vengono allocate arrotondando ai 100 pixel superiori
                newWidth = Math.Ceiling(CDbl(newWidth) / 100.0) * 100
                Return newWidth
            End Get
        End Property

        ''' <summary>
        ''' Ritorna l'altezza che la bitmap deve avere per disegnare correttamente il righello verticale sulla PictureBox
        ''' </summary>
        Private ReadOnly Property NeededBitmapHeight() As Integer
            Get
                ' Controllo se ho una pictureBox collegata a questi righelli
                ' Devo avere una bitmap di dimensione pari a quello che devo disegnare
                Dim newHeight As Integer = PictureBoxControl.Height
                ' Per evitare di allocare e deallocare continuamente delle bitmap quando l'utente ridimensiona la finestra
                ' trascinando con il mouse, le bitmap vengono allocate arrotondando ai 100 pixel superiori
                newHeight = Math.Ceiling(CDbl(newHeight) / 100.0) * 100
                Return newHeight
            End Get
        End Property

        ''' <summary>
        ''' Ritorna la dimensione comune che la bitmap deve avere per disegnare correttamente i righelli.
        ''' </summary>
        Private ReadOnly Property NeededBitmapRulerSize() As Integer
            Get
                ' Devo avere una bitmap di dimensione pari a quello che devo disegnare
                Return Size
            End Get
        End Property

#End Region

#Region "Costruttori"

        ''' <summary>
        ''' Costruttore dato un controllo su cui disegnare il righello
        ''' </summary>
        Public Sub New(ByVal pictureBox As ZRGPictureBoxControl)
            myPictureBoxControl = pictureBox

            ' Creo la penna che uso per disegnare le righe di drag and drop dei righelli
            If myDragPen Is Nothing Then
                CreateDragPen()
            End If

            ' Carico la bitmap che mi serve per l'origine
            myOriginBmp = LoadImageRes("Rulers.RulerOrigin.png")
            myOriginBmpSnapped = LoadImageRes("Rulers.RulerOriginSnap.png")

            ' Aggiorno la dimensione dei righelli in modo che siano coerenti con la dimensione della bitmap
            Size = Math.Max(myOriginBmp.Width, myOriginBmp.Height)
        End Sub

#End Region

#Region "Funzioni private"
        Private Function LoadImageRes(ByVal imageName As String) As Image
            Try
                Dim thisAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim assemblyName As String = thisAssembly.GetName.Name
                imageName = assemblyName + "." + imageName
                Static resourcesNames As String() = thisAssembly.GetManifestResourceNames()
                Dim found As Boolean = False
                For Each name As String In resourcesNames
                    If name.ToLower = imageName.ToLower Then 'Case insensitive
                        Dim file As System.IO.Stream = thisAssembly.GetManifestResourceStream(name)
                        Return Image.FromStream(file)
                    End If
                Next
                Return Nothing
            Catch ex As Exception
                MsgBox(ex.Message + vbCr + ex.StackTrace)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Ritorna il valore del passo necessario pwer visualizzare correttamente le quote su di un righello
        ''' </summary>
        Private Function CalculateBaseStep(ByVal horizontalRuler As Boolean) As Integer

            ' NOTA: Si suppone che il numero piu' largo cada ad uno degli estremi
            '       dell'intervallo di valori visualizzato. In realta' questo non e' vero,
            '       perche' se ho numeri con la virgola (con zoom alti e inches, ad esempio)
            '       posso avere come estremi 2 e 3, ed in mezzo trovarmi con 2.555 e 2.666.
            '       Comunque scelgo di ignorare questo fatto, tanto sono riuscito a compensare
            '       questo problema con il freeSpaceFactor

            ' NOTA2: Al limite, potrei provare a fare un test anche su di un numero a meta'
            '        della finestra logica, ma rischio che abbia un numero di cifre troppo elevato
            '        rispetto alle cifre effettivamente stampate

            ' Trovo la dimensione in pixel del numero piu' grande che devo visualizzare
            Dim startValue, stopValue, maxNumberWidth As Single

            ' Valore all'inizio e alla fine del righello
            If horizontalRuler Then
                startValue = LogicalOrigin.X
                stopValue = LogicalOrigin.X + LogicalWidth
            Else
                startValue = LogicalOrigin.Y
                stopValue = LogicalOrigin.Y + LogicalHeight
            End If

            ' Aggiorno il valore in modo da tener conto della misura in cui sto visualizzando i dati
            ' NOTA: Come al solito, il valore di partenza e' in micron
            startValue /= MeasureSystem.CustomUnitToMicron(1, UnitOfMeasure)
            stopValue /= MeasureSystem.CustomUnitToMicron(1, UnitOfMeasure)


            ' Dimensione del numero all'inizio e numero alla fine del righello
            startValue = digitMaskCreator.MaskWidth(startValue)
            stopValue = digitMaskCreator.MaskWidth(stopValue)

            ' Massima dimensione del numero da scrivere [pixel]
            maxNumberWidth = Math.Max(startValue, stopValue)

            ' Spazio disponibile [pixel]
            Dim availableSpace As Integer
            If horizontalRuler Then
                availableSpace = CInt(PictureBoxControl.GraphicInfo.ToPhysicalDimension(LogicalWidth))
            Else
                availableSpace = CInt(PictureBoxControl.GraphicInfo.ToPhysicalDimension(LogicalHeight))
            End If

            ' Numero di quote che si suppone che vengano visualizzate nel righello
            Dim totalNumQuotes As Integer = availableSpace / CInt(maxNumberWidth)

            ' Check se ho un numero di quote valido
            If totalNumQuotes = 0 Then
                Return 0
            End If

            ' Trovo il passo di base con cui scrivero' i numeri nel righello [micron]
            If horizontalRuler Then
                Return LogicalWidth / totalNumQuotes
            Else
                Return LogicalHeight / totalNumQuotes
            End If
        End Function

        ''' <summary>
        ''' Ridisegna la bitmap che rappresenta il righello orizzontale
        ''' </summary>
        Private Function RedrawHorizontalRuler() As Bitmap
            Dim GR As Graphics = Nothing
            Try
                ' Check se questo righello ha dimensioni valide
                If Width <= 0 Then
                    Return Nothing
                End If

                ' Imposto il Graphics su cui disegnare e la dimensione logica del righello
                ' NOTA: La dimensione logica deve essere un pixel di meno rispetto alla dimensione fisica della bitmap,
                '       altrimenti la linea che vado a disegnare sul contorno interno non risulta visibile.
                ' NOTA: Il fatto di diminuire queso valore non influlenza la posizione delle tacche del righello, ma solo la loro "altezza"
                GR = PictureBoxControl.GetScaledGraphicObject(myHRulerBmp)
                Dim rulerLogicSize As Integer = (Size - 1) / ScaleFactor

                ' Disegno lo sfondo del righello
                GR.Clear(RulerColor)


                ' Disegno il contorno del righello
                ' NOTA: Non devo disegnare un rettangolo completo attorno al righello, perhce' altrimenti il bordo rivolto 
                '       verso l'esterno mi da' un brutto effetto ottico quando la PictureBox e' in modalita' BorderStyle.Simple.
                '       Quindi disegno solo una riga nera rivolta verso l'interno della PictureBox

                GR.DrawLine(RulerPen, LogicalOrigin.X + rulerLogicSize, LogicalOrigin.Y + rulerLogicSize, LogicalOrigin.X + LogicalWidth, LogicalOrigin.Y + rulerLogicSize)

                ' Fattore moltiplicativo necessario per disegnare il giusto numero sulla quota
                Dim RulerValueFactor As Single = 1.0 / MeasureSystem.CustomUnitToMicron(1, UnitOfMeasure)


                ' Calcolo il passo del righello
                Dim rulerStep As Single = GetRulerStep()
                If (rulerStep <= 0) Then
                    Return myHRulerBmp
                End If


                Dim XDisplacement As Integer = 0
                Dim OverNeedles As Integer = 0

                XDisplacement = 0
                OverNeedles = 1


                ' Trovo il primo punto da disegnare che non cade sotto al righello verticale
                Dim startPoint As Single = Math.Ceiling((LogicalOrigin.X + rulerLogicSize) / rulerStep) * rulerStep + XDisplacement

                ' Valori della Y che restano fissi: per disegnare il numero, la linea alta 1/2 righello,
                ' la linea alta 1/4 di righello, la linea di base del righello
                Dim yCoord As Integer = LogicalOrigin.Y + 2 / ScaleFactor
                Dim yHalfLine As Integer = LogicalOrigin.Y + rulerLogicSize / 2
                Dim yQuarterLine As Integer = LogicalOrigin.Y + rulerLogicSize - rulerLogicSize / 4
                Dim yRulerBase As Integer = LogicalOrigin.Y + rulerLogicSize


                ' Disegno l'interno del righello
                For xCoord As Single = startPoint To LogicalOrigin.X + LogicalWidth Step rulerStep
                    ' Disegno le due lineette verticali
                    GR.DrawLine(RulerPen, xCoord, yHalfLine, xCoord, yRulerBase)
                    GR.DrawLine(RulerPen, CInt(xCoord + rulerStep / 2), yQuarterLine, CInt(xCoord + rulerStep / 2), yRulerBase)
                    ' Disegno la quota sul righello
                    'GR.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    digitMaskCreator.DrawScaledNumber(GR, (xCoord - XDisplacement) * RulerValueFactor - OverNeedles + 1, xCoord, yCoord, ScaleFactor, True)
                    'GR.SmoothingMode = Drawing2D.SmoothingMode.None
                Next

                ' Il righello orizzontale e' aggiornato
                NeedsHorizontalRedraw = False

                Return myHRulerBmp
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Nothing
            Finally
                ' Libero la memoria allocata per l'oggetto Graphics
                If GR IsNot Nothing Then
                    GR.Dispose()
                End If
            End Try
        End Function
        ''' <summary>
        ''' Ridisegna la bitmap che rappresenta il righello verticale
        ''' </summary>
        Private Function RedrawVerticalRuler() As Bitmap
            Dim GR As Graphics = Nothing
            Try
                ' Check se questo righello ha dimensioni valide
                If Height <= 0 Then
                    Return Nothing
                End If

                ' Imposto il Graphics su cui disegnare e la dimensione logica del righello
                ' NOTA: La dimensione logica deve essere un pixel di meno rispetto alla dimensione fisica della bitmap,
                '       altrimenti la linea che vado a disegnare sul contorno interno non risulta visibile.
                ' NOTA: Il fatto di diminuire queso valore non influlenza la posizione delle tacche del righello, ma solo la loro "altezza"
                GR = PictureBoxControl.GetScaledGraphicObject(myVRulerBmp)
                Dim rulerLogicSize As Integer = (Size - 1) / ScaleFactor

                ' Disegno lo sfondo del righello
                GR.Clear(RulerColor)

                ' Disegno il contorno del righello
                ' NOTA: Non devo disegnare un rettangolo completo attorno al righello, perhce' altrimenti il bordo rivolto 
                '       verso l'esterno mi da' un brutto effetto ottico quando la PictureBox e' in modalita' BorderStyle.Simple.
                '       Quindi disegno solo una riga nera rivolta verso l'interno della PictureBox

                GR.DrawLine(RulerPen, LogicalOrigin.X + rulerLogicSize, LogicalOrigin.Y + rulerLogicSize, LogicalOrigin.X + rulerLogicSize, LogicalOrigin.Y + LogicalHeight)

                ' Fattore moltiplicativo necessario per disegnare il giusto numero sulla quota
                Dim RulerValueFactor As Single = 1.0 / MeasureSystem.CustomUnitToMicron(1, UnitOfMeasure)

                ' Calcolo il passo del righello
                Dim rulerStep As Single = GetRulerStep()

                ' Check se ho un passo valido
                If (rulerStep <= 0) Then
                    Return myVRulerBmp
                End If

                ' Trovo il primo punto da disegnare che non cade sotto al righello orizzontale
                Dim startPoint As Integer = Math.Ceiling((LogicalOrigin.Y + rulerLogicSize) / rulerStep) * rulerStep

                ' Valori della X che restano fissi: per disegnare il numero, la linea alta 1/2 righello,
                ' la linea alta 1/4 di righello, la linea di base del righello
                Dim xCoord As Integer = LogicalOrigin.X + 2 / ScaleFactor
                Dim xHalfLine As Integer = LogicalOrigin.X + rulerLogicSize / 2
                Dim xQuarterLine As Integer = LogicalOrigin.X + rulerLogicSize - rulerLogicSize / 4
                Dim xRulerBase As Integer = LogicalOrigin.X + rulerLogicSize

                ' Disegno l'interno del righello
                For yCoord As Single = startPoint To LogicalOrigin.Y + LogicalHeight Step rulerStep
                    ' Disegno le due lineette orizzontali
                    GR.DrawLine(RulerPen, xHalfLine, yCoord, xRulerBase, yCoord)
                    GR.DrawLine(RulerPen, xQuarterLine, CInt(yCoord + rulerStep / 2), xRulerBase, CInt(yCoord + rulerStep / 2))
                    ' Disegno la quota sul righello
                    'GR.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    digitMaskCreator.DrawScaledNumber(GR, yCoord * RulerValueFactor, xCoord, yCoord, ScaleFactor, False)
                    'GR.SmoothingMode = Drawing2D.SmoothingMode.None
                Next

                ' Il righello verticale e' aggiornato
                NeedsVerticalRedraw = False

                Return myVRulerBmp
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Nothing
            Finally
                ' Libero la memoria allocata per l'oggetto Graphics
                If GR IsNot Nothing Then
                    GR.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Crea la penna da usare per le operazioni di drag and drop dei righelli
        ''' e la assegna a myDragPen
        ''' </summary>
        Private Sub CreateDragPen()
            myDragPen = New Pen(myDragLineColor)
            ' Imposto il pattern che andro' ad usare per le linee di drag
            Dim DashPattern() As Single = {20, 7, 1, 7}
            myDragPen.DashPattern = DashPattern
            myDragPen.DashStyle = Drawing2D.DashStyle.Custom
        End Sub

        ''' <summary>
        ''' Controlla se la bitmap del righello orizzontale va ricreata oppure no.
        ''' </summary>
        Private Sub CheckHorizontalRulerBitmap()
            ' Se ho cambiato la dimensione del righello orizzontale, devo distruggere la bitmap esistente
            If (myHRulerBmp IsNot Nothing) AndAlso (myHRulerBmp.Width <> NeededBitmapWidth) Then
                myHRulerBmp.Dispose()
                myHRulerBmp = Nothing
            End If
            ' Se e' necessario, creo la bitmap su cui disegnare
            If (myHRulerBmp Is Nothing) Then
                myHRulerBmp = New Bitmap(NeededBitmapWidth, NeededBitmapRulerSize)
                'MsgBox(String.Format("Rulers: Allocata una bitmap orizzontale da {0}x{1}", myHRulerBmp.Width, myHRulerBmp.Height))
                ' Aggiorno il flag "necessario ridisegno"
                NeedsHorizontalRedraw = True
            End If
        End Sub

        ''' <summary>
        ''' Controlla se la bitmap del righello verticale va ricreata oppure no.
        ''' </summary>
        Private Sub CheckVerticalRulerBitmap()
            ' Se ho cambiato la dimensione del righello verticale, devo distruggere la bitmap esistente
            If (myVRulerBmp IsNot Nothing) AndAlso (myVRulerBmp.Height <> NeededBitmapHeight) Then
                myVRulerBmp.Dispose()
                myVRulerBmp = Nothing
            End If
            ' Se e' necessario, creo la bitmap su cui disegnare
            If (myVRulerBmp Is Nothing) Then
                myVRulerBmp = New Bitmap(NeededBitmapRulerSize, NeededBitmapHeight)
                'MsgBox(String.Format("Rulers: Allocata una bitmap verticale da {0}x{1}", myVRulerBmp.Width, myVRulerBmp.Height))
                ' Aggiorno il flag "necessario ridisegno"
                NeedsVerticalRedraw = True
            End If
        End Sub

#End Region

#Region "Gestione degli eventi della PictureBox"

        ''' <summary>
        ''' Evento generato quando cambia l'unita' di misura della PictureBox associata.
        ''' </summary>
        Private Sub PictureBox_MeasureUnitChanged(ByVal unit As MeasureSystem.enUniMis) Handles myPictureBoxControl.OnMeasureUnitChanged
            ' Quando cambia l'unita' di misura, ho bisogno di un ridisegno di entrambi i righelli
            NeedsHorizontalRedraw = True
            NeedsVerticalRedraw = True
        End Sub

#End Region

#Region "Funzioni pubbliche"

        ''' <summary>
        ''' Disegna i righelli sul Graphics passatogli.
        ''' NOTA: Il graphics passato puo' essere stato scalato oppure no, per questa routine la cosa e' ininfluente.
        ''' </summary>
        Public Sub Draw(ByVal GR As Graphics)
            Try
                ' Check se i righelli hanno dimensioni valide
                If (Width <= 0) OrElse (Height <= 0) Then
                    Exit Sub
                End If

                ' Check se dall'ultimo ridisegno effettuato e' cambiata la View corrente o i dati di ridisegno delle bitmap
                ' NOTA BENE: In questo caso l'operatore "<>" fa un confronto di ConversionInfo, non di GraphicInfo
                If (myLastGraphicInfo Is Nothing) OrElse (myLastGraphicInfo <> PictureBoxControl.GraphicInfo) Then
                    NeedsHorizontalRedraw = True
                    NeedsVerticalRedraw = True
                End If

                ' Controlla se le bitmap su cui disegnare i righelli vanno ricreate o aggiornate.
                ' NOTA: Se le bitmap vengono modificate, imposta NeedsHorizontalRedraw e NeedsVerticalRedraw a true
                CheckHorizontalRulerBitmap()
                CheckVerticalRulerBitmap()

                ' Se serve, ridisegno le bitmap dei righelli
                If NeedsHorizontalRedraw Then
                    RedrawHorizontalRuler()
                End If
                If NeedsVerticalRedraw Then
                    RedrawVerticalRuler()
                End If

                ' Aggiorno i dati relativi all'ultimo ridisegno effettuato
                myLastGraphicInfo = PictureBoxControl.GraphicInfo.Clone()


                ' Check se tutte le bitmap sono ok
                Dim bitmapOk As Boolean = True
                bitmapOk = bitmapOk AndAlso (myHRulerBmp IsNot Nothing) AndAlso (myHRulerBmp.Width > 0) AndAlso (myHRulerBmp.Height > 0)
                bitmapOk = bitmapOk AndAlso (myVRulerBmp IsNot Nothing) AndAlso (myVRulerBmp.Width > 0) AndAlso (myVRulerBmp.Height > 0)
                bitmapOk = bitmapOk AndAlso (myOriginBmp IsNot Nothing) AndAlso (myOriginBmp.Width > 0) AndAlso (myOriginBmp.Height > 0)
                bitmapOk = bitmapOk AndAlso (myOriginBmpSnapped IsNot Nothing) AndAlso (myOriginBmpSnapped.Width > 0) AndAlso (myOriginBmpSnapped.Height > 0)
                If Not bitmapOk Then
                    MsgBox("Found invalid bitmap in " + Me.GetType.FullName() + ".Paint()")
                    Return
                End If

                ' Salvo lo stato precedente e cancello le eventuali trasformazioni
                Dim oldState As GraphicsState = GR.Save()
                GR.ResetTransform()

                Try
                    ' Disegno i righelli
                    GR.DrawImageUnscaled(myHRulerBmp, 0, 0)
                    GR.DrawImageUnscaled(myVRulerBmp, 0, 0)

                    GR.DrawImage(myOriginBmp, 0, 0, myOriginBmp.Width, myOriginBmp.Height)
                Finally
                    ' Ripristino lo stato precedente
                    GR.Restore(oldState)
                End Try
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End Sub

        ''' <summary>
        ''' Disegna la linea di riferimento per il drag and drop dal righello orizzontale
        ''' </summary>
        Public Sub DrawHorizontalDragDropLine(ByVal GR As Graphics, ByVal Y As Integer)
            Try
                Dim yCoord As Single = (Y - LogicalOrigin.Y) * ScaleFactor
                GR.DrawLine(myDragPen, 0, yCoord, Width, yCoord)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End Sub

        ''' <summary>
        ''' Disegna la linea di riferimento per il drag and drop dal righello verticale
        ''' </summary>
        Public Sub DrawVerticalDragDropLine(ByVal GR As Graphics, ByVal X As Integer)
            Try
                Dim xCoord As Single = (X - LogicalOrigin.X) * ScaleFactor
                GR.DrawLine(myDragPen, xCoord, 0, xCoord, Height)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End Sub

        ''' <summary>
        ''' Ritorna il passo del righello orizzontale o di quello verticale
        ''' </summary>
        Public Function GetRulerStep() As Single
            ' Trovo il passo di base con cui scrivero' i numeri nel righello [micron]
            ' Il passo di base e' dato dal passo maggiore necessario per i due righelli (orizzontale e verticale)
            Dim stepValue As Double = Math.Max(CalculateBaseStep(True), CalculateBaseStep(False))

            ' Check se ho un passo di base valido
            If stepValue = 0 Then
                Return 1.0!
            End If

            ' Faccio in modo che tra un numero e l'altro ci sia un po' di spazio libero
            stepValue *= FreeSpaceFactor

            ' Se non una unita' di misura metrica, devo ricordarmi che base numerica uso,
            ' e adeguare il valore del passo alla base numerica in uso
            Dim baseUnit As Double = 1
            If UnitOfMeasure = MeasureSystem.enUniMis.inches Then
                stepValue /= 25400
                baseUnit = 25400
            End If

            ' Riporto il passo ad un multiplo di 1, 2, 5, 10, 20, 50, ecc...
            Dim valuesArray As Double() = {1, 2, 5}

            ' Riporto il passo ad un multiplo dei numeri presenti nell'array
            ' Il numero scelto tra quelli presenti e' quello che mi minimizza l'errore
            Dim actualLog, actualRounded, actualError, bestStep As Double
            Dim minError As Double = Double.MaxValue
            For Each actualValue As Double In valuesArray
                ' Potenza di 10 che mi porta al passo attuale, in funzione del possibile step
                actualLog = Math.Log10(stepValue / actualValue)
                ' Calcolo l'arrotondamento piu' vicino e l'errore associato a questo possibile step
                actualRounded = Math.Round(actualLog)
                actualError = Math.Abs(actualRounded - actualLog)
                ' Scelgo lo step con errore minore
                If actualError < minError Then
                    bestStep = actualValue * baseUnit * Math.Pow(10, actualRounded)
                    minError = actualError
                End If
            Next
            Return bestStep
        End Function

#End Region

    End Class

End Class
