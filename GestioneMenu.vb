Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports System.Reflection
Imports System.Resources
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing.Imaging

Public Class GestioneMenu : Inherits MenuItem
    ' ESEMPIO DI COME RICHIAMARE IL TUTTO
    'Dim myMenu As New ContextMenu()
    'Dim mnuCarico As New GestioneMenu

    ' METTERE IL TRATTINO PER AVERE UN SEPARATORE
    'mnuCarico = New GestioneMenu("Verdana", 12, "Carica", "Immagini\Carica.ico", 16, New EventHandler(AddressOf mnuCarico_Click), Nothing)

    'myMenu.MenuItems.Add(mnuCarico)

    'lstLog.ContextMenu = myMenu
    ' ESEMPIO DI COME RICHIAMARE IL TUTTO

    ' PER ABILITARE / DISABILITARE
    ' mnuCarico.Abilita() 
    ' PER ABILITARE / DISABILITARE

    Private m_Icon As Bitmap
    Private m_Font As Font
    Private larghezzaIcona As Integer
    Private AltezzaIcona As Integer
    Private Abilitato As Boolean
    Private fOnClick As EventHandler
    Private Testo As String

    Private m_Gradient_Color1 As Color = SystemColors.Highlight
    Private m_Gradient_Color2 As Color = SystemColors.Control

    Public Sub New()
        ' MyClass.New("", 1, "", 16, Nothing, Nothing, System.Windows.Forms.Shortcut.None)
    End Sub

    Public Sub New(ByVal FontName As String, ByVal FontSize As Integer, ByVal text As String, ByVal icon As String, ByVal iconSize As Integer, ByVal onClickP As EventHandler, ByVal shortcut As Shortcut)
        ' MyBase.New(text, onClickP, shortcut)

        fOnClick = onClickP
        OwnerDraw = True
        m_Font = New Font(FontName, FontSize)
        If File.Exists(icon) Then
            m_Icon = CaricaImmagineSenzaLockarla(icon, iconSize)
        Else
            If iconSize = 0 Then iconSize = 16
            Dim bmp As New Bitmap(iconSize, iconSize, PixelFormat.Format24bppRgb)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.FillRectangle(Brushes.Transparent, 0, 0, iconSize, iconSize)
                g.FillEllipse(Brushes.Red, 0, 0, iconSize, iconSize)
            End Using
            bmp.MakeTransparent()
            m_Icon = bmp
        End If
        Abilitato = True
        Testo = text

        If Not m_Icon Is Nothing Then
            AltezzaIcona = m_Icon.Width
            larghezzaIcona = m_Icon.Height
        End If
    End Sub

    Private Function RidimensionaImmagine(imgOriginale As Bitmap, iconSize As Integer) As Bitmap
        'If imgOriginale Is Nothing Then
        '    If iconSize = 0 Then iconSize = 16
        '    Dim bmp As New Bitmap(iconSize, iconSize, PixelFormat.Format24bppRgb)
        '    Using g As Graphics = Graphics.FromImage(bmp)
        '        g.FillRectangle(Brushes.Transparent, 0, 0, iconSize, iconSize)
        '        g.FillEllipse(Brushes.Red, 0, 0, iconSize, iconSize)
        '    End Using
        '    bmp.MakeTransparent()
        '    Return bmp
        'Else
        Return New Bitmap(imgOriginale, Val(iconSize), Val(iconSize))
        'End If
    End Function

    Private Function CaricaImmagineSenzaLockarla(NomeImmagine As String, iconSize As Integer) As Image
        Dim bmp As Image = Nothing
        Dim fs As System.IO.FileStream = Nothing

        Try
            fs = New System.IO.FileStream(NomeImmagine, IO.FileMode.Open, IO.FileAccess.Read)
            bmp = System.Drawing.Image.FromStream(fs)
        Catch ex As Exception
            'Stop
        End Try

        If fs Is Nothing = False Then
            Try
                fs.Close()
                fs.Dispose()
            Catch ex As Exception

            End Try
        End If
        fs = Nothing

        bmp = RidimensionaImmagine(bmp, iconSize)

        Return bmp
    End Function

    Private Sub Me_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        If Abilitato Then
            If Not fOnClick Is Nothing Then
                fOnClick.Invoke(Me, EventArgs.Empty)
            End If
        End If
    End Sub

    Public Sub ImpostaImmagine(icon As String, iconsize As Integer)
        m_Icon = CaricaImmagineSenzaLockarla(icon, iconsize)
    End Sub

    Public Function LeggeTesto() As String
        Return Testo
    End Function

    Public Sub ImpostaTesto(sTesto As String)
        Testo = sTesto
    End Sub

    Public Sub Abilita()
        Abilitato = True
    End Sub

    Public Sub Disabilita()
        Abilitato = False
    End Sub

    Private Function GetRealText() As String
        Dim s As String = Testo

        If ShowShortcut And Shortcut <> Shortcut.None Then
            Dim k As Keys = CType(Shortcut, Keys)
            s = s & Convert.ToChar(9) & TypeDescriptor.GetConverter(GetType(Keys)).ConvertToString(k)
        End If

        Return s
    End Function

    Private Sub DisegnaTutto(e As DrawItemEventArgs)
        Dim br As Brush
        Dim rcBk As Rectangle = e.Bounds
        rcBk.X += larghezzaIcona + 10

        If Testo <> "-" Then
            If Not m_Icon Is Nothing Then
                e.Graphics.DrawImage(m_Icon, e.Bounds.Left + 2, e.Bounds.Top + 2)
                ' e.Graphics.DrawIcon(m_Icon, e.Bounds.Left + 2, e.Bounds.Top + 2)
            End If

            If Abilitato Then
                If CBool(e.State And DrawItemState.Selected) Then
                    br = New LinearGradientBrush(rcBk, m_Gradient_Color1, m_Gradient_Color2, 0)
                Else
                    br = SystemBrushes.Control
                End If
            Else
                br = SystemBrushes.Control
            End If

            e.Graphics.FillRectangle(br, rcBk)

            Dim sf As StringFormat = New StringFormat()
            sf.HotkeyPrefix = HotkeyPrefix.Show

            If Abilitato Then
                br = New SolidBrush(e.ForeColor)
            Else
                br = New SolidBrush(Color.DarkGray)
            End If

            e.Graphics.DrawString(GetRealText(), m_Font, br, e.Bounds.Left + larghezzaIcona + 15, e.Bounds.Top + 2, sf)
        Else
            br = SystemBrushes.Control
            e.Graphics.FillRectangle(br, rcBk)
        End If
    End Sub

    Protected Overrides Sub OnDrawItem(ByVal e As DrawItemEventArgs)
        MyBase.OnDrawItem(e)

        DisegnaTutto(e)
    End Sub

    Protected Overrides Sub OnMeasureItem(ByVal e As MeasureItemEventArgs)
        Dim sf As New StringFormat()
        sf.HotkeyPrefix = HotkeyPrefix.Show
        MyBase.OnMeasureItem(e)
        e.ItemHeight = AltezzaIcona + 2
        e.ItemWidth = CInt(e.Graphics.MeasureString(GetRealText(), m_Font, 10000, sf).Width) + 30
    End Sub
End Class
