Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

Public Class Bitmap32

	Public ImageBytes As Byte()
	Public RowSizeBytes As Integer
	Public Const PixelDataSize As Integer = 32

	Public Bitmap As Bitmap

	Private m_IsLocked As Boolean = False


	Public ReadOnly Property IsLocked As Boolean

		Get
			Return m_IsLocked
		End Get
	End Property

	Public Sub New(ByVal bm As Bitmap)
		Bitmap = bm
	End Sub

	Private m_BitmapData As BitmapData

	Public ReadOnly Property Width As Integer
		Get

			Return Bitmap.Width

		End Get

	End Property


	Public ReadOnly Property Height As Integer

		Get
			Return Bitmap.Height
		End Get
	End Property

	Public Sub GetPixel(ByVal x As Integer, ByVal y As Integer, <Out> ByRef red As Byte, <Out> ByRef green As Byte, <Out> ByRef blue As Byte, <Out> ByRef alpha As Byte)

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		blue = ImageBytes(Math.Min(System.Threading.Interlocked.Increment(i), i - 1))
		green = ImageBytes(Math.Min(System.Threading.Interlocked.Increment(i), i - 1))
		red = ImageBytes(Math.Min(System.Threading.Interlocked.Increment(i), i - 1))
		alpha = ImageBytes(i)

	End Sub


	Public Sub SetPixel(ByVal x As Integer, ByVal y As Integer, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte, ByVal alpha As Byte)

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		ImageBytes(Math.Min(System.Threading.Interlocked.Increment(i), i - 1)) = blue
		ImageBytes(Math.Min(System.Threading.Interlocked.Increment(i), i - 1)) = green
		ImageBytes(Math.Min(System.Threading.Interlocked.Increment(i), i - 1)) = red
		ImageBytes(i) = alpha
	End Sub

	Public Function GetBlue(ByVal x As Integer, ByVal y As Integer) As Byte

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		Return ImageBytes(i)

	End Function


	Public Sub SetBlue(ByVal x As Integer, ByVal y As Integer, ByVal blue As Byte)

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		ImageBytes(i) = blue
	End Sub

	Public Function GetGreen(ByVal x As Integer, ByVal y As Integer) As Byte

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		Return ImageBytes(i + 1)

	End Function


	Public Sub SetGreen(ByVal x As Integer, ByVal y As Integer, ByVal green As Byte)

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		ImageBytes(i + 1) = green
	End Sub

	Public Function GetRed(ByVal x As Integer, ByVal y As Integer) As Byte

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		Return ImageBytes(i + 2)

	End Function


	Public Sub SetRed(ByVal x As Integer, ByVal y As Integer, ByVal red As Byte)

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		ImageBytes(i + 2) = red
	End Sub

	Public Function GetAlpha(ByVal x As Integer, ByVal y As Integer) As Byte

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		Return ImageBytes(i + 3)

	End Function


	Public Sub SetAlpha(ByVal x As Integer, ByVal y As Integer, ByVal alpha As Byte)

		Dim i As Integer = y * m_BitmapData.Stride + x * 4

		ImageBytes(i + 3) = alpha
	End Sub

	Public Sub LockBitmap()

		If IsLocked Then Return

		Dim bounds As Rectangle = New Rectangle(0, 0, Bitmap.Width, Bitmap.Height)

		m_BitmapData = Bitmap.LockBits(bounds, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
		RowSizeBytes = m_BitmapData.Stride

		Dim total_size As Integer = m_BitmapData.Stride * m_BitmapData.Height

		ImageBytes = New Byte(total_size - 1) {}
		Marshal.Copy(m_BitmapData.Scan0, ImageBytes, 0, total_size)
		m_IsLocked = True
	End Sub

	Public Sub UnlockBitmap()

		If Not IsLocked Then Return
		Dim total_size As Integer = m_BitmapData.Stride * m_BitmapData.Height

		Marshal.Copy(ImageBytes, 0, m_BitmapData.Scan0, total_size)

		Bitmap.UnlockBits(m_BitmapData)
		ImageBytes = Nothing

		m_BitmapData = Nothing
		m_IsLocked = False

	End Sub


	Public Sub Average()

		Dim was_locked As Boolean = IsLocked

		LockBitmap()


		For y As Integer = 0 To Height - 1

			For x As Integer = 0 To Width - 1
				Dim red, green, blue, alpha As Byte

				GetPixel(x, y, red, green, blue, alpha)

				Dim gray As Byte = CByte(((red + green + blue) / 3))

				SetPixel(x, y, gray, gray, gray, alpha)

			Next
		Next


		If Not was_locked Then UnlockBitmap()

	End Sub


	Public Sub Grayscale()

		Dim was_locked As Boolean = IsLocked

		LockBitmap()


		For y As Integer = 0 To Height - 1

			For x As Integer = 0 To Width - 1
				Dim red, green, blue, alpha As Byte

				GetPixel(x, y, red, green, blue, alpha)

				Dim gray As Byte = CByte((0.3 * red + 0.5 * green + 0.2 * blue))

				SetPixel(x, y, gray, gray, gray, alpha)

			Next
		Next


		If Not was_locked Then UnlockBitmap()

	End Sub


	Public Sub ClearRed()

		Dim was_locked As Boolean = IsLocked

		LockBitmap()


		For y As Integer = 0 To Height - 1

			For x As Integer = 0 To Width - 1
				SetRed(x, y, 0)

			Next
		Next


		If Not was_locked Then UnlockBitmap()

	End Sub


	Public Sub ClearGreen()

		Dim was_locked As Boolean = IsLocked

		LockBitmap()


		For y As Integer = 0 To Height - 1

			For x As Integer = 0 To Width - 1
				SetGreen(x, y, 0)

			Next
		Next


		If Not was_locked Then UnlockBitmap()

	End Sub


	Public Sub ClearBlue()

		Dim was_locked As Boolean = IsLocked

		LockBitmap()


		For y As Integer = 0 To Height - 1

			For x As Integer = 0 To Width - 1
				SetBlue(x, y, 0)

			Next
		Next


		If Not was_locked Then UnlockBitmap()

	End Sub


	Public Sub Invert()

		Dim was_locked As Boolean = IsLocked

		LockBitmap()


		For y As Integer = 0 To Height - 1

			For x As Integer = 0 To Width - 1
				Dim red As Byte = CByte((255 - GetRed(x, y)))

				Dim green As Byte = CByte((255 - GetGreen(x, y)))

				Dim blue As Byte = CByte((255 - GetBlue(x, y)))

				Dim alpha As Byte = GetAlpha(x, y)

				SetPixel(x, y, red, green, blue, alpha)

			Next
		Next


		If Not was_locked Then UnlockBitmap()

	End Sub


	Public Class Filter
		Public Kernel As Single(,)

		Public Weight, Offset As Single


		Public Sub Normalize()

			Weight = 0

			For row As Integer = 0 To Kernel.GetUpperBound(0)


				For col As Integer = 0 To Kernel.GetUpperBound(1)

					Weight += Kernel(row, col)

				Next
			Next

		End Sub


		Public Sub ZeroKernel()

			Dim total As Single = 0


			For row As Integer = 0 To Kernel.GetUpperBound(0)


				For col As Integer = 0 To Kernel.GetUpperBound(1)

					total += Kernel(row, col)

				Next
			Next


			Dim row_mid As Integer = CInt((Kernel.GetUpperBound(0) / 2))

			Dim col_mid As Integer = CInt((Kernel.GetUpperBound(1) / 2))

			total -= Kernel(row_mid, col_mid)

			Kernel(row_mid, col_mid) = -total
		End Sub
	End Class

	Public Function Clone() As Bitmap32

		Dim was_locked As Boolean = Me.IsLocked

		Me.LockBitmap()
		Dim result As Bitmap32 = CType(Me.MemberwiseClone(), Bitmap32)

		result.Bitmap = New Bitmap(Me.Bitmap.Width, Me.Bitmap.Height)

		result.m_IsLocked = False
		If Not was_locked Then Me.UnlockBitmap()

		Return result

	End Function

	Public Shared ReadOnly Property EmbossingFilter As Filter
		Get
			Return New Filter() With {
			.Weight = 1,
			.Offset = 127,
			.Kernel = New Single(,) {
			{-1, 0, 0},
			{0, 0, 0},
			{0, 0, 1}}
		}
		End Get
	End Property

	Public Shared ReadOnly Property EmbossingFilter2 As Filter
		Get
			Return New Filter() With {
			.Weight = 1,
			.Offset = 127,
			.Kernel = New Single(,) {
			{2, 0, 0},
			{0, -1, 0},
			{0, 0, -1}}
		}
		End Get
	End Property
	Public Shared ReadOnly Property BlurFilter5x5Gaussian As Filter
		Get
			Dim result As Filter = New Filter() With {
				.Offset = 0,
				.Kernel = New Single(,) {
				{1, 4, 7, 4, 1},
				{4, 16, 26, 16, 4},
				{7, 26, 41, 26, 7},
				{4, 16, 26, 16, 4},
				{1, 4, 7, 4, 1}}
			}
			result.Normalize()
			Return result
		End Get
	End Property

	Public Shared ReadOnly Property BlurFilter5x5Mean As Filter
		Get
			Dim result As Filter = New Filter() With {
				.Offset = 0,
				.Kernel = New Single(,) {
				{1, 1, 1, 1, 1},
				{1, 1, 1, 1, 1},
				{1, 1, 1, 1, 1},
				{1, 1, 1, 1, 1},
				{1, 1, 1, 1, 1}}
			}
			result.Normalize()
			Return result
		End Get
	End Property

	Public Shared ReadOnly Property EdgeDetectionFilterULtoLR As Filter
		Get
			Return New Filter() With {
				.Weight = 1,
				.Offset = 0,
				.Kernel = New Single(,) {
				{-5, 0, 0},
				{0, 0, 0},
				{0, 0, 5}}
			}
		End Get
	End Property

	Public Shared ReadOnly Property EdgeDetectionFilterTopToBottom As Filter
		Get
			Return New Filter() With {
				.Weight = 1,
				.Offset = 0,
				.Kernel = New Single(,) {
				{-1, -1, -1},
				{0, 0, 0},
				{1, 1, 1}}
			}
		End Get
	End Property

	Public Shared ReadOnly Property EdgeDetectionFilterLeftToRight As Filter
		Get
			Return New Filter() With {
				.Weight = 1,
				.Offset = 0,
				.Kernel = New Single(,) {
				{-1, 0, 1},
				{-1, 0, 1},
				{-1, 0, 1}}
			}
		End Get
	End Property

	Public Shared ReadOnly Property HighPassFilter3x3 As Filter
		Get
			Return New Filter() With {
				.Weight = 16,
				.Offset = 127,
				.Kernel = New Single(,) {
				{-1, -2, -1},
				{-2, 12, -2},
				{-1, -2, -1}}
			}
		End Get
	End Property

	Public Shared ReadOnly Property HighPassFilter5x5 As Filter
		Get
			Dim result As Filter = New Filter() With {
				.Offset = 127,
				.Kernel = New Single(,) {
				{-1, -4, -7, -4, -1},
				{-4, -16, -26, -16, -4},
				{-7, -26, -41, -26, -7},
				{-4, -16, -26, -16, -4},
				{-1, -4, -7, -4, -1}}
			}
			result.Normalize()
			result.Weight = -result.Weight
			result.ZeroKernel()
			Return result
		End Get
	End Property

	Public Function ApplyFilter(ByVal filter As Filter, ByVal lock_result As Boolean) As Bitmap32
		Dim result As Bitmap32 = Me.Clone()
		Dim was_locked As Boolean = Me.IsLocked
		Me.LockBitmap()
		result.LockBitmap()
		Dim xoffset As Integer = -CInt((filter.Kernel.GetUpperBound(1) / 2))
		Dim yoffset As Integer = -CInt((filter.Kernel.GetUpperBound(0) / 2))
		Dim xmin As Integer = -xoffset
		Dim xmax As Integer = Bitmap.Width - filter.Kernel.GetUpperBound(1)
		Dim ymin As Integer = -yoffset
		Dim ymax As Integer = Bitmap.Height - filter.Kernel.GetUpperBound(0)
		Dim row_max As Integer = filter.Kernel.GetUpperBound(0)
		Dim col_max As Integer = filter.Kernel.GetUpperBound(1)

		For x As Integer = xmin To xmax

			For y As Integer = ymin To ymax
				Dim skip_pixel As Boolean = False
				Dim red As Single = 0, green As Single = 0, blue As Single = 0

				For row As Integer = 0 To row_max

					For col As Integer = 0 To col_max
						Dim ix As Integer = x + col + xoffset
						Dim iy As Integer = y + row + yoffset
						Dim new_red, new_green, new_blue, new_alpha As Byte
						Me.GetPixel(ix, iy, new_red, new_green, new_blue, new_alpha)

						If new_alpha = 0 Then
							skip_pixel = True
							Exit For
						End If

						red += new_red * filter.Kernel(row, col)
						green += new_green * filter.Kernel(row, col)
						blue += new_blue * filter.Kernel(row, col)
					Next

					If skip_pixel Then Exit For
				Next

				If Not skip_pixel Then
					red = filter.Offset + red / filter.Weight
					If red < 0 Then red = 0
					If red > 255 Then red = 255
					green = filter.Offset + green / filter.Weight
					If green < 0 Then green = 0
					If green > 255 Then green = 255
					blue = filter.Offset + blue / filter.Weight
					If blue < 0 Then blue = 0
					If blue > 255 Then blue = 255
					result.SetPixel(x, y, CByte(red), CByte(green), CByte(blue), Me.GetAlpha(x, y))
				End If
			Next
		Next

		If Not lock_result Then result.UnlockBitmap()
		If Not was_locked Then Me.UnlockBitmap()
		Return result
	End Function

	Public Sub Pixellate(ByVal rank As Integer, ByVal lock_result As Boolean)
		Dim was_locked As Boolean = Me.IsLocked
		Me.LockBitmap()
		Dim y As Integer = 0

		While y < Height
			Dim x As Integer = 0

			While x < Width
				Dim total_r As Integer = 0
				Dim total_g As Integer = 0
				Dim total_b As Integer = 0
				Dim num_pixels As Integer = 0

				For row As Integer = y To y + rank - 1

					If row < Height Then

						For col As Integer = x To x + rank - 1

							If col < Width Then
								total_r += GetRed(col, row)
								total_g += GetGreen(col, row)
								total_b += GetBlue(col, row)
								num_pixels += 1
							End If
						Next
					End If
				Next

				Dim byte_r As Byte = CByte((total_r / num_pixels))
				Dim byte_g As Byte = CByte((total_g / num_pixels))
				Dim byte_b As Byte = CByte((total_b / num_pixels))

				For row As Integer = y To y + rank - 1

					If row < Height Then

						For col As Integer = x To x + rank - 1

							If col < Width Then
								SetPixel(col, row, byte_r, byte_g, byte_b, 255)
							End If
						Next
					End If
				Next

				x += rank
			End While

			y += rank
		End While

		If Not was_locked Then Me.UnlockBitmap()
	End Sub

	Public Function Pointellate(ByVal rank As Integer, ByVal point_diameter As Integer, ByVal lock_result As Boolean) As Bitmap
		Dim was_locked As Boolean = Me.IsLocked
		Me.LockBitmap()
		Dim bm As Bitmap = New Bitmap(Width, Height)

		Using gr As Graphics = Graphics.FromImage(bm)
			gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
			Dim y As Integer = 0

			While y < Height
				Dim x As Integer = 0

				While x < Width
					Dim total_r As Integer = 0
					Dim total_g As Integer = 0
					Dim total_b As Integer = 0
					Dim num_pixels As Integer = 0

					For row As Integer = y To y + rank - 1

						If row < Height Then

							For col As Integer = x To x + rank - 1

								If col < Width Then
									total_r += GetRed(col, row)
									total_g += GetGreen(col, row)
									total_b += GetBlue(col, row)
									num_pixels += 1
								End If
							Next
						End If
					Next

					Dim byte_r As Byte = CByte((total_r / num_pixels))
					Dim byte_g As Byte = CByte((total_g / num_pixels))
					Dim byte_b As Byte = CByte((total_b / num_pixels))
					Dim offset As Integer = (rank - point_diameter) / 2

					Using br As Brush = New SolidBrush(Color.FromArgb(255, byte_r, byte_g, byte_b))
						gr.FillEllipse(br, x + offset, y + offset, point_diameter, point_diameter)
					End Using

					x += rank
				End While

				y += rank
			End While
		End Using

		If Not was_locked Then Me.UnlockBitmap()
		Return bm
	End Function
End Class
