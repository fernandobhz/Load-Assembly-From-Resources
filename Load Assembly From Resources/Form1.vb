Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging
Imports System.Windows.Forms

Public Class Form1

    Private Cam As WebCamCapture
    Private m_ip As IntPtr = IntPtr.Zero
    Private PicBox As New PictureBox

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        CloseDx()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Cam = New WebCamCapture(0, 640, 480, 24, Me.PictureBox1)
    End Sub

    Private Sub CloseDx()
        Try
            If Cam IsNot Nothing Then
                Cam.Dispose()
            End If

            RealeaseBuffer()
        Catch ex As Exception
        End Try
    End Sub

    Sub RealeaseBuffer()
        Try
            If m_ip <> IntPtr.Zero Then
                Marshal.FreeCoTaskMem(m_ip)
                m_ip = IntPtr.Zero
            End If
        Catch ex As Exception
        End Try
    End Sub

    Function TakeCamShot() As Bitmap
        RealeaseBuffer()

        m_ip = Cam.Click()
        Dim b As New Bitmap(Cam.Width, Cam.Height, Cam.Stride, PixelFormat.Format24bppRgb, m_ip)
        b.RotateFlip(RotateFlipType.RotateNoneFlipY)

        Return b

    End Function
End Class