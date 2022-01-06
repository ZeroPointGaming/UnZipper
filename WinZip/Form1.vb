Imports System.IO
Imports System.IO.Compression


Public Class Form1
    Public WorkingDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\WinZip Files"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not My.Computer.FileSystem.DirectoryExists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\WinZip Files") Then
            My.Computer.FileSystem.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\WinZip Files")
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

        For Each file In files
            If file.Contains(".zip") Then
                Dim RandomIdentifier As String = (Rnd() * 1 * Rnd()).ToString

                ZipFile.ExtractToDirectory(file, WorkingDirectory + "/" + RandomIdentifier)

                MessageBox.Show(file.Replace(".zip", "").ToString + " Extracted to " + WorkingDirectory.ToString + "/" + RandomIdentifier.ToString)
            Else
                MessageBox.Show("Error, Not a zip file")
            End If
        Next
    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
End Class
