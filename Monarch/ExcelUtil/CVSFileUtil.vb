Imports System.IO

Public Class CVSFileUtil

    Private Shared _message As String

    Public Shared ReadOnly Property Message() As String
        Get
            Return _message
        End Get
    End Property

    Public Shared Function Write(ByVal strData As String, ByVal strFileName As String) As Boolean
        Try
            _message = String.Empty
            If File.Exists(strFileName) Then
                File.Delete(strFileName)
            End If

            Dim sw As New StreamWriter(File.Create(strFileName))

            sw.AutoFlush = True
            sw.WriteLine(strData)
            sw.Close()
            Return True
        Catch ex As Exception
            _message = ex.Message
            Return False
        End Try
    End Function

    Public Shared Function WriteCNC(ByVal strData As String, ByVal strFileName As String) As Boolean
        Try
            _message = String.Empty
            If File.Exists(strFileName) Then
                File.Delete(strFileName)
            End If

            Dim sw As New StreamWriter(File.Create(strFileName))

            sw.AutoFlush = True
            sw.Write(strData)
            sw.Close()
            Return True
        Catch ex As Exception
            _message = ex.Message
            Return False
        End Try
    End Function

    Public Shared Function isFolder(ByVal strFolderPath As String) As Boolean
        Try
            _message = String.Empty
            If Directory.Exists(strFolderPath) Then
                Return True
            End If
            Return False
        Catch ex As Exception
            _message = "Failed to test the folder path"
            Return False
        End Try
    End Function

    Public Shared Function Check(ByVal strSourcePath As String, ByVal strDestinationPath As String) As Boolean
        Try
            _message = String.Empty
            If Not File.Exists(strSourcePath) Then
                Return False
            End If
            If Not File.Exists(strDestinationPath) Then
                Return False
            End If

            Dim Source As New StreamReader(File.OpenRead(strSourcePath))
            Dim Destination As New StreamReader(File.OpenRead(strDestinationPath))

            While (Not Source.EndOfStream)
                Dim strSource As String = Source.ReadLine
                Dim strDestination As String = Destination.ReadLine
                If Not strSource.Equals(strDestination) Then
                    MessageBox.Show(strSource & " - " & strDestination)
                    Return False
                End If
            End While

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    'anup 02-03-2011 start
    Public Shared Function DoesFileContainDoubleQuotation(ByVal strSourcePath As String) As Boolean
        Try
            Dim strAreFilesClosed As Boolean = False    '03_03_2011   RAGAVA
            If Not File.Exists(strSourcePath) Then
                Return False
            End If

            Dim DoubleQuotes As String = Char.ConvertFromUtf32(34)
            Dim strMainData As String = String.Empty
            Dim blnAreFilesClosed As Boolean = False

            If strSourcePath.IndexOf(Char.ConvertFromUtf32(34), 0) > -1 Or strSourcePath.IndexOf(",", 0) Then
                Dim Source As New StreamReader(File.Open(strSourcePath, FileMode.Open))

                While (Not Source.EndOfStream)
                    Dim strSource As String = Source.ReadToEnd
                    If strSource.Contains(DoubleQuotes) Then
                        strSource = strSource.Replace(DoubleQuotes, String.Empty)
                        strMainData = strSource
                        Source.Close()
                        Dim Destin As New StreamWriter(File.Open(strSourcePath, FileMode.Create))
                        Destin.AutoFlush = True
                        Destin.WriteLine(strMainData)
                        Destin.Close()
                        strAreFilesClosed = True
                        Exit While
                    End If
                End While
                If strAreFilesClosed = False Then
                    Source.Close()
                End If
            End If

        Catch ex As Exception

        End Try
    End Function

    'anup 28-02-2011 till here


End Class
