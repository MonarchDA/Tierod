Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
''' 

Public Class clsimgcapture
#Region "Variables"
    Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
    Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Integer) As Integer
    Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
    Public Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
    Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Public Declare Function BitBlt Lib "GDI32" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Integer) As Integer
    Public Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Integer) As Integer
    Public Declare Function DeleteObject Lib "GDI32" (ByVal hObj As Integer) As Integer
    Const SRCCOPY As Integer = &HCC0020
    Public Background As Bitmap
    Public fw, fh As Integer
    Public Shared img_arry(0) As Object
    Public intI As Integer
    Public blnTab As Boolean
    Public Shared tab_imgaarry(50) As Object
    Public Shared node As TreeNode
    Public Shared pic As PictureBox
#End Region

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Public Sub CaptureScreen()

        Try
            Dim hsdc, hmdc As Integer
            Dim hbmp, hbmpold As Integer
            Dim r As Integer
            hsdc = CreateDC("DISPLAY", "", "", "")
            hmdc = CreateCompatibleDC(hsdc)
            fw = GetDeviceCaps(hsdc, 8)
            fh = GetDeviceCaps(hsdc, 10)
            hbmp = CreateCompatibleBitmap(hsdc, fw, fh)
            hbmpold = SelectObject(hmdc, hbmp)
            r = BitBlt(hmdc, 0, 0, fw, fh, hsdc, 0, 0, 13369376)
            hbmp = SelectObject(hmdc, hbmpold)
            r = DeleteDC(hsdc)
            r = DeleteDC(hmdc)
            Background = Image.FromHbitmap(New IntPtr(hbmp))
            DeleteObject(hbmp)
        Catch ex As Exception
            '  getErrorLog(ex)
        End Try

    End Sub

    Public Sub StoreImages()
        Try
            CaptureScreen()
            pic = New PictureBox
            pic.Image = Background
            img_arry(intI) = pic.Image
        Catch ex As Exception
        End Try
    End Sub

    Private Function getTreeViewInfo_Tabs(ByVal strScreenName As String) As Boolean
        Dim k As Integer = 0
    End Function
End Class
