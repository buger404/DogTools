VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wallpaperer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Function GetHTML(ByVal url As String) As String

    Dim html As String, XMLObject As Object
    Set XMLObject = CreateObject("Microsoft.XMLHTTP")
    XMLObject.open "GET", url, False
    XMLObject.send
    html = XMLObject.responsetext
    GetHTML = html
    Set XMLObject = Nothing
    
End Function
Private Function GetWebFile(ByVal url As String) As Byte()

    Dim XMLObject As Object
    Set XMLObject = CreateObject("Microsoft.XMLHTTP")
    XMLObject.open "GET", url, False
    XMLObject.send
    GetWebFile = XMLObject.responsebody
    Set XMLObject = Nothing
    
End Function
Public Function GetWallpaperList() As String()
    Randomize
    
    Dim html As String, temp() As String, temp2() As String, papers() As String, formats() As String, file As String, url As String
    ReDim papers(0)
    
    html = GetHTML("http://desk.zol.com.cn/fengjing/" & Int(Rnd * 80) & ".html")
    
    temp = Split(html, "src=""")
    For i = 1 To UBound(temp)
        temp2 = Split(temp(i), """")
        If temp2(0) Like "*://desk-fd.zol-img.com.cn/t_s*c5/*.jpg" Then
            formats = Split(temp2(0), "desk-fd.zol-img.com.cn/t_s")
            formats = Split(formats(1), "c5/")
            url = "http://desk-fd.zol-img.com.cn/t_s" & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY & "c5/" & formats(1)
            ReDim Preserve papers(UBound(papers) + 1)
            papers(UBound(papers)) = url
        End If
    Next
    
    GetWallpaperList = papers
End Function
Public Function GetWallpaper() As String

    Dim papers() As String
    papers = GetWallpaperList
    GetWallpaper = papers(Int(Rnd * (UBound(papers) - 1)) + 1)
    
End Function

Public Sub DownloadWallpaper(ByVal Path As String)

    Dim papers() As String, data() As Byte
    papers = GetWallpaperList
    data = GetWebFile(papers(Int(Rnd * (UBound(papers) - 1)) + 1))
    Open Path For Binary As #1
    Put #1, , data
    Close #1
    
End Sub
