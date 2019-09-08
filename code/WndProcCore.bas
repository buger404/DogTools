Attribute VB_Name = "WndProcCore"
Public WndProc As Long, SafeLock As Boolean

Public Function FunWndProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_DROPFILES Then
        Dim hDrop As Long, nLoopCtr As Integer, IReturn As Long, sFileName As String
        Dim FileList As String
        hDrop = wParam: sFileName = Space$(255)
        nDropCount = DragQueryFileA(hDrop, -1, sFileName, 254)
        For nLoopCtr = 0 To nDropCount - 1
            sFileName = Space$(255)
            IReturn = DragQueryFileA(hDrop, nLoopCtr, sFileName, 254)
            FileList = FileList & Left$(sFileName, IReturn) & vbCrLf
        Next
        Call DragFinish(hDrop)
        If mNowShow = "RegPage" Then RegPage.FileDrop FileList
        If mNowShow = "BatchPage" Then BatchPage.FileDrop FileList
    End If
    
    FunWndProc = CallWindowProcA(WndProc, Hwnd, uMsg, wParam, lParam)
End Function
