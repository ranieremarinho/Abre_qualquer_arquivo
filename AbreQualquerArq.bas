Attribute VB_Name = "AbreQualquerArq"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1

Public Sub AbreArq(FormhWnd As Long, ByVal Arq As String, ByVal CaminhoArq As String, Optional ModoExib As Byte)

    If ModoExib = Empty Then ModoExib = 1
    
    ExecutarInstru��o FormhWnd, "open", Arq, "", CaminhoArq, ModoExib
    '2 abre minimizado
    '0 carrega mas n�o abre (s� vemos no Ctrl+Alt+Del)
    '1 e 5 abre normal com foco
    '3 e 6 e 7 abre maximizado
    '4 abre normal sem foco

End Sub

Public Sub AbreArquiv(FormhWnd As Long, ByVal CaminhoCompleto As String, Optional ModoExib As Byte)

    If ModoExib = Empty Then ModoExib = 1
    
    ExecutarInstru��o FormhWnd, "open", GetNomeArq(CaminhoCompleto), "", GetSohCaminho(CaminhoCompleto), ModoExib
    '2 abre minimizado
    '0 carrega mas n�o abre (s� vemos no Ctrl+Alt+Del)
    '1 e 5 abre normal com foco
    '3 e 6 e 7 abre maximizado
    '4 abre normal sem foco

End Sub

Private Function ExecutarInstru��o(hWnd As Long, Instru��o As String, Arquivo As String, Par�metros As String, Caminho As String, ModoDeExibi��o As Byte)

    ExecutarInstru��o = ShellExecute(hWnd, Instru��o, Arquivo, Par�metros, Caminho, ModoDeExibi��o)
    
End Function

Public Function GetNomeArq(ByVal CaminhoComp As String) As String

    ' "c:\teste\cfg.dll" -> "cfg.dll"
    
    Dim I As Integer
    Dim NomeArq As String
    
    NomeArq = Mid$(CaminhoComp, InStrRev(CaminhoComp, "\") + 1)
    
    GetNomeArq = NomeArq
    
End Function

Public Function GetSohCaminho(ByVal CaminhoCompleto As String) As String
    
    CaminhoCompleto = Trim(CaminhoCompleto)
    
    If Right(CaminhoCompleto, 1) = "\" Then GoTo Conclui
    If InStr(CaminhoCompleto, "\") = 0 Then GoTo Conclui

    CaminhoCompleto = Mid$(CaminhoCompleto, 1, InStrRev(CaminhoCompleto, "\"))
    
Conclui:
    
    GetSohCaminho = CaminhoCompleto
    
End Function
