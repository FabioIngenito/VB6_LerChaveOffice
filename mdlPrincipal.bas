Attribute VB_Name = "mdlPrincipal"
Option Explicit
Public objWinReg           As WinReg32
Public objBuscaChaves      As clsBuscaChaves

Sub main()

    If objWinReg Is Nothing Then Set objWinReg = New WinReg32

    If objBuscaChaves Is Nothing Then Set objBuscaChaves = New clsBuscaChaves
    
    frmLerChaveOffice.Show
End Sub
