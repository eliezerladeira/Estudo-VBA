Attribute VB_Name = "mdlCalend�rio"
Option Explicit


Dim R�tulos() As New cCalend�rio

Function GetCalend�rio() As Date
        
          
    Dim lTotalR�tulos As Long
    Dim ctrl As control
    Dim FRM As frmCalend�rio
    
    Set FRM = New frmCalend�rio
    
       For Each ctrl In FRM.Controls
        If ctrl.Name Like "l?c?" Then
            lTotalR�tulos = lTotalR�tulos + 1
            ReDim Preserve R�tulos(1 To lTotalR�tulos)
            Set R�tulos(lTotalR�tulos).lblGrupo = ctrl
        End If
    Next ctrl

    FRM.Show
    
        If IsDate(FRM.Tag) Then
        GetCalend�rio = FRM.Tag
    Else
        GetCalend�rio = Date
    End If
        
    Unload FRM



End Function
