Private Sub CommandButton1_Click()
' D�k�manda yazd�rma alan�ndaki k�s�mlar� masa �st�ne aktif sayfa ismi ile pdf kaydeder
On Error Resume Next
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
"C:\Users\Deniz\Desktop\" & ActiveSheet.Name

End Sub
