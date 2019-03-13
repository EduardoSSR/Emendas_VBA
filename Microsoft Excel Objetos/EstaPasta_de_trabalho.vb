Private Sub Workbook_BeforeClose(Cancel As Boolean)
	lsDesligarTelaCheia
End Sub

Private Sub Workbook_Open()

	Call set_combo1
	lsLigarTelaCheia
	Call verifica_validade
	Sheets("Principal").Select
End Sub
