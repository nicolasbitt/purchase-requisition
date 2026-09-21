# Connection with Excel
$xl = New-Object -ComObject Excel.Application

$caminho = "C:\Projetos\ME5A\ME5A_REPORT.xlsm"

$wb = $null

# Colocar aqui para verificar se a conexao com o SAP existe

try {
	
	# Colocar como false quando subir para producao, para rodar em segundo plano
	$xl.Visible = $false
	$xl.DisplayAlerts = $false

	# Abrindo ME5A_REPORT.xlsm
	$wb = $xl.Workbooks.Open($caminho)
	
	# Executando o modulo responsavel por atualizar a planilha
	$xl.Run("'$($wb.Name)'!SAP_ME5A_UPDATE")

	# Executando o modulo responsavel por enviar o e-mail
	$xl.Run("'$($wb.Name)'!Email")

	Add-Type -AssemblyName PresentationFrameWork; [void][System.Windows.MessageBox]::Show("E-mail ME5A dessa semana pronto para envio!", "Sucesso!", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

}
catch {

	Write-Host "Erro:"
	Write-Host $_.Exception.Message

}
finally {
	
	if ($wb -ne $null) {
	$wb.Close($false)
	}

	if ($xl -ne $null) {
	$xl.Quit()
	}
}