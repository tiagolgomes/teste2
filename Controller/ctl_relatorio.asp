<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/clRelatorio.asp"-->
<!--#include file="../Util/srcUtil.asp"-->
<%
tipo = Request("tipo")

if tipo = "gerarRelatorioCampanha" then
	gerarRelatorioCampanha()
elseif tipo = "buscarNumerosEnviados" then
	buscarNumerosEnviados()
elseif tipo = "gerarRelatorioRetorno" then
	gerarRelatorioRetorno()
elseif tipo = "gerarRelatorioNumerosInutilizados" then
	gerarRelatorioNumerosInutilizados()
elseif tipo = "gerarRelatorioCampanhaXEnvio" then
	gerarRelatorioCampanhaXEnvio()
elseif tipo = "gerarRelatorioRetornosCampanha" then
	gerarRelatorioRetornosCampanha()
elseif tipo = "gerarRelatorioCampanhasCSV" then
	gerarRelatorioCampanhasCSV()
elseif tipo = "gerarRelatorioSMSEnviados" then
	gerarRelatorioSMSEnviados()
elseif tipo= "buscarConsumoMensal" then
	buscarConsumoMensal()
end if

function gerarRelatorioCampanha()

	strDataInicial = request("data_inicial")
	strDataFinal = request("data_final")
	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.relBuscaCampanhaXEnvio(strDataInicial,strDataFinal)
		Set a(Null) = jsObject()
		a(Null)("id_campanha") 	= item.id_campanha
		a(Null)("email_retorno") 	= item.email_retorno
		a(Null)("total_campanha") 	= item.total_campanha
		a(Null)("data_envio") 	= item.data_envio
		a(Null)("texto") 	= item.texto
		a(Null)("qtde_msg_enviadas") 	= item.qtde_msg_enviadas
		a(Null)("qtde_msg_erros") 	= item.qtde_msg_erros
		a(Null)("qtde_msg_estouro") 	= item.qtde_msg_estouro
	next
	a.flush
	set cRelatorio = nothing
end function


function gerarRelatorioCampanhaXEnvio()

	strDataInicial = request("data_inicial")
	strDataFinal = request("data_final")
	set cRelatorio = new Relatorio
	Set a = jsArray()

	hora = hour(Now)
	minuto = minute(Now)
	segundo =  second(Now)
	arquivo_excel = "CampanaXEnvio_"&replace(date,"/","")&"_"&hora&"_"&minuto&"_"&segundo&".csv"

	Set a(Null) = jsObject()
		a(Null)("arquivo_excel") = arquivo_excel

	set fso = createobject("scripting.filesystemobject")
	Set act = fso.CreateTextFile(server.mappath("../Excel/"&arquivo_excel), true)

	'// CRIA O HERADER DAS COLUNAS
	act.WriteLine("nu_contrato;telefone;texto;")

	for each item in cRelatorio.gerarRelatorioCampanhaXEnvioContrato(strDataInicial,strDataFinal)
			act.WriteLine(chr(34)&item.nu_contrato&chr(34)&";"&_
						  chr(34)&item.telefone&chr(34)&";"&_
						  chr(34)&item.texto&chr(34))
	next

	a.flush
	set cRelatorio = nothing
end function

function buscarNumerosEnviados()
	strCampanha = request("id_campanha")

	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.buscarNumerosEnviados(strCampanha)
		Set a(Null) = jsObject()
		a(Null)("id_campanha") 	= item.id_campanha
		a(Null)("numero") 	= item.telefone
	next
	a.flush
	set cRelatorio = nothing
end function



function gerarRelatorioRetorno()

	strDataInicial = request("data_inicial")
	strDataFinal = request("data_final")
	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.relBuscaRetornos(strDataInicial,strDataFinal)
		Set a(Null) = jsObject()
		a(Null)("telefone") = item.telefone
		a(Null)("mensagem") = item.mensagem
		a(Null)("nu_contrato") = ""
		a(Null)("texto") = item.texto
		if item.nu_contrato <> "null" and item.nu_contrato <> "" then
			a(Null)("nu_contrato") = item.nu_contrato
		end if
	next
	a.flush
	set cRelatorio = nothing
end function

function gerarRelatorioRetornosCampanha()
	strDataInicial = request("data_inicial")
	strDataFinal = request("data_final")
	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.relBuscaRetornosCampanha(strDataInicial,strDataFinal)
		Set a(Null) = jsObject()
		a(Null)("numero") = item.telefone
		a(Null)("resposta_cliente") = item.mensagem
		a(Null)("mensagem_enviada") = item.texto
		a(Null)("data_resposta") = item.data_resposta
	next
	a.flush
	set cRelatorio = nothing
end function


function gerarRelatorioNumerosInutilizados()

	strDataInicial = request("data_inicial")
	strDataFinal = request("data_final")
	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.relBuscaNumerosInutilizados(strDataInicial,strDataFinal)
		Set a(Null) = jsObject()
		a(Null)("telefone") = item.telefone
		a(Null)("id_campanha") = item.id_campanha
		a(Null)("data_envio") = item.data_envio
	next
	a.flush
	set cRelatorio = nothing
end function



function gerarRelatorioCampanhasCSV()

	strIdCampanha = trim(request("id_campanha"))

	set cRelatorio = new Relatorio
	arquivo_excel = "RelatorioNumeros_"&strIdCampanha&".csv"

	Set a = jsArray()
	Set a(Null) = jsObject()
	a(Null)("arquivo_excel") = arquivo_excel

	set fso = createobject("scripting.filesystemobject")
	Set act = fso.CreateTextFile(server.mappath("../Excel/"&arquivo_excel), true)

	act.WriteLine("telefone;contrato;texto;status_sms")

	for each item in cRelatorio.gerarRelatorioCampanhasCSV(strIdCampanha)
		act.WriteLine(chr(34)&item.telefone&chr(34)&";"&_
					  chr(34)&item.nu_contrato&chr(34)&";"&_
					  chr(34)&item.texto&chr(34)&";"&_
					  chr(34)&item.status_envio&chr(34))

	next
	act.close


	a.flush
	set cRelatorio = nothing
	set fso = nothing
end function

function gerarRelatorioSMSEnviados()
	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.gerarRelatorioSMSEnviados()
		Set a(Null) = jsObject()
		a(Null)("data_envio") = item.data_envio &" / "& MonthName(Cint(item.mes_envio),True)
		a(Null)("qtde_msg_enviadas") = item.qtde_msg_enviadas

	next
	a.flush
	set cRelatorio = nothing
end function

function buscarConsumoMensal()

	set cRelatorio = new Relatorio
	Set a = jsArray()

	for each item in cRelatorio.buscarConsumoMensal()
		Set a(Null) = jsObject()
		a(Null)("data_envio") = item.data_envio
		a(Null)("qtde_msg_enviadas") = item.qtde_msg_enviadas

	next
	a.flush
	set cRelatorio = nothing




end function 


%>