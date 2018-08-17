<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/clCampanha.asp"-->
<!--#include file="../Util/freeASPUpload.asp"-->
<%
tipo = Request("tipo")

if tipo = "salvarCampanha" then 
	salvarCampanha()
elseif tipo = "excel_campanha_envio" then 
	excel_campanha_envio()	
elseif tipo = "enviarCMD" then 	
	enviarCMD()		
elseif tipo = "UparArquivo" then 
	UparArquivo()	
elseif tipo="buscarCampanhasPadrao" then 
	buscarCampanhasPadrao()
elseif tipo="excluirPadrao" then 
	excluirPadrao()	
elseif tipo= "ValidarQuantidadesEnvio" then 
	ValidarQuantidadesEnvio()	
elseif tipo= "buscarNumero" then 
	buscarNumero()	
elseif tipo="inutilizarNumero" then 
	inutilizarNumero()	
elseif tipo="gravarContratos" then
	gravarContratos()
elseif tipo="excluirCampanha"	then 
	excluirCampanha()
elseif tipo="IntegrarSMS" then
	IntegrarSMS()
elseif tipo= "codeURL" then 
	codeURL()	
end if 

function salvarCampanha()

	strIdMeioComunicacao = "1"
	strEmailRetorno = request("email_retorno")
	strTexto = replace(request("texto_campanha"),"'","''")
	strNomeArquivo = request("nome_arquivo")
	strObservacaoPadrao = request("observacao_padrao")
	flRetorno = request("fl_retorno")
	strDataAgendamento = request("data_agendamento")
	strHoraAgendamento = request("hora_agendamento")
	
	set cCampanha = new Campanha
	 
	 if not strObservacaoPadrao  then 
		cCampanha.GravarCampanhasPadrao replace(strTexto,"'","''")
	 end if 
	 cCampanha.GravarCampanha strIdMeioComunicacao,strEmailRetorno,strTexto,flRetorno,strDataAgendamento,strHoraAgendamento
	 strIdCampanha = cCampanha.id_campanha
	 Set a = jsArray()
	 Set a(Null) = jsObject()
	 a(Null)("id_campanha") = strIdCampanha
	 Session("id_campanha") = strIdCampanha
	 Session("nome_arquivo") = strNomeArquivo
	
	set cCampanha = nothing
	a.flush
end function

function excel_campanha_envio()
	
	stridCampanha = request("id_campanha")
	set cCampanha = new Campanha	
	
	arquivo_excel= "campanha_envio"&stridCampanha&".csv"	
	set fso = createobject("scripting.filesystemobject")
	Set act = fso.CreateTextFile(server.mappath("../Excel/"&arquivo_excel), true)
	for each item in cCampanha.BuscaCampanhaEnvio(stridCampanha)
		if item.telefone <> "" then 
			
			strTG3 = item.tg_3
			if(IsDate(strTG3)) then strTG3 = ""&strTG3&"" 
				
			act.WriteLine(item.telefone&","&chr(34)&replace(replace(replace(item.texto,"#tg1#",item.tg_1),"#tg2#",item.tg_2),"#tg3#",strTG3)&chr(34))
		end if
	next
	act.close
	set cCampanha = nothing
	
end function

function UparArquivo()
	Set Upload = New FreeASPUpload
	Count = upload.Save(server.mapPath("..\Excel"))

	response.redirect "../view/frmcarregando.asp"
	
end function

function enviarCMD()
	stridCampanha = request("id_campanha")
	response.redirect "../cmd_php/executacmd.php?id_campanha="&stridCampanha
end function 

function buscarCampanhasPadrao()
	
	set cCampanha = new Campanha	
	Set a = jsArray()
	
	for each item in cCampanha.buscarCampanhasPadrao()
		Set a(Null) = jsObject()
		a(Null)("id_padrao") = item.id_padrao
		a(Null)("texto") = item.texto
	next
	
	a.flush
	set cCampanha = nothing
end function 

function excluirPadrao()
	set cCampanha = new Campanha	
	Set a = jsArray()
	strPadrao = split(request("id_padrao"),"_")
	cCampanha.excluirPadrao(strPadrao(1))
	set cCampanha = nothing
end function 

function ValidarQuantidadesEnvio()
	set cCampanha = new Campanha	
	Set a = jsArray()
	
	cCampanha.BuscaCampanhaQuantidadeDeEnviosMensalEmpresa
	cCampanha.BuscaCampanhaQuantidadeDeEnviosMensal
	strQuantidadeEnvioMensal = Session("numeros_envio_mensal")
	strQuantidadePermitidaEmpresa = Session("numeros_envio_mensal_empresa")
	Set a(Null) = jsObject()
	a(Null)("retorno") = "false"
	if (CLng(strQuantidadeEnvioMensal) >= CLng(strQuantidadePermitidaEmpresa)) then 
		a(Null)("retorno") = "true"
	end if 
	a.flush
	set cCampanha = nothing
end function  

function buscarNumero()
	strNumero = request("telefone")
	
	set cCampanha = new Campanha	
	Set a = jsArray()
	
	for each item in cCampanha.buscarNumero(strNumero)
		Set a(Null) = jsObject()
		a(Null)("id_campanha") = item.id_campanha
		a(Null)("telefone") = item.telefone
		a(Null)("data_envio") = item.data_envio
		a(Null)("nome_arquivo") = "check"
		if(item.fl_numero_inutilizado = "-1") then  a(Null)("nome_arquivo") = "no_check"
	next
	
	a.flush
	set cCampanha = nothing
end function 

function inutilizarNumero()
	set cCampanha = new Campanha	
	Set a = jsArray()
	strIdCampanha = request("id_campanha")
	strTelefone = request("telefone")
	strNomeArquivo = request("nome_arquivo")
	
	cCampanha.inutilizarNumero strIdCampanha,strTelefone,strNomeArquivo
	set cCampanha = nothing
end function

function gravarContratos()
	set cCampanha = new Campanha	
	Set a = jsArray()
	strIdCampanha = request("id_campanha")
	
	cCampanha.gravarContratos strIdCampanha
	set cCampanha = nothing
end function

function excluirCampanha()
	set cCampanha = new Campanha	
	Set a = jsArray()
	strIdCampanha = request("id_campanha")
	
	cCampanha.excluirCampanha strIdCampanha
	set cCampanha = nothing
end function 

function IntegrarSMS()
	 set cCampanha = new Campanha	
	 for each item in cCampanha.IntegrarSMS()
		strIdCampanha = item.id_campanha
		strEmailRetorno = item.email_retorno
		if(trim(strEmailRetorno) <> "") then 
			montarArquivoIntegracao(strIdCampanha)
		end if 
 	 next
	 set cCampanha = nothing
end function

function montarArquivoIntegracao(strIdCampanha)
	arquivo_excel= "campanha_envio"&strIdCampanha&".csv"	
	set fso = createobject("scripting.filesystemobject")
	Set act = fso.CreateTextFile(server.mappath("../Excel/"&arquivo_excel), true)
	
	set cCampanha = new Campanha	
	for each item in cCampanha.BuscaCampanhaEnvio(stridCampanha)
		if item.telefone <> "" then 
			
			strTG3 = item.tg_3
			if(IsDate(strTG3)) then strTG3 = ""&strTG3&"" 
				
			act.WriteLine(item.telefone&","&chr(34)&replace(replace(replace(item.texto,"#tg1#",item.tg_1),"#tg2#",item.tg_2),"#tg3#",strTG3)&chr(34))
		end if
	next
	act.close
	set cCampanha = nothing
	
end function

function codeURL()
		
		strUrl = trim(Request("url"))
		
		DataToSend = "{'longUrl':'"&strUrl&"'}"
		
		Set oXmlMensage = Server.CreateObject("Microsoft.XMLHTTP")
		oXmlMensage.Open "GET", "https://www.googleapis.com/urlshortener/v1/url?key=AIzaSyArYIiq8qwKFFxhkL76t1XwujgFEh2hhOw", false 
		oXmlMensage.setRequestHeader "Content-type","application/json"		
		oXmlMensage.send DataToSend 
		
		Set a = jsArray()
		Set a(Null) = jsObject()
		a(Null)("codeURL") = oXmlMensage.ResponseText
		a.flush
end function 
%>