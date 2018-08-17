<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Model/clParametros.asp"-->
<!--#include file="../Util/freeASPUpload.asp"-->
<%
tipo = Request("tipo")

if tipo = "salvarParametros" then 
	salvarParametros()
elseif tipo = "BuscarParametros" then 
	BuscarParametros()
end if 

function salvarParametros()
	
	stremail_envio = request("email_envio")
	strservidor_envio = request("servidor_envio")
	strQtdTentativasCampanha = request("qtd_tentativas_envio_campanha")
	strIntervaloThread = request("intervalo_espera_thread")
	strLimiteEnvioCanal = request("limite_envio_canal")
	
	
	set cParametros = new Parametros	
		cParametros.GravarParametros stremail_envio,strservidor_envio,strQtdTentativasCampanha,strIntervaloThread,strLimiteEnvioCanal
	set cParametros = nothing
	
end function

function BuscarParametros()

	set cParametros = new Parametros	
	cParametros.BuscarParametros
	strEmailEnvio = ""
	strServidorEnvio = ""
	if not isnull(cParametros.email_envio)then
		strEmailEnvio = cParametros.email_envio
		strServidorEnvio = cParametros.servidor_envio
	end if 
	
	Set a = jsArray()
	Set a(Null) = jsObject()
	a(Null)("email_envio") = strEmailEnvio
	a(Null)("de_servidor_envio")= strServidorEnvio
	a(Null)("qtd_tentativas_envio_campanha")= cParametros.qtd_tentativas_envio_campanha
	a(Null)("intervalo_espera_thread")= cParametros.intervalo_espera_thread
	a(Null)("qtd_limite_envio_canal")= cParametros.qtd_limite_envio_canal
	
	set cParametros = nothing
	a.flush


end function


%>