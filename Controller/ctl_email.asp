<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Model/clEmail.asp"-->
<!--#include file="../Model/clCampanha.asp"-->
<!--#include file="../Model/clParametros.asp"-->
<%
tipo = Request("tipo")

if tipo = "EnviarEmail" then 
	EnviarEmail()
elseif tipo = "EnviarEmailTeste" then 
	EnviarEmailTeste()
elseif tipo= "EnviarEmailRetorno" then 
	EnviarEmailRetorno()
elseif tipo= "EnviarEmailRetornoTeste" then 
	EnviarEmailRetornoTeste()		
end if 

function EnviarEmail()

	strIdCampanha = request("id_campanha")
	set cCampanha = new Campanha	
	cCampanha.BuscaCampanhaId strIdCampanha
		strEmailRetorno = cCampanha.email_retorno
		strTexto = cCampanha.texto
		strFlLimiteSms = cCampanha.fl_limite_sms
	set cCampanha = nothing
	
	
	Set a = jsArray()
	Set a(Null) = jsObject()
	if(strFlLimiteSms = "-1") then 
		a(Null)("retornoCampanha") = false
	else 
		set cParametros = new Parametros	
			cParametros.BuscarParametros
			strEmailEnvio = cParametros.email_envio
		set cParametros = nothing
	
		strCorpoTeste = strTexto
		strAssunto = "Campanha "&strIdCampanha
		strEmailEnvio = strEmailEnvio
		strEmailEnvioTo = strEmailRetorno
		set cEmail = new Email	
		cEmail.EnviarEmail strCorpoTeste,strAssunto,strEmailEnvio,strEmailEnvioTo
		set cEmail = nothing
		a(Null)("retornoCampanha") = true
	end if 

	a.flush
end function 

function EnviarEmailTeste()
	strCorpoTeste =  "Email teste OK"
	strEmailEnvio = request("email_envio")
	strEmailEnvioTo = request("email_teste")
	set cEmail = new Email	
	cEmail.EnviarEmail strCorpoTeste,"Email Teste",strEmailEnvio,strEmailEnvioTo
	set cEmail = nothing
end function

function EnviarEmailRetorno()

	set cParametros = new Parametros	
		cParametros.BuscarParametros
		strEmailEnvio = cParametros.email_envio
	set cParametros = nothing
	
	strAssunto = "Retorno Campanha"
	set cEmail = new Email	
	for each item in cEmail.BuscaCampanhaEnvioRetorno()
		strCorpo = "Texto Campanha <br/>"&item.email_texo&"<br/><br/>Resposta número: "&item.numero_retorno&"<br/><br/>"& item.mensagem_retorno
		if item.nu_contrato <> "null" and item.nu_contrato <> "" then 
			strCorpo = strCorpo & "<br/> Número do Contrato: "& item.nu_contrato 
		end if 
		
		strEmailEnvioTo = item.email_retorno
		cEmail.EnviarEmail strCorpo,strAssunto,strEmailEnvio,strEmailEnvioTo
		cEmail.AtualizarEnvio item.numero_retorno
	next
	set cEmail = nothing
	

end function

function EnviarEmailRetornoTeste()
	set cParametros = new Parametros	
		cParametros.BuscarParametros
		strEmailEnvio = cParametros.email_envio
	set cParametros = nothing
	
	strAssunto = "Retorno Campanha"
	set cEmail = new Email	
	for each item in cEmail.BuscaCampanhaEnvioRetornoTeste()
		strEmailEnvioTo = item.email_retorno
		response.write strEmailEnvioTo&"<br/>"
	next
	set cEmail = nothing
end function




%>