<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/clMensagem.asp"-->
<!--#include file="../Util/srcUtil.asp"-->
<%
tipo = Request("tipo")

if tipo = "salvarMensagem" then 
	salvarMensagem()
end if 

function salvarMensagem()
	
	Set a = jsArray()
	
	strTelefone = request("telefone")
	strMensagem = request("mensagem")
	strNuContrato = request("nu_contrato")
	
	if trim(strNuContrato) <> "" then 
		set cMensagem = new Mensagem
			cMensagem.GravarContratos strTelefone,strNuContrato
		set cMensagem = nothing
	end if 
	Set a(Null) = jsObject()
	a(Null)("retorno_classe") = "Msg não enviada! Tente novamente"
    set cMensagem = new Mensagem
		cMensagem.GravarMensagem strTelefone,removeAcentos(strMensagem),strNuContrato
		a(Null)("retorno_classe") = cMensagem.retorno_classe
	set cMensagem = nothing
	a.flush	
end function

%>