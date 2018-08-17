<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Model/clBalckList.asp"-->
<!--#include file="../Util/freeASPUpload.asp"-->
<%
tipo = Request("tipo")

if tipo = "gravarBlackList" then 
	gravarBlackList()
elseif tipo = "buscarTelefone" then 
	buscarTelefone()		
end if 

function gravarBlackList()
	
	strTelefone = trim(request("telefone"))
	strIdBlackList = trim(request("id_blackList"))
	strFlgAtivo = trim(request("flg_ativo"))
	
	
	set cBlackList = new BlackList	
		cBlackList.GravarBalckList strTelefone,strIdBlackList,strFlgAtivo
	set cBlackList = nothing
	
	
end function


function buscarTelefone()

	strTelefone = request("telefone")
	
	set cBlackList = new BlackList	
	Set a = jsArray()
	
	for each item in cBlackList.buscarTelefone(strTelefone)
		Set a(Null) = jsObject()
		a(Null)("id_blackList") = item.id_blackList
		a(Null)("telefone") = item.telefone
		a(Null)("data_alteracao") = item.data_alteracao
		a(Null)("flg_ativo") = item.flg_ativo
		a(Null)("nome_ultimo_usuario_alterou") = item.nome_ultimo_usuario_alterou
		a(Null)("checked") = "no_checked"
		if(trim(item.flg_ativo) = "S") then 
			a(Null)("checked") = "checked"
		end if 
	next
	set cBlackList = nothing
	a.flush
	
end function


%>