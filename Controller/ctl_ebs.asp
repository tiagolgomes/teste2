<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/ClEbs.asp"-->
<%
tipo = Request("tipo")

if tipo = "autenticarUsuario" then 
	autenticarUsuario()
elseif tipo = "buscarEBS" then 
	buscarEBS()
elseif tipo ="buscarCanais"	then 
	buscarCanais()	
elseif tipo = "mudarOpcaoStatus" then 
	mudarOpcaoStatus()	
elseif tipo = "zerarEnvioCanal" then 
	zerarEnvioCanal()	
end if 


function autenticarUsuario()
	' --- AUTENTICAÇÃO -----	
	  
	  Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
	  oXmlHTTP.Open "GET", "http://200.150.102.98:8088/sms/manager?action=login&username=hdr&secret=hdr17rdh", false 
	  oXmlHTTP.send()
	  set oXmlHTTP = nothing
	  
	' ---- FIM DA AUTENTICAÇÃO ----	
end function


function buscarEBS()
	
	set cEbs = new Ebs	
	Set a = jsArray()
	
	for each item in cEbs.buscarEBS()
		Set a(Null) = jsObject()
		a(Null)("nome_ebs") = item.nome_ebs
	next
	set cEbs = nothing
	a.flush

end function

function buscarCanais()

	strNomeEbs =  trim(Request("nome_ebs"))
	
	set cEbs = new Ebs	
	Set a = jsArray()
	
	for each item in cEbs.buscarCanais(strNomeEbs)
		
		Set a(Null) = jsObject()
		a(Null)("id_ebs") = item.id_ebs
		a(Null)("nome_ebs") = item.nome_ebs
		a(Null)("canal_ebs") = item.canal_ebs
		a(Null)("qtd_envio") = item.qtd_envio
		a(Null)("qtd_tentativas") = item.qtd_tentativas
		
		if(item.flg_ativo = "N") then 
			a(Null)("flg_ativo") = false
		else
			a(Null)("flg_ativo") = "checked"
		end if 
		a(Null)("flg_em_uso") = item.flg_em_uso
	next
	set cEbs = nothing
	a.flush
end function

function mudarOpcaoStatus()
	strHabilitaOpcao = trim(request("habilita_opcao"))
	strIDCanal = trim(request("id_canal"))
	set cEbs = new Ebs	
		cEbs.mudarOpcaoStatus strIDCanal,strHabilitaOpcao
	set cEbs = nothing
end function  


function zerarEnvioCanal()
	strIDCanal = trim(request("id_canal"))
	set cEbs = new Ebs	
		cEbs.zerarEnvioCanal strIDCanal
	set cEbs = nothing
end function  



 
%>