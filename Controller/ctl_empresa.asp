<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Model/clEmpresa.asp"-->
<%


tipo = Request("tipo")

if tipo = "salvarEmpresa" then 
	salvarEmpresa()	
elseif tipo = "carregarEmpresas" then
	carregarEmpresas()	
elseif tipo = "carregarEmpresaPorId" then
	carregarEmpresaPorId()
elseif tipo="buscarTotalSMSEnviados" then 
	buscarTotalSMSEnviados()
end if 


function salvarEmpresa()
	strEmpresa = request("cd_empresa")
	strDeEmpresa= request("de_empresa")
	strQtEnvioMensal = request("qt_envio_mensal")
	
	set cEmpresa = new Empresa
	Set a = jsArray()
	cEmpresa.salvarEmpresa strEmpresa, strDeEmpresa, strQtEnvioMensal
	
	set cLogin = nothing
end function 


function carregarEmpresas()

	set cEmpresa = new Empresa
	Set a = jsArray()
	
	for each item in cEmpresa.carregarEmpresas()
		Set a(Null) = jsObject()
		a(Null)("cd_empresa") 	= item.cd_empresa
		a(Null)("de_empresa") 	= item.de_empresa
		a(Null)("qt_envio_mensal") = item.qt_envio_mensal
	next
	
	set cEmpresa = nothing
	a.flush
end function 

function carregarEmpresaPorId()
	
	strEmpresa = request("id_empresa")

	set cEmpresa = new Empresa
	Set a = jsArray()
	
	cEmpresa.carregarEmpresaPorId strEmpresa
	
	Set a(Null) = jsObject()
	a(Null)("cd_empresa") 	= cEmpresa.cd_empresa
	a(Null)("de_empresa") 	= cEmpresa.de_empresa
	a(Null)("qt_envio_mensal") = cEmpresa.qt_envio_mensal
	
	
	set cEmpresa = nothing
	a.flush
end function 

function buscarTotalSMSEnviados()
	
	set cEmpresa = new Empresa
	Set a = jsArray()
	
	cEmpresa.buscarTotalSMSEnviados()
		Set a(Null) = jsObject()
		a(Null)("qtd_total_sms_enviados_mes") 	= cEmpresa.qtd_total_sms_enviados_mes
	
	set cEmpresa = nothing
	a.flush
	
end function


%>