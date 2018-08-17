<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/clLogin.asp"-->
<%


tipo = Request("tipo")



if tipo = "ValidaUsuario" then 
	ValidaUsuario()
elseif tipo = "salvarUsuario" then 
	salvarUsuario()	
elseif tipo = "carregrUltimosAcessos" then
	carregrUltimosAcessos()
elseif tipo = "zerarSenha" then 
	zerarSenha()	
elseif tipo = "validarSessao" then 
	validarSessao()
elseif tipo ="buscarUsuario" then 
	buscarUsuario()
elseif tipo = "sairSistema" then  
	sairSistema()	
end if 

function ValidaUsuario()
	
	
	strUsuario = request("usuario")
	strSenha = request("senha")
	strCdEmpresa = request("cd_empresa")
	Session("id_usuario")  = ""
	set cLogin = new Login	
	Set a = jsArray()
	cLogin.ValidarUsuario strUsuario,md5(strSenha),strCdEmpresa
	Set a(Null) = jsObject()
	a(Null)("id_usuario") 	= cLogin.id_usuario
	if cLogin.id_usuario <> "" then 
		Session("id_usuario") = cLogin.id_usuario
		Session("cd_empresa") = cLogin.cod_empresa
		Session("usuario") = cLogin.usuario
	end if 	
	
	set cLogin = nothing
	a.flush	
end function

function salvarUsuario()
	
	strUsuario = request("usuario")
	strSenha = request("senha")
	
	set cLogin = new Login	
	Set a = jsArray()
	cLogin.salvarUsuaro strUsuario,md5(strSenha)
	
	set cLogin = nothing
end function 

function carregrUltimosAcessos()

	set cLogin = new Login
	Set a = jsArray()
	
	for each item in cLogin.carregrUltimosAcessos()
		Set a(Null) = jsObject()
		a(Null)("id_usuario") 	= item.id_usuario
		a(Null)("usuario") 	= item.usuario
		a(Null)("ultimo_acesso") = item.ultimo_acesso
	next
	
	set cLogin = nothing
	a.flush
end function 

function zerarSenha()
	strIdUsuario = request("id_usuario")
	set cLogin = new Login	
	cLogin.zerarSenha strIdUsuario
	set cLogin = nothing
end function

function validarSessao()
	Set a = jsArray()
	Set a(Null) = jsObject()
		a(Null)("cd_empresa") = Session("cd_empresa")
	a.flush
end function 

function buscarUsuario()
	Set a = jsArray()
	Set a(Null) = jsObject()
		a(Null)("usuario") 	= Session("usuario")
		a(Null)("id_usuario") 	= Session("id_usuario")
	a.flush
end function

function sairSistema()
	Session.Abandon
end function

%>