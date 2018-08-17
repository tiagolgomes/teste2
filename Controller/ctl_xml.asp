<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/clEBS.asp"-->
<!--#include file="../Model/clSMS.asp"-->
<!--#include file="../Util/srcUtil.asp"-->
<%
tipo = Request("tipo")

if tipo = "atulizarXML" then 
	atulizarXML()
end if 

function atulizarXML()
	Set Fso = Server.CreateObject ("Scripting.FileSystemObject")
	Set arquivo = Fso.GetFolder("C:\inetpub\wwwroot\SpeedSms\XML/")
	Set arquivos = arquivo.files
 	For Each a in arquivos
		caminho = server.MapPath("../XML/"&a.name) 
		strArrayNome = Split(a.name,".")
		
		Set objXML = Server.CreateObject("Microsoft.XMLDOM")
		objXML.async = False 
		objXML.load(caminho)
		
		if CStr(objXML.parseError.errorCode) <> "0" Then 
			'Se NAO conseguiu ler o arquivo (ele NAO FOI CRIADO...)
			response.end
		Else
			Set Root = objXML.documentElement
			strRetornoXML =  Root.childNodes(0).selectSingleNode("generic").attributes(0).text
			strMensagemRetorno =  Root.childNodes(0).selectSingleNode("generic").attributes(1).text
			
			if(trim(strRetornoXML) = "Error") then 
				strRetornoXML = 400
				if(trim(strMensagemRetorno) <> "No free channel found") then 
					strMensagemEBS =  Root.childNodes(0).selectSingleNode("generic").attributes(2).text
				end if 
			elseif (trim(strRetornoXML) = "Success") then 
				strRetornoXML = 200
				strMensagemEBS =  Root.childNodes(0).selectSingleNode("generic").attributes(3).text
			end if 	
			set cSMS = new SMS
				if(trim(strMensagemRetorno) <> "No free channel found") then 
					cSMS.AtualizarCampanhaEnvio strRetornoXML,strMensagemRetorno,strArrayNome(0),strMensagemEBS
				end if 
			Set cSMS = nothing	
		
			If Err.Number = 0 Then
				Set FSO = Server.CreateObject("Scripting.FileSystemObject")
					FSO.DeleteFile caminho
				Set FSO = nothing
			End If
		end if
		Set xmlDoc = nothing
	Next
	Set arquivos = nothing
	Set arquivo = nothing
	Set Fso = nothing

end function
%>