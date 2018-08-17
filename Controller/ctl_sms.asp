<!--#include file="../Util/JSON_2.0.2.asp"-->
<!--#include file="../Util/md5.asp"-->
<!--#include file="../Model/clEBS.asp"-->
<!--#include file="../Model/clSMS.asp"-->
<!--#include file="../Util/srcUtil.asp"-->
<%
tipo = Request("tipo")

if tipo = "autenticarUsuario" then 
	autenticarUsuario()
elseif tipo = "enviarSMS" then 
	enviarSMS()
elseif tipo="salvarCampanha" then 
	salvarCampanha()
elseif tipo = "buscarNumerosComErro" then 
	buscarNumerosComErro()	
elseif tipo = "enviarSMSTESTE" then 
	enviarSMSTESTE()
elseif tipo = "salvarCampanhaEnvio" then 
	salvarCampanhaEnvio()
elseif tipo = "enviarCMDPHP" then 
	enviarCMDPHP()	
elseif tipo = "buscarQuantidadeEBS" then 
	buscarQuantidadeEBS()
elseif tipo = "alterarCanalDeSaida" then	
	alterarCanalDeSaida()	
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
	
	set cSMS = new SMS
	 
	 if not strObservacaoPadrao  then 
		cSMS.GravarCampanhasPadrao replace(strTexto,"'","''")
	 end if 
	 cSMS.GravarCampanha strIdMeioComunicacao,strEmailRetorno,strTexto,flRetorno,strDataAgendamento,strHoraAgendamento
	 strIdCampanha = cSMS.id_campanha
	 Set a = jsArray()
	 Set a(Null) = jsObject()
	 a(Null)("id_campanha") = strIdCampanha
	 Session("id_campanha") = strIdCampanha
	 Session("nome_arquivo") = strNomeArquivo
	
	set cSMS = nothing
	a.flush
end function

function autenticarUsuario()
	' --- AUTENTICAÇÃO -----	
	  
	  Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
	  oXmlHTTP.Open "GET", "http://200.150.102.98:8088/sms/manager?action=login&username=hdr&secret=hdr17rdh", false 
	  oXmlHTTP.send()
	  set oXmlHTTP = nothing
	  
	' ---- FIM DA AUTENTICAÇÃO ----	
end function



function enviarSMS()
	strDados = Split(trim(request("strDados")),";")
	strIdCampanha = trim(request("id_campanha"))
	strTextoCampanha = 	trim(request("texto_campanha"))
	
	Set a = jsArray()
	Set a(Null) = jsObject()
	
	strTotalArquivo = Cint(UBound(strDados))
	
	if(strTotalArquivo >= 0 )then 
		strTelefone = strDados(0)
	end if 
	if(strTotalArquivo >= 1 )then 
		strTG1 = strDados(1)
	end if 
	if(strTotalArquivo >= 2 )then 
		strTG2 = strDados(2)
	end if 
	if(strTotalArquivo >= 3 )then 
		strTG3 = strDados(3)
	end if 
	if(strTotalArquivo >= 4 )then 
		strNuContrato = strDados(4)
	end if 
	
	strTexto  = replace(replace(replace(strTextoCampanha,"#tg1#",strTG1),"#tg2#",strTG2),"#tg3#",strTG3)
	
	set cEBS = new Ebs
	cEBS.BuscarEbsCanal 
		strEBS =  cEBS.nome_ebs
		strCanal = cEBS.canal_ebs
		strIdEbs = cEBS.id_ebs
	
	'// SÓ CHAMA A FUNÇÃO CASO EXISTA EBS DISPONIVEL	
		
		if(trim(strIdEbs) <> "" and not isnull(strIdEbs)) then
			'cEBS.gravarUtilizacaoEbs strIdEbs
			'Set oXmlMensage = Server.CreateObject("Microsoft.XMLHTTP")
			'oXmlMensage.Open "GET", "http://200.150.102.98:8088/sms/mxml?action=ksendsms&device="&strEBS&"C"&strCanal&"&destination="&strTelefone&"&message="&strTexto&"", false 	
			'oXmlMensage.send()
			
			'hora = hour(Now) 
			'minuto = minute(Now)
			'segundo =  second(Now)
			
			'arquivo_xml = replace(date,"/","")&"_"&hora&"_"&minuto&"_"&segundo&"_"&strCanal&".xml"	
			
			'Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
			'xmlDoc.async = False
			'success = xmlDoc.loadXML(oXmlMensage.ResponseXML.xml)
			'If success = True Then    
				 'CASO SEJA NECESSÀRIO SALVAR O XML
				 'xmlDoc.save(server.MapPath("../XML/"&arquivo_xml))
				 
				 'Set Root = xmlDoc.documentElement
				 'strRetornoXML =  Root.childNodes(0).selectSingleNode("generic").attributes(0).text
				 
				 'if(trim(strRetornoXML) = "Error") then 
					'strRetornoXML = 200
				 'elseif (trim(strRetornoXML) = "Sucess") then 
					'strRetornoXML = 400
				 'end if 	
				 
				 set cSMS = new SMS
					 cSMS.GravarCampanhasEnvio strIdCampanha,strTelefone,strTG1,strTG2,strTG3,strNuContrato,strRetornoXML,strEBS&"--"&"C"&strCanal
				 set cSMS = nothing
				 
			'end if 
			
			a(Null)("retorno_canal") = true
			cEBS.gravarLiberacaoEBS strIdEbs
		else
			strRetornoXML = 200
			strSemEbsSaida = "00-00"
			set cSMS = new SMS
				 cSMS.GravarCampanhasEnvio strIdCampanha,strTelefone,strTG1,strTG2,strTG3,strNuContrato,strRetornoXML,strSemEbsSaida
			set cSMS = nothing
			a(Null)("retorno_canal") = false
		end if 
	
	Set xmlDoc = Nothing		
	Set oXmlMensage = Nothing
	set cEBS = nothing
	a.flush
	
end function 


function buscarNumerosComErro()

	strIdCampanha = trim(request("id_campanha"))
	strTextoCampanha = 	trim(request("texto_campanha"))
	
	
	set cSMS = new SMS	
	Set a = jsArray()
	
	for each item in cSMS.buscarNumerosComErro(strIdCampanha)
		if item.telefone <> "" then 
			Set a(Null) = jsObject()
			a(Null)("telefone") = item.telefone
			a(Null)("tg_1") = item.tg_1
			a(Null)("tg_2") = item.tg_2
			a(Null)("tg_3") = item.tg_3
			a(Null)("nu_contrato") = item.nu_contrato
		
		end if
	next
	set cSMS = nothing
	a.flush
	
	
end function 


function salvarCampanhaEnvio()
	strDados = Split(trim(request("strDados")),";")
	strIdCampanha = trim(request("id_campanha"))
	strTextoCampanha = 	trim(request("texto_campanha"))
	strUrlCode = 	trim(request("urlCode"))
	
	Set a = jsArray()
	Set a(Null) = jsObject()
	
	strTotalArquivo = Cint(UBound(strDados))
	
	if(strTotalArquivo >= 0 )then 
		strTelefone = strDados(0)
	end if 
	if(strTotalArquivo >= 1 )then 
		strTG1 = strDados(1)
	end if 
	if(strTotalArquivo >= 2 )then 
		strTG2 = strDados(2)
	end if 
	if(strTotalArquivo >= 3 )then 
		strTG3 = strDados(3)
	end if 
	if(strTotalArquivo >= 4 )then 
		if(IsNumeric(strDados(4))) then
			strNuContrato = strDados(4)
		end if 
	end if 
	
	strTexto  = replace(replace(replace(replace(strTextoCampanha,"#tg1#",strTG1),"#tg2#",strTG2),"#tg3#",strTG3),"#url#",strUrlCode)
	
	strEBS = ""
	strRetornoXML = ""
	set cSMS = new SMS
	  cSMS.GravarCampanhasEnvio strIdCampanha,strTelefone,strTG1,strTG2,strTG3,strNuContrato,strRetornoXML,strEBS,removeAcentos(strTexto)
	set cSMS = nothing
			
	a.flush
end function 


function enviarSMSTESTE()
		
	strCanal = trim(request("canal"))
	
	Set oXmlMensage = Server.CreateObject("Microsoft.XMLHTTP")
	oXmlMensage.Open "GET", "http://200.150.102.98:8088/sms/mxml?action=ksendsms&device=b06C"&strCanal&"&destination=41992178092&message=Testando", false 	
	oXmlMensage.send()
	
	hora = hour(Now) 
	minuto = minute(Now)
	segundo =  second(Now)
			
	arquivo_xml = replace(date,"/","")&"_"&hora&"_"&minuto&"_"&segundo&"_"&strCanal&".xml"	
			
	Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = true
	success = xmlDoc.loadXML(oXmlMensage.ResponseXML.xml)
	If success = True Then    
			xmlDoc.save(server.MapPath("../XML/"&arquivo_xml))
	end if 
end function 


function enviarCMDPHP()
	
	strIdCampanha = trim(request("id_campanha"))
	
	Set oXmlMensage = Server.CreateObject("Microsoft.XMLHTTP")
	oXmlMensage.Open "GET", "http://localhost:81/sms/ctl_sms.php?id_campanha="&strIdCampanha&"", false 	
	oXmlMensage.send()
	
end function 

function alterarCanalDeSaida()
	strIdCampanha = trim(request("id_campanha"))
	
	set cSMS = new SMS
	  cSMS.alterarCanalDeSaida strIdCampanha
	set cSMS = nothing
	

end function 

function buscarQuantidadeEBS()
	
	set cEBS = new Ebs
		cEBS.buscarQuantidadeCanais 
		strQuantidadeCanais = cEBS.quantidade_canais_disponiveis
	set cEBS = nothing	
	
	
	Set a = jsArray()
	Set a(Null) = jsObject()
	a(Null)("strQuantidadeCanais") = strQuantidadeCanais
	a.flush

end function 

 
%>