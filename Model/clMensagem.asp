<%
class Mensagem

	private m_id_envio
	private m_telefone
	private m_mensagem
	private m_retorno_classe
	private m_listMensagem
	
	public property get id_envio()
		id_envio = trim(m_id_envio)
	end property
	public property let id_envio(pDado)
		m_id_envio = pDado
	end property
		
	public property get telefone()
		telefone = trim(m_telefone)
	end property
	public property let telefone(pDado)
		m_telefone = pDado
	end property

	public property get mensagem()
		mensagem = trim(m_mensagem)
	end property
	public property let mensagem(pDado)
		m_mensagem = pDado
	end property
		
	public property get retorno_classe()
		retorno_classe = trim(m_retorno_classe)
	end property
	public property let retorno_classe(pDado)
		m_retorno_classe = pDado
	end property
		
	sub Class_Initialize()
		m_id_envio = ""
		m_telefone = ""
		m_mensagem = ""
		m_retorno_classe = ""
		set m_listMensagem = Server.CreateObject("Scripting.Dictionary")
	end sub
	
	sub Class_Terminate()
		m_listMensagem.RemoveAll()
		set m_listMensagem = nothing
	end sub
	
	public function GravarMensagem(strTelefone, strMensagem,strNuContrato)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		strSQL = _
			" set nocount on " &_
			" insert into sms_avulso values('"&strTelefone&"',getdate(),'"&Session("cd_empresa")&"','"&strMensagem&"',0,'"&strNuContrato&"','','')" &_
			" select @@identity as id_envio " &_
			" set nocount off " 
			
			objretorno.Open strSQL, conMSSQL	
				m_id_campanha = objretorno("id_envio")
			objretorno.Close
		conMSSQL.close
		set conMSSQL = nothing
		
		
		autenticarUsuario()
		enviarSmsAvulso m_id_campanha,strTelefone,strMensagem
		
	end function 
	
	public function GravarContratos(strTelefone, strNuContrato)
		strEmpresa = Session("cd_empresa")
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _
			" select nu_contrato from telefoneXcontratos "&_
			" where cd_empresa = '"&strEmpresa&"'" &_
			" and telefone = '"&strTelefone&"'"
		objretorno.Open strSQL, conMSSQL			
		if objretorno.eof then 
			strSQLExec = _
				"insert into telefoneXcontratos "&_
				" values('"&strEmpresa&"','"&strTelefone&"','"&strNuContrato&"')"
			conMSSQL.execute strSQLExec
		else 
			if objretorno("nu_contrato") <> "" then
				strSQLExec = _
					" update telefoneXcontratos set nu_contrato = '"&strNuContrato&"' where cd_empresa = '"&strEmpresa&"' and telefone = '"&strTelefone&"' "
				conMSSQL.execute strSQLExec
			end if
		end if 
		objretorno.Close
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	
	function autenticarUsuario()
		' --- AUTENTICAÇÃO -----	
	  
		Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
		oXmlHTTP.Open "GET", "http://192.168.11.232:8088/sms/manager?action=login&username=hdr&secret=hdr17rdh", false 
		oXmlHTTP.send()
		set oXmlHTTP = nothing
	  
		' ---- FIM DA AUTENTICAÇÃO ----	
	end function
	
	function enviarSmsAvulso(id_envio,strTelefone,strMensagem)
		strEBS = "B00|0"
		m_retorno_classe = ""
		strEBSenvio = split(strEBS,"|")
		
		
		Set oXmlMensage = Server.CreateObject("Microsoft.XMLHTTP")
		oXmlMensage.Open "GET", "http://192.168.11.232:8088/sms/mxml?action=ksendsms&device="&strEBSenvio(0)&"&destination="&strTelefone&"&message="&strMensagem&"", false 		
		oXmlMensage.send()
	
		arquivo_xml = strTelefone&".xml"
		Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		xmlDoc.async = False
		success = xmlDoc.loadXML(oXmlMensage.ResponseXML.xml)
		
		If success = True Then    
			 'CASO SEJA NECESSÀRIO SALVAR O XML
			 'xmlDoc.save(server.MapPath("../XML/"&arquivo_xml))
				 
			Set Root = xmlDoc.documentElement
			strRetornoXML =  Root.childNodes(0).selectSingleNode("generic").attributes(0).text
			strMensagemRetorno =  Root.childNodes(0).selectSingleNode("generic").attributes(1).text
				 
			if(trim(strRetornoXML) = "Error") then 
				strRetornoXML = 400
				retornoClasse = "Msg não enviada Erro : "&strMensagemRetorno
			elseif (trim(strRetornoXML) = "Success") then 
				strRetornoXML = 200
				retornoClasse = "Mensagem Enviada com sucesso"
			end if 	
		
			atualizarRetorno id_envio,strRetornoXML,strMensagemRetorno,strEBSenvio
			m_retorno_classe = retornoClasse
		end if 
	
	end function
	
	function BuscarEbsCanal()
		
		retorno = "00-00"
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _
		" select top 1 id_ebs,cod_ebs, canal from tb_ebs "&_
		" where flg_ativo = 'S' "&_
		" and qtd_envio <= (select qtd_limite_envio_canal from parametros) "&_
		" order by qtd_tentativas asc "
		objretorno.open strSQL, conMSSQL, 1, 1
		if not objretorno.eof then 
			retorno = objretorno("cod_ebs")& "C"&objretorno("canal")&"|"&objretorno("id_ebs")
		end if 
		objretorno.close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		BuscarEbsCanal = retorno
	end function 
	
	function atualizarRetorno(id_envio,strRetornoXML,strMensagemRetorno,strEBSenvio)
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		strRetorno = split(strMensagemRetorno,"--")
		strSQL = _
			" update sms_avulso set retorno_xml = '"&strRetornoXML&"--"&strMensagemRetorno&"', ebs_saida = '"&strEBSenvio(0)&"', fl_mensagem_enviada = 1 where id_envio = "&id_envio
		conMSSQL.execute strSQL	
		
		
		strSQLTentativas = _
			" update tb_ebs set qtd_tentativas = (qtd_tentativas + 1 )  where id_ebs = "&strEBSenvio(1)
		conMSSQL.execute strSQLTentativas	
		
		if(strRetorno(0) = "200") then 
			strSQLTentativasSucesso = _
			" update tb_ebs set qtd_envio = (qtd_envio + 1 )  where id_ebs = "&strEBSenvio(1)
			conMSSQL.execute strSQLTentativasSucesso	
		end if
		
		conMSSQL.close
		set conMSSQL = nothing
	end function
	
end class
%>