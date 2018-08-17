<%
class Mensagem

	private m_id_envio
	private m_telefone
	private m_mensagem
	
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
		
	sub Class_Initialize()
		m_id_envio = ""
		m_telefone = ""
		m_mensagem = ""
		set m_listMensagem = Server.CreateObject("Scripting.Dictionary")
	end sub
	
	sub Class_Terminate()
		m_listMensagem.RemoveAll()
		set m_listMensagem = nothing
	end sub
	
	public function GravarMensagem(strTelefone, strMensagem)
		
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		strSQL = _
			" set nocount on " &_
			" insert into sms_avulso values('"&strTelefone&"',getdate(),'"&Session("cd_empresa")&"','"&strMensagem&"',0)" &_
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
		oXmlHTTP.Open "GET", "http://200.150.102.98:8088/sms/manager?action=login&username=hdr&secret=hdr17rdh", false 
		oXmlHTTP.send()
		set oXmlHTTP = nothing
	  
		' ---- FIM DA AUTENTICAÇÃO ----	
	end function
	
	function enviarSmsAvulso(id_envio,strTelefone,strMensagem)
		strEBS = BuscarEbsCanal()
		
		Set oXmlMensage = Server.CreateObject("Microsoft.XMLHTTP")
		oXmlMensage.Open "GET", "http://200.150.102.98:8088/sms/mxml?action=ksendsms&device="&strEBS&"&destination="&strTelefone&"&message="&strMensagem&"", false 	
		oXmlMensage.send()
			
		Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		xmlDoc.async = False
		success = xmlDoc.loadXML(oXmlMensage.ResponseXML.xml)
		If success = True Then    
			 'CASO SEJA NECESSÀRIO SALVAR O XML
			 'xmlDoc.save(server.MapPath("../XML/"&arquivo_xml))
				 
			Set Root = xmlDoc.documentElement
			strRetornoXML =  Root.childNodes(0).selectSingleNode("generic").attributes(0).text
				 
			if(trim(strRetornoXML) = "Error") then 
				strRetornoXML = 200
			elseif (trim(strRetornoXML) = "Success") then 
				strRetornoXML = 400
			end if 	
		
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
		" select top 1 id_ebs,cod_ebs, canal from tb_ebs "
		" where flg_ativo = 'S' "
		" and qtd_envio <= (select qtd_limite_envio_canal from parametros) "
		" order by qtd_tentativas asc "
		objretorno.open strSQL, conMSSQL, 1, 1
		if not objretorno.eof then 
			retorno = objretorno["cod_ebs"]& "C"&objretorno["canal"] 
		end if 
		objretorno.close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		BuscarEbsCanal = retorno
	end function 
	

end class
%>