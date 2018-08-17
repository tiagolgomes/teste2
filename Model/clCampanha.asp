<%
class Campanha

	private m_id_campanha
	private m_email_retorno
	private m_texto
	private m_telefone
	private m_tg_1
	private m_tg_2
	private m_tg_3
	private m_id_padrao
	private m_fl_limite_sms
	private m_fl_numero_inutilizado
	private m_data_envio
	
	private m_listCampanha
	
	public property get id_campanha()
		id_campanha = trim(m_id_campanha)
	end property
	public property let id_campanha(pDado)
		m_id_campanha = pDado
	end property
	
	public property get email_retorno()
		email_retorno = trim(m_email_retorno)
	end property
	public property let email_retorno(pDado)
		m_email_retorno = pDado
	end property
	
	public property get texto()
		texto = trim(m_texto)
	end property
	public property let texto(pDado)
		m_texto = pDado
	end property
	
	public property get telefone()
		telefone = trim(m_telefone)
	end property
	public property let telefone(pDado)
		m_telefone = pDado
	end property
	
	public property get tg_1()
		tg_1 = trim(m_tg_1)
	end property
	public property let tg_1(pDado)
		m_tg_1 = pDado
	end property
	
	public property get tg_2()
		tg_2 = trim(m_tg_2)
	end property
	public property let tg_2(pDado)
		m_tg_2 = pDado
	end property
	
	public property get tg_3()
		tg_3 = trim(m_tg_3)
	end property
	public property let tg_3(pDado)
		m_tg_3 = pDado
	end property
	
	public property get id_padrao()
		id_padrao = trim(m_id_padrao)
	end property
	public property let id_padrao(pDado)
		m_id_padrao = pDado
	end property
	
	public property get numeros_envio_mensal()
		numeros_envio_mensal = trim(m_numeros_envio_mensal)
	end property
	public property let numeros_envio_mensal(pDado)
		m_numeros_envio_mensal = pDado
	end property
	
	public property get fl_limite_sms()
		fl_limite_sms = trim(m_fl_limite_sms)
	end property
	public property let fl_limite_sms(pDado)
		m_fl_limite_sms = pDado
	end property
	
	public property get fl_numero_inutilizado()
		fl_numero_inutilizado = trim(m_fl_numero_inutilizado)
	end property
	public property let fl_numero_inutilizado(pDado)
		m_fl_numero_inutilizado = pDado
	end property
	
	public property get data_envio()
		data_envio = trim(m_data_envio)
	end property
	public property let data_envio(pDado)
		m_data_envio = pDado
	end property
	
	
		
	sub Class_Initialize()
		m_id_campanha = ""
		m_email_retorno = ""
		m_texto = ""
		m_telefone = ""
		m_tg_1 = ""
		m_tg_2 = ""
		m_tg_3 = ""
		m_id_padrao = ""
		m_numeros_envio_mensal = ""
		m_fl_limite_sms = ""
		m_fl_numero_inutilizado = ""
		m_data_envio = ""
		set m_listCampanha = Server.CreateObject("Scripting.Dictionary")
	end sub

	'destrutor
	sub Class_Terminate()
		m_listCampanha.RemoveAll()
		set m_listCampanha = nothing
	end sub
	
	public function GravarCampanha(strIdMeioComunicacao, strEmailRetorno, strTexto,flRetorno,strDataAgendamento,strHoraAgendamento)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		strSQL = _
			" set nocount on " &_
			" insert into campanhas values("&strIdMeioComunicacao&",'"&strEmailRetorno&"','"&strTexto&"','0','"&Session("cd_empresa")&"',0,"&flRetorno&",'"&strDataAgendamento&"','"&strHoraAgendamento&"','N')" &_
			" select @@identity as id_campanha " &_
			" set nocount off " 
			
			objretorno.Open strSQL, conMSSQL	
				m_id_campanha = objretorno("id_campanha")
			objretorno.Close
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function BuscaCampanhaId(id_campanha)
		strEmpresa = Session("cd_empresa")
	
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select id_campanha, email_retorno,texto,fl_limite_sms from campanhas "&_
		" where id_campanha = "&id_campanha&_
		" and cd_empresa = '"&strEmpresa&"'"
		
		objretorno.Open strSQL, conMSSQL	
			m_id_campanha = objretorno("id_campanha")
			m_email_retorno = objretorno("email_retorno")
			m_texto = objretorno("texto")
			m_fl_limite_sms = objretorno("fl_limite_sms")
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
	end function	
	
	public function BuscaCampanhaEnvio(id_campanha)
		m_listCampanha.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select telefone, campanhas.texto,tg_1, tg_2,tg_3 from campanha_envio "&_
		" inner join campanhas on campanha_envio.id_campanha = campanhas.id_campanha "&_
		" where campanha_envio.id_campanha = "&id_campanha&" "&_
		" order by ordem "
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			set cCampanha = new Campanha
			cCampanha.telefone = objretorno("telefone")
			cCampanha.texto = objretorno("texto")
			cCampanha.tg_1 = objretorno("tg_1")
			cCampanha.tg_2 = objretorno("tg_2")
			cCampanha.tg_3 = objretorno("tg_3")
			m_listCampanha.Add aux, cCampanha
			aux = aux + 1
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		BuscaCampanhaEnvio = m_listCampanha.Items
	end function	
	
		
	public function buscarCampanhasPadrao()
		
		strEmpresa = Session("cd_empresa")
		
		m_listCampanha.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select top 5 id_padrao, texto from campanhas_padrao where cd_empresa = '"&strEmpresa&"' order by id_padrao desc"
		
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			set cCampanha = new Campanha
			cCampanha.id_padrao = objretorno("id_padrao")
			cCampanha.texto = objretorno("texto")
			m_listCampanha.Add aux, cCampanha
			aux = aux + 1
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarCampanhasPadrao = m_listCampanha.Items
	end function	
	
	
	public function excluirPadrao(strPadrao)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		strSQL = _
			" delete from campanhas_padrao where id_padrao = "&strPadrao&" "
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function GravarCampanhasPadrao(strTexto)
		strEmpresa = Session("cd_empresa")	
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		strSQL = _
			" insert into campanhas_padrao values('"&strTexto&"','"&strEmpresa&"') "
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function BuscaCampanhaQuantidadeDeEnviosMensal()
	
		strEmpresa = Session("cd_empresa")
		mesAtual = Month(now)
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select count(telefone) as total from campanha_envio "&_
		" where convert(varchar(MAX),cd_empresa) = '"&strEmpresa&"'" &_
		" and Month(data_envio) = '"& mesAtual &"'"
		 
		 Session("numeros_envio_mensal") = 0
 		 objretorno.Open strSQL, conMSSQL	
	 	 if not objretorno.eof then
			Session("numeros_envio_mensal") = objretorno("total")
		 end if 
		 objretorno.Close
		 
		 set objretorno = nothing
		 conMSSQL.close
		 set conMSSQL = nothing		
	end function	
	
	public function BuscaCampanhaQuantidadeDeEnviosMensalEmpresa()
	
		strEmpresa = Session("cd_empresa")
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select qt_envio_mensal  from empresas "&_
		" where cd_empresa = '"&strEmpresa&"'" 
		 
		 Session("numeros_envio_mensal_empresa") = 0
 		 objretorno.Open strSQL, conMSSQL	
	 	 if not objretorno.eof then
		 	Session("numeros_envio_mensal_empresa") = objretorno("qt_envio_mensal")
		 end if 
		 objretorno.Close
		 
		 set objretorno = nothing
		 conMSSQL.close
		 set conMSSQL = nothing		
	end function	
	
	public function buscarNumero(strNumero)
	
		m_listCampanha.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select id_campanha, telefone,convert(char(10),data_envio,103) as data_envio, coalesce(fl_numero_inutilizado,'') as fl_numero_inutilizado from campanha_envio "&_
		" where telefone = '"&strNumero&"'"&_
		" order by id_campanha desc"
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			set cCampanha = new Campanha
			cCampanha.id_campanha = objretorno("id_campanha")
			cCampanha.telefone = objretorno("telefone")
			cCampanha.data_envio = objretorno("data_envio")
			cCampanha.fl_numero_inutilizado = objretorno("fl_numero_inutilizado")
			m_listCampanha.Add aux, cCampanha
			aux = aux + 1
			objretorno.MoveNext()
		wend	
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarNumero = m_listCampanha.Items
	end function 
	
	public function inutilizarNumero(strIdCampanha,strTelefone,strNomeArquivo)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		
		if(strNomeArquivo = "check") then 
		strSQL = _
			" update campanha_envio set fl_numero_inutilizado = '-1' "&_
			" where id_campanha = "&strIdCampanha&_
			" and telefone = '"&strTelefone&"'"
		else 
		strSQL = _
			" update campanha_envio set fl_numero_inutilizado = '0' "&_
			" where id_campanha = "&strIdCampanha&_
			" and telefone = '"&strTelefone&"'"
		end if 
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function gravarContratos(strIdCampanha)
		
		strEmpresa = Session("cd_empresa")
		
		dim conMSSQL, objretorno,objretornoExec
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")	
		set objretornoExec = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _
		" select telefone,nu_contrato from campanha_envio "&_
		" where id_campanha = "&strIdCampanha
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			strSQL = _
			" select nu_contrato from telefoneXcontratos "&_
			" where cd_empresa = '"&strEmpresa&"'" &_
			" and telefone = '"&objretorno("telefone")&"'"
			objretornoExec.Open strSQL, conMSSQL		
			if objretornoExec.eof then 
				strSQLExec = _
					"insert into telefoneXcontratos "&_
					" values('"&strEmpresa&"','"&objretorno("telefone")&"','"&objretorno("nu_contrato")&"')"
				conMSSQL.execute strSQLExec
			else 
				if objretorno("nu_contrato") <> "" then
					strSQLExec = _
						" update telefoneXcontratos set nu_contrato = '"&objretorno("nu_contrato")&"' where cd_empresa = '"&strEmpresa&"' and telefone = '"&objretorno("telefone")&"' "
					conMSSQL.execute strSQLExec
				end if
			end if 
			objretornoExec.Close
			objretorno.MoveNext()
		wend	
		objretorno.Close
		
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function excluirCampanha(strIdCampanha)
		strEmpresa = Session("cd_empresa")
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		strSQL = _
			" delete from campanhas where id_campanha = "&strIdCampanha&" and cd_empresa = '"&strEmpresa&"' "
		conMSSQL.execute strSQL
		strSQL = _
			" delete from campanha_envio where id_campanha = "&strIdCampanha&" and cd_empresa = '"&strEmpresa&"' "
		conMSSQL.execute strSQL
		
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function IntegrarSMS()
	
		dim conMSSQL, objretorno,objretornoExec
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _
		" select id_campanha,email_retorno from campanhas "&_
		" where fl_campanha_enviada = 0 "&_
		" and fl_integracao = 'S' "
	
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			set cCampanha = new Campanha
			cCampanha.id_campanha = objretorno("id_campanha")
			cCampanha.email_retorno = objretorno("email_retorno")
			m_listCampanha.Add aux, cCampanha
			aux = aux + 1
			objretorno.MoveNext()
		wend	
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		IntegrarSMS = m_listCampanha.Items
		
	end function
	
end class
%>