<%
class Empresa

	private m_cd_empresa
	private m_de_empresa
	private m_qt_envio_mensal
	private m_qtd_total_sms_enviados_mes
	
	private m_listEmpresa
	
	public property get cd_empresa()
		cd_empresa = trim(m_cd_empresa)
	end property
	public property let cd_empresa(pDado)
		m_cd_empresa = pDado
	end property
	
	public property get de_empresa()
		de_empresa = trim(m_de_empresa)
	end property
	public property let de_empresa(pDado)
		m_de_empresa = pDado
	end property
	
	public property get qt_envio_mensal()
		qt_envio_mensal = trim(m_qt_envio_mensal)
	end property
	public property let qt_envio_mensal(pDado)
		m_qt_envio_mensal = pDado
	end property
	
	public property get qtd_total_sms_enviados_mes()
		qtd_total_sms_enviados_mes = trim(m_qtd_total_sms_enviados_mes)
	end property
	public property let qtd_total_sms_enviados_mes(pDado)
		m_qtd_total_sms_enviados_mes = pDado
	end property
	
	'construtor
	sub Class_Initialize()
		m_cd_empresa = ""
		m_de_empresa = ""
		m_qt_envio_mensal = ""
		m_qtd_total_sms_enviados_mes = ""
		set m_listEmpresa = Server.CreateObject("Scripting.Dictionary")
	end sub

	'destrutor
	sub Class_Terminate()
		m_listEmpresa.RemoveAll()
		set m_listEmpresa = nothing
	end sub

	public function salvarEmpresa(strEmpresa, strDeEmpresa,strQtEnvioMensal)
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		
		strSQL = _
		" select cd_empresa,de_descricao_empresa, qt_envio_mensal from empresas where cd_empresa = '"&strEmpresa&"' "
		objretorno.open strSQL, conMSSQL, 1, 1
		if objretorno.eof then 
			strSQL = _
			" insert into empresas values('"&strEmpresa&"','"&strDeEmpresa&"', "&strQtEnvioMensal&", getdate())" 
		else
			strSQL = _
			" update empresas set de_descricao_empresa = '"&strDeEmpresa&"', qt_envio_mensal = "&strQtEnvioMensal&" where cd_empresa = '"&strEmpresa&"' " 
		end if 
		objretorno.close
		
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function carregarEmpresas()
	
		m_listEmpresa.RemoveAll()
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select cd_empresa,de_descricao_empresa, qt_envio_mensal from empresas order by cd_empresa "
			
		objretorno.open strSQL, conMSSQL, 1, 1
		while not objretorno.eof
			set cEmpresa = new Empresa
			cEmpresa.cd_empresa = objretorno("cd_empresa")
			cEmpresa.de_empresa = objretorno("de_descricao_empresa")
			cEmpresa.qt_envio_mensal = objretorno("qt_envio_mensal")
			m_listEmpresa.Add objretorno.AbsolutePosition, cEmpresa
			objretorno.MoveNext()
		wend
		objretorno.close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		
		carregarEmpresas = m_listEmpresa.items
	end function 
	
	public function carregarEmpresaPorId(strEmpresa)
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select cd_empresa,de_descricao_empresa, qt_envio_mensal from empresas where cd_empresa = '"&strEmpresa&"' "
		
		objretorno.open strSQL, conMSSQL, 1, 1
		if not objretorno.eof then 
			set cEmpresa = new Empresa
			m_cd_empresa = objretorno("cd_empresa")
			m_de_empresa = objretorno("de_descricao_empresa")
			m_qt_envio_mensal = objretorno("qt_envio_mensal")
		end if 
		objretorno.close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function buscarTotalSMSEnviados()
		strEmpresa = Session("cd_empresa")
		
		m_listEmpresa.RemoveAll()
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strEnvioCampanhas = 0
		strEnvioAvulso = 0
		strSQL = _
		" select COUNT(id_campanha_envio) as total_envio_mes_campanha from campanhas_envio "&_
		" where MONTH(data_envio) = (select MONTH(GETDATE())) "&_
		" and YEAR(data_envio) = (select (YEAR(GETDATE()))) "&_
		" and cd_empresa = '"&strEmpresa&"' "&_
		" and status = '200' "
		objretorno.open strSQL, conMSSQL, 1, 1
		if not objretorno.eof then 
			strEnvioCampanhas = objretorno("total_envio_mes_campanha")
		end if 
		objretorno.close
		
		
		strSQL = _
		" select COUNT(id_envio) as total_envio_avulso from sms_avulso "&_
		" where MONTH(data_envio) = (select MONTH(GETDATE())) "&_
		" and YEAR(data_envio) = (select (YEAR(GETDATE()))) "&_
		" and cd_empresa = '"&strEmpresa&"' "&_
		" and  Rtrim(Ltrim(SUBSTRING(retorno_xml,0,4))) = '200' "
		objretorno.open strSQL, conMSSQL, 1, 1
		if not objretorno.eof then 
			strEnvioAvulso = objretorno("total_envio_avulso")
		end if 
		objretorno.close
		
		m_qtd_total_sms_enviados_mes = CLNG(CLNG(strEnvioCampanhas) + CLNG(strEnvioAvulso))
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	
end class
%>