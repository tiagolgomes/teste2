<%
class SMS

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
	private m_id_campanha_envio
	private m_nu_contrato
	
	
	private m_listCampanha
	private m_listNumerosErro
	
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
	
	public property get id_campanha_envio()
		id_campanha_envio = trim(m_id_campanha_envio)
	end property
	public property let id_campanha_envio(pDado)
		m_id_campanha_envio = pDado
	end property
		
	public property get nu_contrato()
		nu_contrato = trim(m_nu_contrato)
	end property
	public property let nu_contrato(pDado)
		m_nu_contrato = pDado
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
		m_id_campanha_envio = ""
		m_nu_contrato = ""
		set m_listCampanha = Server.CreateObject("Scripting.Dictionary")
		set m_listNumerosErro = Server.CreateObject("Scripting.Dictionary")
	end sub

	'destrutor
	sub Class_Terminate()
		m_listCampanha.RemoveAll()
		set m_listCampanha = nothing
		m_listNumerosErro.RemoveAll()
		set m_listNumerosErro = nothing
	end sub
	
	public function GravarCampanha(strIdMeioComunicacao, strEmailRetorno, strTexto,flRetorno,strDataAgendamento,strHoraAgendamento)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		strSQL = _
			" set nocount on " &_
			" insert into campanhas values("&strIdMeioComunicacao&",'"&strEmailRetorno&"','"&strTexto&"','0','"&Session("cd_empresa")&"',0,"&flRetorno&",'"&strDataAgendamento&"','"&strHoraAgendamento&"','N',0)" &_
			" select @@identity as id_campanha " &_
			" set nocount off " 
			
			objretorno.Open strSQL, conMSSQL	
				m_id_campanha = objretorno("id_campanha")
			objretorno.Close
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
	
	public function GravarCampanhasEnvio(strIdCampanha,strTelefone,strTag1,strTag2,strTag3,strContrato,strRetornoXML,strEbsSaida,strTexto)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strEmpresa = Session("cd_empresa")
		
		'strEbsSaida = buscarEbsSaida()
		
		strSQL = _
			" set nocount on " &_
			" insert into campanhas_envio "&_
			" values('"&strTelefone&"','"&strTag1&"','"&strTag2&"','"&strTag3&"',getdate(),"&strIdCampanha&",'"&strEmpresa&"',0,'"&strContrato&"',1,'"&strRetornoXML&"','"&strEbsSaida&"','"&strTexto&"',400,0) "&_
			" select @@identity as id_campanha_envio " &_
			" set nocount off " 
		
		objretorno.Open strSQL, conMSSQL	
			m_id_campanha_envio = objretorno("id_campanha_envio")
		objretorno.Close
		
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function buscarNumerosComErro(id_campanha)
		m_listNumerosErro.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select telefone,tg_1, tg_2,tg_3,nu_contrato  from campanhas_envio WITH (NOLOCK) "&_
		" where id_campanha = "&id_campanha&" "&_
		" and retorno_xml = '200' "
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			set cSMS = new SMS
			cSMS.telefone = objretorno("telefone")
			cSMS.tg_1 = objretorno("tg_1")
			cSMS.tg_2 = objretorno("tg_2")
			cSMS.tg_3 = objretorno("tg_3")
			cSMS.nu_contrato = objretorno("nu_contrato")
			m_listNumerosErro.Add aux, cSMS
			aux = aux + 1
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarNumerosComErro = m_listNumerosErro.Items
	end function	
	
	public function alterarCanalDeSaida(id_campanha)
		m_listNumerosErro.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		
		strSQL = _
		" select id_campanha_envio,ebs_saida from campanhas_envio WITH (NOLOCK) "&_
		" where id_campanha = "&id_campanha&" "&_
		" and status = '400' "
		
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			
			strEbsSaida = buscarEbsSaida()
			strSQLExecucaoReenvio =  " update campanhas_envio set ebs_saida = '"&strEbsSaida&"' where id_campanha_envio = "&trim(objretorno("id_campanha_envio"))
			conMSSQL.execute strSQLExecucaoReenvio	
			
			strCodEBS = split(strEbsSaida,"--")
			 
			strSQLExecucao =  " update tb_ebs set qtd_tentativas = (qtd_tentativas + 1 )  where cod_ebs =	'"&strCodEBS(0)&"' and canal = "&Cint(strCodEBS(1))&" "
			conMSSQL.execute strSQLExecucao	
			
			objretorno.MoveNext()
		wend	
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
	end function	
	
	function buscarEbsSaida()
		m_listNumerosErro.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")		
		strRetorno = ""
		strSQL = _
		" select top 1 Ltrim(Rtrim(convert(varchar(MAX),cod_ebs) +'--'+convert(varchar(MAX),canal))) as retorno_ebs,id_ebs from tb_ebs "&_
		" where flg_ativo = 'S' "&_ 
		" and flg_em_uso = 'N' "&_
		" order by qtd_tentativas asc "
		objretorno.Open strSQL, conMSSQL	
		if not objretorno.eof then 
			strRetorno = trim(objretorno("retorno_ebs"))
			
			strIdEbs = trim(objretorno("id_ebs"))
			strSQLExecucao =  " update tb_ebs set qtd_tentativas = (qtd_tentativas + 1 )  where id_ebs = "&strIdEbs
			conMSSQL.execute strSQLExecucao
		end if 	
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarEbsSaida = strRetorno
	end function 
	
	public function AtualizarCampanhaEnvio(strRetornoXML,strMensagemRetorno,id_campanha_envio,strMensagemEBS)
		
		strEbs = Split(trim(strMensagemEBS),"C")
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQLExecucao = "update campanhas_envio set retorno_xml = '"&strMensagemRetorno&"', qtde_tentativas = (qtde_tentativas + 1 ), "&_
		" status = '"&strRetornoXML&"', "&_
		" ebs_saida = '"&strEbs(0)&"--"&strEbs(1)&"' "&_
		"where id_campanha_envio = "&id_campanha_envio
		conMSSQL.execute strSQLExecucao
		
		if(trim(strRetornoXML) = "200") then 
			strSQLExecucaoEnvios = " update tb_ebs set qtd_envio = (qtd_envio + 1 )  where cod_ebs = (select SUBSTRING(ebs_saida,1,3) from campanhas_envio where id_campanha_envio = "&id_campanha_envio&") "&_
			" and canal = (select SUBSTRING(ebs_saida,6,2) from campanhas_envio where id_campanha_envio = "&id_campanha_envio&") "
			conMSSQL.execute strSQLExecucaoEnvios
			
			strEbsAtualizacao = ""
			if(len(trim(strEbs(0))) = 2 )then 
				strEbsAtualizacao =  CSTR(Mid(strEbs(0),1,1)) &"0" & CSTR(Mid(strEbs(0),2,1))
			else
				strEbsAtualizacao = trim(strEbs(0))
			end if 
			
			strSQLExecucaoEbs =  "update tb_ebs set qtd_envio = (qtd_envio + 1)  where cod_ebs = '"&strEbsAtualizacao&"' and canal = '"&strEbs(1)&"' "
			conMSSQL.execute strSQLExecucaoEbs
		end if 
		
		conMSSQL.close
		set conMSSQL = nothing
		set objretorno = nothing
	
	end function
	
	
end class
%>