<%
class Relatorio

	private m_data_incial
	private m_data_final
	private m_total_campanha
	private m_id_campanha
	private m_data_envio
	private m_email_retorno
	private m_telefone
	private m_mensagem
	private m_nu_contrato 
	private m_texto
	private m_data_resposta
	
	private m_qtde_msg_enviadas
	private m_qtde_msg_erros
	private m_qtde_msg_estouro
	
	private m_status_envio
	private m_mes_envio




	private m_listRelatorio

	public property get data_incial()
		data_incial = trim(m_data_incial)
	end property
	public property let data_incial(pDado)
		m_data_incial = pDado
	end property

	public property get data_final()
		data_final = trim(m_data_final)
	end property
	public property let data_final(pDado)
		m_data_final = pDado
	end property

	public property get total_campanha()
		total_campanha = trim(m_total_campanha)
	end property
	public property let total_campanha(pDado)
		m_total_campanha = pDado
	end property

	public property get email_retorno()
		email_retorno = trim(m_email_retorno)
	end property
	public property let email_retorno(pDado)
		m_email_retorno = pDado
	end property

	public property get id_campanha()
		id_campanha = trim(m_id_campanha)
	end property
	public property let id_campanha(pDado)
		m_id_campanha = pDado
	end property

	public property get data_envio()
		data_envio = trim(m_data_envio)
	end property
	public property let data_envio(pDado)
		m_data_envio = pDado
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

	public property get nu_contrato()
		nu_contrato = trim(m_nu_contrato)
	end property
	public property let nu_contrato(pDado)
		m_nu_contrato = pDado
	end property

	public property get texto()
		texto = trim(m_texto)
	end property
	public property let texto(pDado)
		m_texto = pDado
	end property

	public property get qtde_msg_enviadas()
		qtde_msg_enviadas = trim(m_qtde_msg_enviadas)
	end property
	public property let qtde_msg_enviadas(pDado)
		m_qtde_msg_enviadas = pDado
	end property

	public property get qtde_msg_erros()
		qtde_msg_erros = trim(m_qtde_msg_erros)
	end property
	public property let qtde_msg_erros(pDado)
		m_qtde_msg_erros = pDado
	end property

	public property get qtde_msg_estouro()
		qtde_msg_estouro = trim(m_qtde_msg_estouro)
	end property
	public property let qtde_msg_estouro(pDado)
		m_qtde_msg_estouro = pDado
	end property

	public property get data_resposta()
		data_resposta = trim(m_data_resposta)
	end property
	public property let data_resposta(pDado)
		m_data_resposta = pDado
	end property

	public property get status_envio()
		status_envio = trim(m_status_envio)
	end property
	public property let status_envio(pDado)
		m_status_envio = pDado
	end property

	public property get mes_envio()
		mes_envio = trim(m_mes_envio)
	end property
	public property let mes_envio(pDado)
		m_mes_envio = pDado
	end property



	sub Class_Initialize()
		m_data_incial = ""
		m_data_incial = ""
		m_total_campanha = ""
		m_email_retorno = ""
		m_id_campanha = ""
		m_data_envio = ""
		m_telefone = ""
		m_mensagem = ""
		m_nu_contrato = ""
		m_texto = ""
		m_qtde_msg_enviadas = ""
		m_qtde_msg_erros = ""
		m_qtde_msg_estouro = ""
		m_data_resposta = ""
		m_status_envio = ""
		m_mes_envio = ""
		set m_listRelatorio = Server.CreateObject("Scripting.Dictionary")
	end sub

	sub Class_Terminate()
		m_listRelatorio.RemoveAll()
		set m_listRelatorio = nothing
	end sub


	public function relBuscaCampanhaXEnvio(strDataInicial, strDataFinal)
		strEmpresa = Session("cd_empresa")
		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.CommandTimeout=3600
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select a.id_campanha,b.email_retorno, count(a.id_campanha) as total_campanha, "&_
		" (select top 1 convert(char(10),data_envio,103) from campanhas_envio where id_campanha = a.id_campanha) as data_envio,"&_
		" (select texto from campanhas where id_campanha = a.id_campanha) as texto, "&_
		" (select count(*) from campanhas_envio where id_campanha = a.id_campanha and status = 200) as qtde_msg_enviadas, "&_
		" (select count(*) from campanhas_envio where id_campanha = a.id_campanha and status = 400 and LEN(Ltrim(Rtrim(texto_envio))) <= 160 and qtde_tentativas > 0) as qtde_msg_erros, "&_
		" (select count(*) from campanhas_envio where id_campanha = a.id_campanha and LEN(texto_envio) > 160) as qtde_msg_estouro "&_
		" from campanhas_envio a  "&_
		" inner join campanhas b on a.id_campanha = b.id_campanha "&_
		" where a.data_envio between convert(datetime,'"&strDataInicial&"',103) and convert(datetime,'"&strDataFinal&" 23:59:00',103) "&_
		" and b.cd_empresa = '"&strEmpresa&"'"&_
		" group by a.id_campanha,b.email_retorno,a.id_campanha"&_
		" order by a.id_campanha "
		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			set cRelatorio = new Relatorio
			cRelatorio.id_campanha = objretorno("id_campanha")
			cRelatorio.total_campanha = objretorno("total_campanha")
			cRelatorio.email_retorno = objretorno("email_retorno")
			cRelatorio.data_envio = objretorno("data_envio")
			cRelatorio.texto = objretorno("texto")
			cRelatorio.qtde_msg_enviadas = objretorno("qtde_msg_enviadas")
			cRelatorio.qtde_msg_erros = objretorno("qtde_msg_erros")
			cRelatorio.qtde_msg_estouro = objretorno("qtde_msg_estouro")

			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend
		objretorno.Close


		'// BUSCA CAMPANHAS AVULSAS
		strSQL = _
		" select convert(char(10),data_envio,103) as data_envio,Ltrim(Rtrim(mensagem)) as texto, Substring(retorno_xml,1,3) as status from sms_avulso "&_
		" where data_envio between convert(datetime,'"&strDataInicial&"',103) and convert(datetime,'"&strDataFinal&" 23:59:00',103) "&_
		" and cd_empresa = '"&strEmpresa&"'"&_
		" order by data_envio "
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			set cRelatorio = new Relatorio
			cRelatorio.id_campanha = "AVULSA"
			cRelatorio.total_campanha = 1
			cRelatorio.email_retorno = ""
			cRelatorio.data_envio = objretorno("data_envio")
			cRelatorio.texto = objretorno("texto")

			if(objretorno("status") = "200") then
				cRelatorio.qtde_msg_enviadas = 1
			else
				cRelatorio.qtde_msg_enviadas = 0
			end if

			if(objretorno("status") = "400") then
				cRelatorio.qtde_msg_erros = 1
			else
				cRelatorio.qtde_msg_erros = 0
			end if

			cRelatorio.qtde_msg_estouro = 0

			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend
		objretorno.Close



		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		relBuscaCampanhaXEnvio = m_listRelatorio.Items
	end function

	public function buscarNumerosEnviados(strCampanha)
		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")


		strSQL = _
		"select telefone, id_campanha from campanha_envio "&_
		" where id_campanha = "&strCampanha

		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			set cRelatorio = new Relatorio
			cRelatorio.id_campanha = objretorno("id_campanha")
			cRelatorio.telefone = objretorno("telefone")
			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend

		objretorno.Close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		buscarNumerosEnviados = m_listRelatorio.Items

	end function

	public function relBuscaRetornos(strDataInicial, strDataFinal)
		strEmpresa = Session("cd_empresa")

		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		set objretorno2 = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select SUBSTRING(numero, 2, 15)as numero, SUBSTRING(numero, 4, 15)as numero_ddd,mensagem,telefoneXcontratos.nu_contrato as nu_contrato,telefoneXcontratos.cd_empresa from sms_retornos "&_
		" left join telefoneXcontratos on telefoneXcontratos.telefone = SUBSTRING(numero, 4, 15) "&_
		" where data_retorno between convert(datetime,'"&strDataInicial&"',103) and convert(datetime,'"&strDataFinal&" 23:59:00',103) "&_
		" and SUBSTRING(numero, 4, 15) <> '' "&_
		" and SUBSTRING(numero, 4, 15) <> '0' "&_
		" and Len(SUBSTRING(numero, 4, 15)) > 6 "&_
		" and cd_empresa = '"&strEmpresa&"' "&_
		" order by data_retorno desc"
		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			strSQL = _
			" select top 1 campanhas.texto  from campanha_envio "&_
			" join campanhas on campanha_envio.id_campanha = campanhas.id_campanha "&_
			" where campanha_envio.telefone =  '"&objretorno("numero_ddd")&"' "&_
			" and campanhas.cd_empresa = '"&strEmpresa&"' "&_
			" order by data_envio desc "
			objretorno2.Open strSQL, conMSSQL
			if not 	objretorno2.eof then
				strTextoCampanha = objretorno2("texto")
				set cRelatorio = new Relatorio
				cRelatorio.telefone = objretorno("numero")
				cRelatorio.mensagem = objretorno("mensagem")
				cRelatorio.nu_contrato = objretorno("nu_contrato")
				cRelatorio.texto = strTextoCampanha
				m_listRelatorio.Add aux, cRelatorio
				aux = aux + 1
			end if
			objretorno2.Close


			objretorno.MoveNext()
		wend

		objretorno.Close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		relBuscaRetornos = m_listRelatorio.Items
	end function


	public function relBuscaRetornosCampanha(strDataInicial, strDataFinal)
		strEmpresa = Session("cd_empresa")

		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		set objretorno2 = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select SUBSTRING(numero, 2, 15)as numero, SUBSTRING(numero, 4, 15)as numero_ddd,mensagem as resposta_cliente, "&_
		" texto_envio as mensagem_enviada,CONVERT(char(10),data_retorno,103) +'  '+CONVERT(char(10),data_retorno,108)  as data_resposta "&_
		" from sms_retornos "&_
		" left join campanhas_envio on campanhas_envio.telefone = SUBSTRING(numero, 4, 15) and campanhas_envio.cd_empresa = '"&strEmpresa&"' "&_
		"	and campanhas_envio.id_campanha = ( "&_
			" select top 1 id_campanha from campanhas_envio where campanhas_envio.telefone = SUBSTRING(numero, 4, 15)  "&_
			" and cd_empresa = '"&strEmpresa&"' order by data_envio asc "&_
		" ) "&_
		" where data_retorno between convert(datetime,'"&strDataInicial&"',103) and convert(datetime,'"&strDataFinal&" 23:59:00',103) "&_
		" and SUBSTRING(numero, 4, 15) <> '' "&_
		" and SUBSTRING(numero, 4, 15) <> '0' "&_
		" and Len(SUBSTRING(numero, 4, 15)) > 6 "&_
		" and cd_empresa = '"&strEmpresa&"' "&_
		" order by data_retorno asc"
		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			strDataRetroativa = Cdate(objretorno("data_resposta"))-2
			strMensagensEnviadas = ""
			strSQL = " select convert(char(10),data_envio,103) +'--'+ texto_envio as envio  from campanhas_envio "&_
			" where cd_empresa = '"&strEmpresa&"' "&_
			" and convert(char(10),data_envio,103) between (convert(char(10),'"&strDataRetroativa&"',103)) and (convert(char(10),GETDATE(),103)) "&_
			" and Ltrim(Rtrim(telefone))  = '"&objretorno("numero_ddd")&"' "
			objretorno2.Open strSQL, conMSSQL
				while not objretorno2.eof
					strMensagensEnviadas = strMensagensEnviadas &"<br/>"& objretorno2("envio")
					objretorno2.MoveNext()
				wend
			objretorno2.Close

			set cRelatorio = new Relatorio
				cRelatorio.telefone = objretorno("numero")
				cRelatorio.mensagem = objretorno("resposta_cliente")
				cRelatorio.texto = strMensagensEnviadas
				cRelatorio.data_resposta = objretorno("data_resposta")
			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend
		objretorno.Close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		relBuscaRetornosCampanha = m_listRelatorio.Items
	end function



	public function relBuscaNumerosInutilizados(strDataInicial, strDataFinal)

		strEmpresa = Session("cd_empresa")

		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select telefone as numero, id_campanha, convert(char(10),data_envio,103) as data_retorno from campanha_envio "&_
		" where data_envio between convert(datetime,'"&strDataInicial&"',103) and convert(datetime,'"&strDataFinal&" 23:59:00',103) "&_
		" and fl_numero_inutilizado = -1 "&_
		" and cd_empresa = '"&strEmpresa&"'"&_
		" order by data_envio desc"
		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			set cRelatorio = new Relatorio
			cRelatorio.telefone = objretorno("numero")
			cRelatorio.id_campanha = objretorno("id_campanha")
			cRelatorio.data_envio = objretorno("data_retorno")
			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend

		objretorno.Close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		relBuscaNumerosInutilizados = m_listRelatorio.Items
	end function

	public function gerarRelatorioCampanhaXEnvioContrato(strDataInicial, strDataFinal)
		strEmpresa = Session("cd_empresa")
		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.CommandTimeout=1360000
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select "&_
		" a.nu_contrato as contrato, "&_
		" b.texto,a.telefone "&_
		" from campanha_envio a "&_
		" inner join campanhas b on a.id_campanha = b.id_campanha "&_
		" where a.data_envio between convert(datetime,'"&strDataInicial&"',103) and convert(datetime,'"&strDataFinal&" 23:59:00',103) "&_
		" and b.cd_empresa = '"&strEmpresa&"'"
		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			set cRelatorio = new Relatorio
			cRelatorio.nu_contrato = objretorno("contrato")
			cRelatorio.telefone = objretorno("telefone")
			cRelatorio.texto = trim(objretorno("texto"))

			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend

		objretorno.Close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		gerarRelatorioCampanhaXEnvioContrato = m_listRelatorio.Items
	end function

	public function gerarRelatorioCampanhasCSV(strIdCampanha)
		strEmpresa = Session("cd_empresa")
		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.CommandTimeout=1360000
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select telefone, texto_envio, "&_
		" case when status = 400 then 'N' "&_
		" when status = 200 then 'S' end status_envio, "&_
		" nu_contrato "&_
		" from campanhas_envio "&_
		" where id_campanha = "&strIdCampanha

		aux = 0
		objretorno.Open strSQL, conMSSQL
		while not objretorno.eof
			set cRelatorio = new Relatorio
			cRelatorio.telefone = trim(objretorno("telefone"))
			cRelatorio.texto = trim(objretorno("texto_envio"))
			cRelatorio.status_envio = trim(objretorno("status_envio"))
			cRelatorio.nu_contrato = trim(objretorno("nu_contrato"))

			m_listRelatorio.Add aux, cRelatorio
			aux = aux + 1
			objretorno.MoveNext()
		wend

		objretorno.Close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		gerarRelatorioCampanhasCSV = m_listRelatorio.Items

	end function


	public function gerarRelatorioSMSEnviados()

		strDataRetroativa =  MID(date-19,7) &"-" & MID(date-19,4,2) & "-"&MID(date-19,1,2) & " 00:00:00"
		strDataAtual =  MID(date,7) &"-" & MID(date,4,2) & "-"&MID(date,1,2) & " 23:59:59"

		strEmpresa = Session("cd_empresa")

		m_listRelatorio.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")


		strSQL = _
		" select MONTH(data_envio) as mes,DAY(data_envio) as dia_envio,COUNT(*) as total from campanhas_envio "&_
		" where cd_empresa = '"&strEmpresa&"'"&_
		" and status = 200 "&_
		" and data_envio between '"&strDataRetroativa&"' and '"&strDataAtual&"' "&_
		" group by MONTH(data_envio),DAY(data_envio) "&_
		" order by mes desc,dia_envio asc "

   		 aux = 0
		 objretorno.Open strSQL, conMSSQL
		 while not objretorno.eof
			 set cRelatorio = new Relatorio
			 cRelatorio.data_envio = objretorno("dia_envio")
			 cRelatorio.mes_envio = objretorno("mes")
			 cRelatorio.qtde_msg_enviadas = objretorno("total")
			 m_listRelatorio.Add aux, cRelatorio
			 aux = aux + 1
			 objretorno.MoveNext()
		 wend

		 objretorno.Close

		 set objretorno = nothing
		 conMSSQL.close
		 set conMSSQL = nothing
		 gerarRelatorioSMSEnviados = m_listRelatorio.Items
	end function

	public function buscarConsumoMensal()

		strDataAtual =  MID(date,7) &"-" & MID(date,4,2) & "-"&MID(date,1,2) & " 23:59:59"
		strPrimeiroDIaMes = YEAR(Date)&"-"& completaComZeros(MONTH(date),2) &"-"& "01"& " 23:59:59"
		strEmpresa = Session("cd_empresa")
		strProximoMes = Month(strDataAtual)+1

		 m_listRelatorio.RemoveAll()
		 dim conMSSQL, objretorno
		 set conMSSQL = Server.CreateObject("ADODB.Connection")
		 conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		 conMSSQL.Open
		 set objretorno = server.CreateObject("ADODB.Recordset")


		strSQL = _
		" select DAY(data_envio) as dia_envio,COUNT(*) as total from campanhas_envio "&_
		" where cd_empresa = '"&strEmpresa&"'"&_
		" and status = 200 "&_
		" and data_envio between '"&strPrimeiroDIaMes&"' and '"&strDataAtual&"' "&_
		" group by DAY(data_envio) "&_
		" order by dia_envio asc "
		 aux = 0
		 objretorno.Open strSQL, conMSSQL
		  while not objretorno.eof
			  set cRelatorio = new Relatorio
			  cRelatorio.data_envio = objretorno("dia_envio")
		 	  cRelatorio.qtde_msg_enviadas = objretorno("total")
			  m_listRelatorio.Add aux, cRelatorio
			  aux = aux + 1
			  objretorno.MoveNext()
		  wend

		  objretorno.Close

		  set objretorno = nothing
		  conMSSQL.close
		  set conMSSQL = nothing
		  buscarConsumoMensal = m_listRelatorio.Items
	end function
	
end class
%>