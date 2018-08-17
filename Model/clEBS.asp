<%
class Ebs


	private m_nome_ebs
	private m_canal_ebs
	private m_id_ebs
	
	private m_qtd_envio
	private m_qtd_tentativas
	private m_flg_ativo
	private m_flg_em_uso
	private m_quantidade_canais_disponiveis
	
	public property get nome_ebs()
		nome_ebs = trim(m_nome_ebs)
	end property
	public property let nome_ebs(pDado)
		m_nome_ebs = pDado
	end property
	
	public property get canal_ebs()
		canal_ebs = trim(m_canal_ebs)
	end property
	public property let canal_ebs(pDado)
		m_canal_ebs = pDado
	end property
	
	public property get id_ebs()
		id_ebs = trim(m_id_ebs)
	end property
	public property let id_ebs(pDado)
		m_id_ebs = pDado
	end property
	
	public property get qtd_envio()
		qtd_envio = trim(m_qtd_envio)
	end property
	public property let qtd_envio(pDado)
		m_qtd_envio = pDado
	end property
		
	public property get qtd_tentativas()
		qtd_tentativas = trim(m_qtd_tentativas)
	end property
	public property let qtd_tentativas(pDado)
		m_qtd_tentativas = pDado
	end property
	
	public property get flg_ativo()
		flg_ativo = trim(m_flg_ativo)
	end property
	public property let flg_ativo(pDado)
		m_flg_ativo = pDado
	end property
	
	public property get flg_em_uso()
		flg_em_uso = trim(m_flg_em_uso)
	end property
	public property let flg_em_uso(pDado)
		m_flg_em_uso = pDado
	end property
	
	public property get quantidade_canais_disponiveis()
		quantidade_canais_disponiveis = trim(m_quantidade_canais_disponiveis)
	end property
	public property let quantidade_canais_disponiveis(pDado)
		m_quantidade_canais_disponiveis = pDado
	end property
	
		
	private m_listEbs
	private m_listCanaisEbs
	
	sub Class_Initialize()
		m_nome_ebs = ""
		m_canal_ebs = ""
		m_id_ebs = ""
		m_qtd_envio = ""
		m_qtd_tentativas = ""
		m_flg_ativo = ""
		m_flg_em_uso = ""
		m_quantidade_canais_disponiveis = ""
		set m_listEbs = Server.CreateObject("Scripting.Dictionary")
		set m_listCanaisEbs = Server.CreateObject("Scripting.Dictionary")
	end sub
	
	sub Class_Terminate()
		m_listEbs.RemoveAll()
		set m_listEbs = nothing
		m_listCanaisEbs.RemoveAll()
		set m_listCanaisEbs = nothing
	end sub
	
	public function BuscarEbsCanal()
		m_listEbs.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _
		" select top 1 id_ebs,cod_ebs,canal from tb_ebs WITH (NOLOCK) "&_
		" where flg_ativo = 'S' "&_
		" and flg_em_uso = 'N' "&_
		" order by qtd_envio asc " 
		objretorno.open strSQL, conMSSQL, 1, 1
		if not objretorno.eof then 
			set cEbs = new Ebs
			m_nome_ebs = objretorno("cod_ebs")
			m_canal_ebs = objretorno("canal")
			m_id_ebs = objretorno("id_ebs")
		end if 
		objretorno.close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function gravarUtilizacaoEbs(strIDEbs)
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		
		strSQL = _
		" update tb_ebs set flg_em_uso = 'S'  where id_ebs = "&strIDEbs 
		
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	
	end function 
	
	public function gravarLiberacaoEBS(strIDEbs)
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		
		strSQL = _
		" update tb_ebs set flg_em_uso = 'N',  qtd_envio = (qtd_envio + 1) where id_ebs = "&strIDEbs 
		
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	
	end function 
	
	public function buscarEBS()
		m_listEbs.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _
		" select distinct cod_ebs from tb_ebs "	
		objretorno.Open strSQL, conMSSQL, 1, 1	
		aux = 0
		while not objretorno.eof
			set cEbs = new Ebs
			cEbs.nome_ebs = objretorno("cod_ebs")
			m_listEbs.Add aux, cEbs
			aux = aux + 1
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarEBS = m_listEbs.Items
	end function 
	
	public function buscarCanais(strNomeEbs)
		m_listCanaisEbs.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _		
		" select id_ebs,cod_ebs,canal,qtd_envio,qtd_tentativas,flg_ativo,flg_em_uso from tb_ebs "&_
		" where cod_ebs = '"&strNomeEbs&"' "&_
		" order by canal  "
		
		objretorno.Open strSQL, conMSSQL, 1, 1	
		aux = 0
		while not objretorno.eof
			set cEbs = new Ebs
			cEbs.id_ebs = objretorno("id_ebs")
			cEbs.nome_ebs = objretorno("cod_ebs")
			cEbs.canal_ebs = objretorno("canal")
			cEbs.qtd_envio = objretorno("qtd_envio")
			cEbs.qtd_tentativas = objretorno("qtd_tentativas")
			cEbs.flg_ativo = objretorno("flg_ativo")
			cEbs.flg_em_uso = objretorno("flg_em_uso")
				
			m_listCanaisEbs.Add aux, cEbs
			aux = aux + 1
			
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarCanais = m_listCanaisEbs.Items
	end function 
	
	public function buscarQuantidadeCanais()
		m_listCanaisEbs.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		strSQL = _		
		" select COUNT(*) as total_canais from tb_ebs "&_ 
		" where flg_ativo = 'S' "&_
		" and flg_em_uso = 'N'" 
		objretorno.Open strSQL, conMSSQL, 1, 1	
		aux = 0
		if not objretorno.eof then 
			m_quantidade_canais_disponiveis = objretorno("total_canais")
		end if 	
			
		objretorno.Close
		
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
	end function 
	
	public function mudarOpcaoStatus(strIDCanal,strHabilitaOpcao)
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		
		strSQL = _
		" update tb_ebs set flg_ativo = '"&strHabilitaOpcao&"' where id_ebs = "&strIDCanal 
	
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function zerarEnvioCanal(strIDCanal)
		dim conMSSQL, objretorno, clEmpresa
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		
		strSQL = _
		" update tb_ebs set qtd_envio = 0, qtd_tentativas = 0 where id_ebs = "&strIDCanal 
	
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
end class
%>