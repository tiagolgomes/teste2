<%
class BlackList

	private m_id_blackList
	private m_telefone
	private m_data_alteracao
	private m_id_ultimo_usuario_alterou
	private m_flg_ativo
	private m_cd_empresa
	private m_nome_ultimo_usuario_alterou
	
	public property get id_blackList()
		id_blackList = trim(m_id_blackList)
	end property
	public property let id_blackList(pDado)
		m_id_blackList = pDado
	end property
	
	public property get telefone()
		telefone = trim(m_telefone)
	end property
	public property let telefone(pDado)
		m_telefone = pDado
	end property
	
	public property get data_alteracao()
		data_alteracao = trim(m_data_alteracao)
	end property
	public property let data_alteracao(pDado)
		m_data_alteracao = pDado
	end property
	
	public property get id_ultimo_usuario_alterou()
		id_ultimo_usuario_alterou = trim(m_id_ultimo_usuario_alterou)
	end property
	public property let id_ultimo_usuario_alterou(pDado)
		m_id_ultimo_usuario_alterou = pDado
	end property
	
	public property get flg_ativo()
		flg_ativo = trim(m_flg_ativo)
	end property
	public property let flg_ativo(pDado)
		m_flg_ativo = pDado
	end property
	
	public property get cd_empresa()
		cd_empresa = trim(m_cd_empresa)
	end property
	public property let cd_empresa(pDado)
		m_cd_empresa = pDado
	end property
	
	public property get nome_ultimo_usuario_alterou()
		nome_ultimo_usuario_alterou = trim(m_nome_ultimo_usuario_alterou)
	end property
	public property let nome_ultimo_usuario_alterou(pDado)
		m_nome_ultimo_usuario_alterou = pDado
	end property
	
	
	private m_listBlackList
	
	sub Class_Initialize()
		m_id_blackList = ""
		m_telefone = ""
		m_data_alteracao = ""
		m_id_ultimo_usuario_alterou = ""
		m_flg_ativo = ""
		m_cd_empresa = ""		
		m_nome_ultimo_usuario_alterou = ""
		set m_listBlackList = Server.CreateObject("Scripting.Dictionary")
	end sub

	sub Class_Terminate()
		m_listBlackList.RemoveAll()
		set m_listBlackList = nothing
	end sub
	
	public function GravarBalckList(strTelefone,strIdBlackList,strFlgAtivo)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		strEmpresa = Session("cd_empresa")	
		strIdUsuario = Session("id_usuario") 
		
		if(trim(strIdBlackList) = "N") then 
			strSQL = _
			" insert into blackList_sms values('"&strTelefone&"', getdate(),"&strIdUsuario&",'S','"&strEmpresa&"') "
		else 
			strSQL = _
			" update blackList_sms set flg_ativo = '"&strFlgAtivo&"', "&_
			" data_alteracao = getdate(), "&_
			" id_ultimo_usuario_alterou = '"&strIdUsuario&"' "&_
			" where id_blackList = "&strIdBlackList	
		end if 		
			
		conMSSQL.Execute strSQL
		
		
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function buscarTelefone(strTelefone)

		m_listBlackList.RemoveAll()
		
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")
		
		strEmpresa = Session("cd_empresa")	
		
		strSQL = _
		" select id_blackList,Ltrim(Rtrim(telefone)) as telefone,convert(char(10),data_alteracao,103) +'  '+CONVERT(char(10),data_alteracao,108) as ultima_alteracao,flg_ativo, usuarios.usuario as nome_usuario "&_
		" from blackList_sms "&_
		" left join usuarios on blackList_sms.cd_empresa = usuarios.cd_empresa "&_
		" and blackList_sms.id_ultimo_usuario_alterou = usuarios.id_usuario "&_
		" where Ltrim(Rtrim(telefone)) = '"&strTelefone&"' "&_
		" and blackList_sms.cd_empresa = '"&strEmpresa&"' "
		
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		if not objretorno.eof then
			set cBlackList = new BlackList
			cBlackList.id_blackList = objretorno("id_blackList")
			cBlackList.telefone = objretorno("telefone")
			cBlackList.data_alteracao = objretorno("ultima_alteracao")
			cBlackList.flg_ativo = objretorno("flg_ativo")
			cBlackList.nome_ultimo_usuario_alterou = objretorno("nome_usuario") 
			
			m_listBlackList.Add aux, cBlackList
			aux = aux + 1
		else 
			set cBlackList = new BlackList
			cBlackList.id_blackList = "N"
			cBlackList.telefone = strTelefone
			cBlackList.data_alteracao = ""
			cBlackList.flg_ativo = "N"
			cBlackList.nome_ultimo_usuario_alterou = ""
			m_listBlackList.Add aux, cBlackList
			aux = aux + 1
		end if 
		
		objretorno.Close
		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		buscarTelefone = m_listBlackList.Items
	
	 end function	
	


end class
%>