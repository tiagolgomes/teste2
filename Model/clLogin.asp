<%
class Login

	private m_id_usuario
	private m_usuario
	private m_senha
	private m_ultimo_acesso
	private m_cod_empresa
	
	private m_listLogin
	
	public property get id_usuario()
		id_usuario = trim(m_id_usuario)
	end property
	public property let id_usuario(pDado)
		m_id_usuario = pDado
	end property
	
	public property get usuario()
		usuario = trim(m_usuario)
	end property
	public property let usuario(pDado)
		m_usuario = pDado
	end property
	
	public property get senha()
		senha = trim(m_senha)
	end property
	public property let senha(pDado)
		m_senha = pDado
	end property
	
	public property get ultimo_acesso()
		ultimo_acesso = trim(m_ultimo_acesso)
	end property
	public property let ultimo_acesso(pDado)
		m_ultimo_acesso = pDado
	end property
	
	public property get cod_empresa()
		cod_empresa = trim(m_cod_empresa)
	end property
	public property let cod_empresa(pDado)
		m_cod_empresa = pDado
	end property
	
	
	'construtor
	sub Class_Initialize()
		m_id_usuario = ""
		m_usuario = ""
		m_senha = ""
		m_ultimo_acesso = ""
		m_cod_empresa = ""
		set m_listLogin = Server.CreateObject("Scripting.Dictionary")
	end sub

	'destrutor
	sub Class_Terminate()
		m_listLogin.RemoveAll()
		set m_listLogin = nothing
	end sub

	public function ValidarUsuario(strUsuario, strSenha,strCdEmpresa)
		m_listLogin.RemoveAll()
		dim conMSSQL, rsTemp, vLogin
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set rsTemp = server.CreateObject("ADODB.Recordset")
		
		strSQL = _
			" select id_usuario,usuario,cd_empresa from usuarios "&_
			" where usuario = '"&SQLInject2(strUsuario)&"'"&_
			" and senha = '"&SQLInject2(strSenha)&"'"		
		rsTemp.open strSQL, conMSSQL, 1, 1
		if not rsTemp.eof then
			m_id_usuario = rsTemp("id_usuario")
			m_usuario = rsTemp("usuario")
			m_cod_empresa = rsTemp("cd_empresa")
		end if 
		rsTemp.close

		set rsTemp = nothing
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function salvarUsuaro(strUsuario, strSenha)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		strSQL = _
			" insert into usuarios values('"&strUsuario&"','"&strSenha&"', getdate())" 
			
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	
	public function carregrUltimosAcessos()
	
		m_listLogin.RemoveAll()
		dim conMSSQL, objretorno, cLogin
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")

		strSQL = _
		" select top 50 id_usuario,usuario, convert(char(10),ultimo_acesso,103) as ultimo_acesso from usuarios order by ultimo_acesso desc "
			
		objretorno.open strSQL, conMSSQL, 1, 1
		while not objretorno.eof
			set cLogin = new Login
			cLogin.id_usuario = objretorno("id_usuario")
			cLogin.usuario = objretorno("usuario")
			cLogin.ultimo_acesso = objretorno("ultimo_acesso")
			m_listLogin.Add objretorno.AbsolutePosition, cLogin
			objretorno.MoveNext()
		wend
		objretorno.close

		set objretorno = nothing
		conMSSQL.close
		set conMSSQL = nothing
		
		carregrUltimosAcessos = m_listLogin.items
	end function 
	
	public function zerarSenha(strIdUsuario)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		strSQL = _
			" update usuarios set senha = 'd41d8cd98f00b204e9800998ecf8427e' where id_usuario = "&strIdUsuario
			
		conMSSQL.execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	
end class
%>