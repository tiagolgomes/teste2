<%
class Parametros

	private m_email_envio
	private m_servidor_envio
	private m_qtd_tentativas_envio_campanha 
	private m_intervalo_espera_thread
	private m_qtd_limite_envio_canal
	
	private m_listParametros
	
	public property get email_envio()
		email_envio = trim(m_email_envio)
	end property
	public property let email_envio(pDado)
		m_email_envio = pDado
	end property
	
	public property get servidor_envio()
		servidor_envio = trim(m_servidor_envio)
	end property
	public property let servidor_envio(pDado)
		m_servidor_envio = pDado
	end property
	
	public property get qtd_tentativas_envio_campanha()
		qtd_tentativas_envio_campanha = trim(m_qtd_tentativas_envio_campanha)
	end property
	public property let qtd_tentativas_envio_campanha(pDado)
		m_qtd_tentativas_envio_campanha = pDado
	end property
	
	public property get intervalo_espera_thread()
		intervalo_espera_thread = trim(m_intervalo_espera_thread)
	end property
	public property let intervalo_espera_thread(pDado)
		m_intervalo_espera_thread = pDado
	end property
	
	
	public property get qtd_limite_envio_canal()
		qtd_limite_envio_canal = trim(m_qtd_limite_envio_canal)
	end property
	public property let qtd_limite_envio_canal(pDado)
		m_qtd_limite_envio_canal = pDado
	end property
	
	sub Class_Initialize()
		m_email_envio = ""
		m_servidor_envio = ""
		m_qtd_tentativas_envio_campanha = ""
		m_intervalo_espera_thread = ""
		m_qtd_limite_envio_canal = ""
		set m_listParametros = Server.CreateObject("Scripting.Dictionary")
	end sub

	sub Class_Terminate()
		m_listParametros.RemoveAll()
		set m_listParametros = nothing
	end sub
	
	public function GravarParametros(stremail_envio, strservidor_envio,strQtdTentativasCampanha,strIntervaloThread,strLimiteEnvioCanal)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		strSQL = _
		" select de_email_envio from parametros "
		objretorno.Open strSQL, conMSSQL	
		if not objretorno.eof then
			strSQL = _
			" update parametros set qtd_tentativas_envio_campanha = "&strQtdTentativasCampanha&",intervalo_espera_thread = "&strIntervaloThread&",qtd_limite_envio_canal = "&strLimiteEnvioCanal&" "
			conMSSQL.Execute strSQL
		else
			strSQL = _
			" insert into parametros values('"&stremail_envio&"','"&strservidor_envio&"','"&stremail_envio&"',"&strQtdTentativasCampanha&","&strIntervaloThread&","&strLimiteEnvioCanal&")" 
			conMSSQL.Execute strSQL
		end if 		
		
		objretorno.Close
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	 public function BuscarParametros()
		
		 dim conMSSQL, objretorno
		 set conMSSQL = Server.CreateObject("ADODB.Connection")
		 conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		 conMSSQL.Open
		 set objretorno = server.CreateObject("ADODB.Recordset")		
		
		 strSQL = _
		 " select de_email_envio, de_servidor_envio,qtd_tentativas_envio_campanha, intervalo_espera_thread,qtd_limite_envio_canal from parametros "
		 objretorno.Open strSQL, conMSSQL	
		 if not objretorno.eof then 
			 m_email_envio = objretorno("de_email_envio")
			 m_servidor_envio = objretorno("de_servidor_envio")
			 m_qtd_tentativas_envio_campanha = objretorno("qtd_tentativas_envio_campanha")
			 m_intervalo_espera_thread = objretorno("intervalo_espera_thread")
			 m_qtd_limite_envio_canal = objretorno("qtd_limite_envio_canal")
		end if 	
		objretorno.Close
		
		 set objretorno = nothing
		 conMSSQL.close
		 set conMSSQL = nothing		
	 end function	
	


end class
%>