<%
class Email

	private m_email_envio
	private m_servidor_envio
	private m_email_from 
	private m_corpo_email
	private m_numero_retorno
	private m_mensagem_retorno
	private m_email_retorno
	private m_email_texo
	private m_nu_contrato
	private m_listEmails
	
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
			
	public property get email_from()
	email_from = trim(m_email_from)
	end property
	public property let email_from(pDado)
		m_email_from = pDado
	end property
	
	public property get corpo_email()
		corpo_email = trim(m_corpo_email)
	end property
	public property let corpo_email(pDado)
		m_corpo_email= pDado
	end property
	
	public property get numero_retorno()
		numero_retorno = trim(m_numero_retorno)
	end property
	public property let numero_retorno(pDado)
		m_numero_retorno = pDado
	end property
	
	public property get mensagem_retorno()
		mensagem_retorno = trim(m_mensagem_retorno)
	end property
	public property let mensagem_retorno(pDado)
		m_mensagem_retorno = pDado
	end property
	
	public property get email_retorno()
		email_retorno = trim(m_email_retorno)
	end property
	public property let email_retorno(pDado)
		m_email_retorno = pDado
	end property
	
	public property get email_texo()
		email_texo = trim(m_email_texo)
	end property
	public property let email_texo(pDado)
		m_email_texo = pDado
	end property
	
	public property get nu_contrato()
		nu_contrato = trim(m_nu_contrato)
	end property
	public property let nu_contrato(pDado)
		m_nu_contrato = pDado
	end property
	
	
	sub Class_Initialize()
		m_email_envio = ""
		m_servidor_envio = ""
		m_email_from = ""
		m_corpo_email = ""
		m_numero_retorno = ""
		m_mensagem_retorno = ""
		m_email_retorno = ""
		m_email_texo = ""
		m_nu_contrato = ""
		set m_listEmails = Server.CreateObject("Scripting.Dictionary")
	end sub
	
	sub Class_Terminate()
		m_listEmails.RemoveAll()
		set m_listEmails = nothing
	end sub
	
	public function EnviarEmail(corpo,assunto,email_from,email_to)
	
	Set objEmail = CreateObject("CDO.Message")

	objEmail.From = email_from
	objEmail.Subject = assunto 
	objEmail.To = email_to
	objEmail.CC = "andre@hdr.adv.br;mauricio@hdr.adv.br"
	
	
	objEmail.htmlbody = "<html><head></head><body>"&corpo&"</body></html>"
	
	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.terra.com.br"

	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = "mextit@terra.com.br"

	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = "mextit@terra.com.br"

	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "s8c1j4s0"

	objEmail.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587

	objEmail.Configuration.Fields.Update
	

	
	if(trim(objEmail.To) <> "") then 
		objEmail.Send	
	end if 
	
	end function
	
	public function BuscaCampanhaEnvioRetorno()
		m_listEmails.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		set objretorno2 = server.CreateObject("ADODB.Recordset")			
		
		strSQL = _
		" select top 150 SUBSTRING(numero, 4, 15) as numero,mensagem from sms_retornos with (NOLOCK) "&_
		" where fl_email_retorno_enviado = '0' "&_
		" and SUBSTRING(numero, 4, 15) <> '' "&_
		" and SUBSTRING(numero, 4, 15) <> '0' "&_
		" and Len(SUBSTRING(numero, 4, 15)) > 6 "&_
		" order by cd_retorno desc "
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			strSQL = " select top 1 telefoneXcontratos.nu_contrato,campanhas.email_retorno,campanhas_envio.texto_envio as texto from campanhas_envio with (NOLOCK) "&_
			" inner join campanhas with (NOLOCK) on campanhas.id_campanha = campanhas_envio.id_campanha "&_
			" left join telefoneXcontratos with (NOLOCK) on telefoneXcontratos.telefone = campanhas_envio.telefone "&_
			" where campanhas_envio.telefone = '"&objretorno("numero")&"' "&_
			" and campanhas.fl_retorno = 0 "&_
			" order by campanhas_envio.id_campanha  desc "
			objretorno2.Open strSQL, conMSSQL	
			while not objretorno2.eof
				set cEmail = new Email
				cEmail.numero_retorno = objretorno("numero")
				cEmail.mensagem_retorno = objretorno("mensagem")
				cEmail.email_retorno = objretorno2("email_retorno")
								
				cEmail.email_texo = objretorno2("texto")
				cEmail.nu_contrato = objretorno2("nu_contrato")
				m_listEmails.Add aux, cEmail
				aux = aux + 1
				objretorno2.MoveNext()
			wend
			objretorno2.Close
			
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		set objretorno = nothing
		set objretorno2 = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		BuscaCampanhaEnvioRetorno = m_listEmails.Items
	end function 
	
	public function AtualizarEnvio(strNumero)
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		strSQL = _
			"update sms_retornos set fl_email_retorno_enviado = '-1' "&_
			" where SUBSTRING(numero, 4, 15) = '"&strNumero&"'"
		conMSSQL.Execute strSQL
		conMSSQL.close
		set conMSSQL = nothing
	end function 
	
	public function BuscaCampanhaEnvioRetornoTeste()
		m_listEmails.RemoveAll()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		
		set objretorno = server.CreateObject("ADODB.Recordset")	
		set objretorno2 = server.CreateObject("ADODB.Recordset")			
		
		strSQL = _
		" select top 100 SUBSTRING(numero, 4, 15) as numero,mensagem from sms_retornos "&_
		" where fl_email_retorno_enviado = '0' "&_
		" and SUBSTRING(numero, 4, 15) <> '' "&_
		" and SUBSTRING(numero, 4, 15) <> '0' "&_
		" and Len(SUBSTRING(numero, 4, 15)) > 6 "&_
		" order by cd_retorno desc "
		aux = 0
		objretorno.Open strSQL, conMSSQL	
		while not objretorno.eof
			strSQL = " select top 1 telefoneXcontratos.nu_contrato,* from campanha_envio "&_
			" inner join campanhas on campanhas.id_campanha = campanha_envio.id_campanha "&_
			" left join telefoneXcontratos on telefoneXcontratos.telefone = campanha_envio.telefone "&_
			" where campanha_envio.telefone = '"&objretorno("numero")&"' "&_
			" order by campanha_envio.id_campanha  desc "
		
			objretorno2.Open strSQL, conMSSQL	
			while not objretorno2.eof
				set cEmail = new Email
				cEmail.numero_retorno = objretorno("numero")
				cEmail.mensagem_retorno = objretorno("mensagem")
				cEmail.email_retorno = objretorno2("email_retorno")
				cEmail.email_texo = objretorno2("texto")
				m_listEmails.Add aux, cEmail
				aux = aux + 1
				objretorno2.MoveNext()
			wend
			objretorno2.Close
			
			objretorno.MoveNext()
		wend	
			
		objretorno.Close
		set objretorno = nothing
		set objretorno2 = nothing
		conMSSQL.close
		set conMSSQL = nothing		
		BuscaCampanhaEnvioRetornoTeste = m_listEmails.Items
	end function 
	
	
	
	
end class
%>