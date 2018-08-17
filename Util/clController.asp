<%
class Controller
	
	public function zerarEnviosMensal()
		dim conMSSQL, objretorno
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open
		set objretorno = server.CreateObject("ADODB.Recordset")	
		
		
		strSQL = _		
		" select DAY(GETDATE()) as primeiro_dia"
		objretorno.Open strSQL, conMSSQL	
		if not objretorno.eof then
			strDia = objretorno("primeiro_dia")
		end if 
		objretorno.Close
	
		if(strDia = "1" or strDia = "01") then 
		
			strSQL = _		
			" select data_mensal from controle_limite "&_
			" where CONVERT(VARCHAR, GETDATE() - DAY(GETDATE()) + 1, 103) = convert(char(10),data_mensal,103) "
			objretorno.Open strSQL, conMSSQL	
			if objretorno.eof then
				strSQL = _
				" update empresas set qt_envio_mensal = 0 "
				conMSSQL.Execute strSQL
				
				strSQL = _
				" insert into controle_limite values(getdate()) "
				conMSSQL.Execute strSQL
			end if 
			objretorno.Close
		
		end if 
		conMSSQL.close
		set conMSSQL = nothing
	end function 
end class
%>