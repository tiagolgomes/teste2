<%
validaSessao()

function validaSessao()
	Response.EXPIRES=0
    Response.Expires = 60
    Response.Expiresabsolute = Now() - 1
    Response.AddHeader "pragma","no-cache"
    Response.AddHeader "cache-control","private"
    
	if trim(session("cd_empresa")) = "" then
	    Response.Write "<script>window.open('../view/index.html','_self');</script>"
	    Response.end 
    end if
end function

%>