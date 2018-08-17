<%

Sub DebugPrint(arrParam)
	strValue = ""
	for each p in arrParam 
		strValue = strValue & " - "  & p
	next
	
	Response.Write strValue 
End Sub

Sub Mensagem(strTexto)
	'Caracteres
	' chr(60) = <
	' chr(62) = >
	' chr(34) = ''
	' chr(47) = /
	' aqui a fun��o tenta montar um alert com a mensagem que o progama que chamou passou
	' Existe um limite de 255 caracteres para strings em javascript()
	' Response.Write(< script > alert( & ' strTexto & ' & ); <  / script >)
    Response.Write(chr(60) & "script" & chr(62) & "alert(" & chr(34) & _
    replace(strTexto,"/",chr(37)) & chr(34) & ");" & chr(60) & chr(47) & "script" & chr(62))
End Sub

Sub redireciona(strLink)
	'Caracteres
	' chr(60) = <
	' chr(62) = >
	' chr(34) = ''
	' chr(47) = /
    Response.Write(chr(60) & "script" & chr(62) & "location.href = " & chr(34) & _
    strLink & chr(34) & ";" & chr(60) & chr(47) & "script" & chr(62))
End Sub

Sub redireciona_target(strLink, pTarget)
	'Caracteres
	' chr(60) = <
	' chr(62) = >
	' chr(34) = ''
	' chr(47) = /
    Response.Write(chr(60) & "script" & chr(62) & "window.open(" & chr(34) & _
    strLink & chr(34) & "," & chr(34) & pTarget & chr(34) & ");" & chr(60) & chr(47) & "script" & chr(62))
End Sub


Function Vlimpa(strcampo)
  if not isnull(strcampo) then
    strcampo = replace(strcampo, "�", "A")
    strcampo = replace(strcampo, "�", "A")
    strcampo = replace(strcampo, "�", "A")
    strcampo = replace(strcampo, "�", "A")
    strcampo = replace(strcampo, "�", "E")
    strcampo = replace(strcampo, "�", "E")
    strcampo = replace(strcampo, "�", "E")
    strcampo = replace(strcampo, "�", "I")
    strcampo = replace(strcampo, "�", "I")
    strcampo = replace(strcampo, "�", "I")
    strcampo = replace(strcampo, "�", "O")
    strcampo = replace(strcampo, "�", "O")
    strcampo = replace(strcampo, "�", "O")
    strcampo = replace(strcampo, "�", "O")
    strcampo = replace(strcampo, "�", "U")
    strcampo = replace(strcampo, "�", "U")
    strcampo = replace(strcampo, "�", "U")
    strcampo = replace(strcampo, "�", "U")
    strcampo = replace(strcampo, "�", "C")
    strcampo = replace(strcampo, "'", " ")
    strcampo = replace(strcampo, "�", "a")
    strcampo = replace(strcampo, "�", "a")
    strcampo = replace(strcampo, "�", "a")
    strcampo = replace(strcampo, "�", "a")
    strcampo = replace(strcampo, "�", "e")
    strcampo = replace(strcampo, "�", "e")
    strcampo = replace(strcampo, "�", "e")
    strcampo = replace(strcampo, "�", "i")
    strcampo = replace(strcampo, "�", "i")
    strcampo = replace(strcampo, "�", "i")
    strcampo = replace(strcampo, "�", "o")
    strcampo = replace(strcampo, "�", "o")
    strcampo = replace(strcampo, "�", "o")
    strcampo = replace(strcampo, "�", "o")
    strcampo = replace(strcampo, "�", "u")
    strcampo = replace(strcampo, "�", "u")
    strcampo = replace(strcampo, "�", "u")
    strcampo = replace(strcampo, "�", "u")
    strcampo = replace(strcampo, "�", "c")
    strcampo = replace(strcampo, "�", " ")
    strcampo = replace(strcampo, "�", " ")
    vlimpa   = strcampo
  end if
End Function



function removeAcentos(str)
	if not IsNull(str) then
	
		response.write str
		str = Replace(str, "�", "a")
		str = Replace(str, "�", "a")
		str = Replace(str, "�", "a")
		str = Replace(str, "�", "a")
		str = Replace(str, "�", "a")	
		str = Replace(str, "�", "e")
		str = Replace(str, "�", "e")
		str = Replace(str, "�", "e")	
		str = Replace(str, "�", "e")	
		str = Replace(str, "�", "i")
		str = Replace(str, "�", "i")
		str = Replace(str, "�", "i")
		str = Replace(str, "�", "i")
		str = Replace(str, "�", "o")
		str = Replace(str, "�", "o")
		str = Replace(str, "�", "o")
		str = Replace(str, "�", "o")		
		str = Replace(str, "�", "o")	
		str = Replace(str, "�", "u")
		str = Replace(str, "�", "u")
		str = Replace(str, "�", "u")	
		str = Replace(str, "�", "u")	
		str = Replace(str, "�", "c")
		str = Replace(str, "�", "A")
		str = Replace(str, "�", "A")
		str = Replace(str, "�", "A")
		str = Replace(str, "�", "A")
		str = Replace(str, "�", "A")	
		str = Replace(str, "�", "E")
		str = Replace(str, "�", "E")
		str = Replace(str, "�", "E")	
		str = Replace(str, "�", "E")	
		str = Replace(str, "�", "I")
		str = Replace(str, "�", "I")
		str = Replace(str, "�", "I")
		str = Replace(str, "�", "I")
		str = Replace(str, "�", "O")
		str = Replace(str, "�", "O")
		str = Replace(str, "�", "O")
		str = Replace(str, "�", "O")		
		str = Replace(str, "�", "O")	
		str = Replace(str, "�", "U")
		str = Replace(str, "�", "U")
		str = Replace(str, "�", "U")	
		str = Replace(str, "�", "U")	
		str = Replace(str, "�", "C")
		removeAcentos = str
	end if
end function

function ignoraAcentos(str)
	str = trim(str)
	dim i, strCh, strSaida
	for i = 1 to len(str)
		strCh = mid(str,i,1)
		select case strCh
			case "a", "�", "�", "�", "�", "�", "A", "�", "�", "�", "�", "�"
				strSaida = strSaida & "[a�����]" 
				
			case "e", "�", "�", "�", "�", "E", "�", "�", "�", "�"    
				strSaida = strSaida & "[e����]"
				 
			case "i", "�", "�", "�", "�", "I", "�", "�", "�", "�"
				strSaida = strSaida & "[i����]" 
				
			case "o", "�", "�", "�", "�", "�", "O", "�", "�", "�", "�", "�"
				strSaida = strSaida & "[o�����]" 
				
			case "u", "�", "�", "�", "�", "U", "�", "�", "�", "�"   
				strSaida = strSaida & "[u����]" 
				
			case "�", "c", "�", "C"
				strSaida = strSaida & "[c�]" 
			
			'caso o nome do aluno tenha " ' " ele acrescenta mais uma aspas simples para que na hora de fazer a busca no banco n�o d� erro'	
			case "'" 
				strSaida = strSaida & "''"
				
			case else
				strSaida = strSaida & strCh  
		end select

	next 
	ignoraAcentos = trim(strSaida)
end function



function completaComZeros(pNum, pTamanho)
	dim i, lTam, result
	result = ""
	if not(IsNull(pNum)) then
		lTam = Len(pNum)+1
	else
		lTam = 1
	end if
	for i = lTam to pTamanho
		result = result & "0"
	next
	if not(IsNull(pNum)) then result = result & pNum
	
	completaComZeros = result
end function

function completaComEspaco(pString, pTamanho)
	dim i, lTam, result
	result = ""
	if not(IsNull(pString)) then
		lTam = Len(pString)+1
	else
		lTam = 1
	end if
	for i = lTam to pTamanho
		result = result & " "
	next
	if not(IsNull(pString)) then result = pString & result 
	
	completaComEspaco = result
end function

function completaComEspacoAntes(pString, pTamanho)
	dim i, lTam, result
	result = ""
	if not(IsNull(pString)) then
		lTam = Len(pString)+1
	else
		lTam = 1
	end if
	for i = lTam to pTamanho
		result = result & " "
	next
	if not(IsNull(pString)) then result = result & pString
	
	completaComEspacoAntes = result
end function

function monta_cnpj(cnpj)
	if len(cnpj) <> 14 then
	  strcnpj = cnpj
	else
	  strcnpj = mid(cnpj,1,2) & "." & mid(cnpj,3,3) & "." & mid(cnpj,6,3) & "/" &_
	  mid(cnpj,9,4) & "-" & mid(cnpj,13,2)
	end if
	monta_cnpj = strcnpj
end function

function monta_cpf(cpf)
	if len(cpf) <> 11 then
		strcpf = cpf
	else
		strcpf = mid(cpf,1,3) & "." & mid(cpf,4,3) & "." &  mid(cpf,7,3) & "-" &  mid(cpf,10,2)
	end if
	monta_cpf = strcpf
end function

function monta_cep(cep)
	if len(cep) <> 8 then
		strcep = cep
	else
		strcep = mid(cep,1,5) & "-" & mid(cep,6,3)
	end if
	monta_cep = strcep
end function

function monta_cep_completo(cep)
	if len(cep) <> 8 then
		strcep = cep
	else
		strcep = mid(cep,1,2)& "." &mid(cep,3,3) & "-" & mid(cep,6,3)
	end if
	monta_cep_completo = strcep
end function



'=========== validar_email ===========

'Funcao que verifica se um e-mail e valido.
'Respostas:
' 0 -> se � nulo ou vazio
' 1 -> Se o e-mail � inv�lido
' 2 -> Se o e-mail � v�lido

function validar_email(email)

	dim con
    validar_email = 2
    email = trim(email)

	if trim(email) = "" then
	    validar_email = 0
	    exit function
	end if
	
	if IsNull(email) then
	    validar_email = 0
	    exit function
	end if	

	email = lcase(email)

	if (left(email, 1) = "." or left(email, 1) = "@") then
	    validar_email = 1
	    exit function
	end if

	if (right(email, 1) = "." or right(email, 1) = "_" or _
	    right(email, 1) = "@") then
	    validar_email = 1
	    exit function
	end if

	con = 0

	for i = 1 to len(email)
	    if (mid(email, i, 1)) = "@" then
	       con = con + 1
	    end if
	next

	if con > 1 or con = 0 then
	    validar_email = 1
	    exit function
	end if

	con = 0

	for i = 1 to len(email)
	    if (mid(email, i, 1)) = "@" then
	       for j = i to len(email)
	          if (mid(email, j, 1)) = "." then
	             con = con + 1
	          end if
	       next
	    end if
	next

	for i = 1 to len(email)
	    if (mid(email, i, 1)) = "/" then
	       for j = i to len(email)
	          if (mid(email, j, 1)) = "." then
	             con = con + 1
	          end if
	       next
	    end if
	next
	if con < 1 then
	    validar_email = 1
	    exit function
	end if

	con = 0
	
	for i = 1 to len(email)
	    if (mid(email, i, 1) = ".") or (mid(email, i, 1) = "@") then
	       if (mid(email, i + 1, 1) = ".") or (mid(email, i + 1, 1) = "@") then
	          validar_email = 1
	          exit function
	       end if
	    end if
	next
	
	for i = 1 to len(email)
	    if (mid(email, i, 1) < "a" or mid(email, i, 1) > "z") and _
	       (mid(email, i, 1) < "0" or mid(email, i, 1) > "9") and _
	       mid(email, i, 1) <> "." and mid(email, i, 1) <> "_" and _
	       mid(email, i, 1) <> "@" and mid(email, i, 1) <> "-" then
	          validar_email = 1
	          exit function
	    end if
	next
	
	'if validar_email <> 2 then mensagem " formato inv�lido de e-mail" & email
	'validar_email = 2
	
end function


function UpperPrimeiraPalavra(pStr)	
	dim strTemp	
	pStr = trim(pStr)
	
	if (trim(pStr) <> "") then
		pos = InStr(1, pStr, " ")
		if pos = 0 then pos = Len(pStr)
		strTemp = UCase(Mid(pStr, 1, pos))
		strTemp = strTemp + Mid(pStr, pos+1, len(pStr))
	end if
	
	UpperPrimeiraPalavra = strTemp
end function



Public Function StrRepeat(strExpression, intTimes)
    Dim j, strData
    For j = 1 To CLng(intTimes)
        strData = strData & strExpression
    Next
    StrRepeat = strData
End Function


'-----------------------------------------------------   
'Funcao: IsCNPJ(ByVal intNumero)   
'Sinopse: Verifica se o valor passado � um CNPJ v�lido    
'Parametro: intNumero   
'Retorno: Booleano   
'-----------------------------------------------------   
Function IsCNPJ(ByVal intNumero)   
    'Validando o formato do CNPJ com express�o regular   
    Set regEx = New RegExp                                'Cria o Objeto Express�o   
    regEx.Pattern = "d{2}.?d{3}.?d{3}/?d{4}-?d{2}"        ' Express�o Regular   
    regEx.IgnoreCase = True                               ' Sensitivo ou n�o   
    regEx.Global = True   
    Retorno = True
    Dim CNPJ_t   
    CNPJ_t  = intNumero   
    CNPJ_t  = Replace(CNPJ_t, ".", "")   
    CNPJ_t  = Replace(CNPJ_t, "/", "")   
    CNPJ_t  = Replace(CNPJ_t, "-", "")   
    
    if len(trim(CNPJ_t)) <> 14 then Retorno = False  
            
    Set regEx = Nothing  
       
    'Caso seja verdadeiro posso validar se o CNPJ � v�lido   
    If Retorno = True Then  
        'Validando a sequencia n�meros   
        Dim CNPJ_temp   
        CNPJ_temp            = intNumero   
        CNPJ_temp            = Replace(CNPJ_temp, ".", "")   
        CNPJ_temp            = Replace(CNPJ_temp, "/", "")   
        CNPJ_temp            = Replace(CNPJ_temp, "-", "")   
        CNPJ_Digito_temp    = Right(CNPJ_temp, 2)   
           
        'Somando os 12 primeiros digitos do CNPJ    
        Soma    = (Clng(Mid(CNPJ_temp,1,1)) * 5) + (Clng(Mid(CNPJ_temp,2,1)) * 4) + (Clng(Mid(CNPJ_temp,3,1)) * 3) + (Clng(Mid(CNPJ_temp,4,1)) * 2) + (Clng(Mid(CNPJ_temp,5,1)) * 9) + (Clng(Mid(CNPJ_temp,6,1)) * 8)+ (Clng(Mid(CNPJ_temp,7,1)) * 7) + (Clng(Mid(CNPJ_temp,8,1)) * 6) + (Clng(Mid(CNPJ_temp,9,1)) * 5) + (Clng(Mid(CNPJ_temp,10,1)) * 4) + (Clng(Mid(CNPJ_temp,11,1)) * 3) + (Clng(Mid(CNPJ_temp,12,1)) * 2)   
        '----------------------------------   
        'Calculando o 1� d�gito verificador   
        '----------------------------------   
        'Pegando o resto da divis�o por 11   
        Resto    = (Soma Mod 11)   
        If Resto < 2 Then  
            DigitoHum = 0   
        Else  
            DigitoHum = Cstr(11-Resto)   
        End If  
        '----------------------------------   
        '----------------------------------   
        'Calculando o 2� d�gito verificador   
        '----------------------------------   
        'Somando os 12 primeiros digitos do CNPJ mais o 1� d�gito   
        Soma    = (Clng(Mid(CNPJ_temp,1,1)) * 6) + (Clng(Mid(CNPJ_temp,2,1)) * 5) + (Clng(Mid(CNPJ_temp,3,1)) * 4) + (Clng(Mid(CNPJ_temp,4,1)) * 3) + (Clng(Mid(CNPJ_temp,5,1)) * 2) + (Clng(Mid(CNPJ_temp,6,1)) * 9) + (Clng(Mid(CNPJ_temp,7,1)) * 8) + (Clng(Mid(CNPJ_temp,8,1)) * 7) + (Clng(Mid(CNPJ_temp,9,1)) * 6) + (Clng(Mid(CNPJ_temp,10,1)) * 5) + (Clng(Mid(CNPJ_temp,11,1)) * 4) + (Clng(Mid(CNPJ_temp,12,1)) * 3) + (DigitoHum * 2)   
        'Pegando o resto da divis�o por 11   
        Resto    = (Soma Mod 11)   
           
        If Resto < 2 Then  
            DigitoDois = 0   
        Else  
            DigitoDois = Cstr(11-Resto)   
        End If  
        '----------------------------------   
        'Verificando se os digitos s�o iguais aos dig�tados.   
        DigitoCNPJ = Cstr(DigitoHum) & Cstr(DigitoDois)   
        If Cstr(CNPJ_Digito_temp) = Cstr(DigitoCNPJ) Then  
            Retorno = True  
        Else  
            Retorno = False  
        End If  
    End If  
    IsCNPJ = Retorno   
End Function  


'busca nome do arquivo em execu��o
function BuscaNomeScript()
	dim scriptName, pos
	
	scriptName = Request.ServerVariables("SCRIPT_NAME")
	pos = InStr(1, scriptName, "/")
	while pos > 0
		scriptName = mid(scriptName,pos+1,len(scriptName))
		pos = InStr(1, scriptName, "/")
	wend
	
	BuscaNomeScript = scriptName
end function


function BuscaExtensao(pNomeArq)
	
	pos = InStr(1, pNomeArq, ".")
	while pos > 0
		pNomeArq = mid(pNomeArq,pos+1,len(pNomeArq))
		pos = InStr(1, pNomeArq, ".")
	wend
	
	BuscaExtensao = pNomeArq
end function


function BuscaNomeArquivo(path)
	
	pos = InStr(1, path, "/")
	while pos > 0
		path = mid(path,pos+1,len(path))
		pos = InStr(1, path, "/")
	wend
	
	BuscaNomeArquivo = path
end function



'Fun��o que corrige o bug no Round do asp
Function MyRound(number,decPoints)
	corretor = 0
	if decPoints > 0 then
		corretor = "0,"
		for i = 1 to decPoints
			corretor = corretor & "0"
		next
		corretor = corretor & "5"
	end if
	
	decPoints = 10^decPoints
	MyRound = round(number*decPoints+corretor)/decPoints
End Function

Function CalculaCPF(Numero_CPF)
	Dim RecebeCPF, Numero(11), Soma, Resultado1, Resultado2
	Dim Vb_Valido, Vs_String, X, CH

	Vb_Valido = True
	RecebeCPF = Numero_CPF

	'Retirar todos os caracteres que nao sejam 0-9
	Vs_String = ""
	For X = 1 to Len(RecebeCPF)
	    Ch=Mid(RecebeCPF,X,1)
	    If Asc(Ch)>=48 And Asc(Ch)<=57 Then
	       Vs_String = Vs_String & Ch
	    End If
	Next

	RecebeCPF = Vs_String

	If Len(RecebeCPF) <> 11 Then
	   Vb_Valido =  false
	ElseIF RecebeCPF = "00000000000" or RecebeCPF = "11111111111" or RecebeCPF = "22222222222"  or _
	       RecebeCPF = "33333333333" or RecebeCPF = "44444444444" or RecebeCPF = "55555555555"  or _
	       RecebeCPF = "66666666666" or RecebeCPF = "77777777777" or RecebeCPF = "88888888888"  or _
	       RecebeCPF = "99999999999" then
	    Vb_Valido = false
	Else

	Numero(1) = Cint(Mid(RecebeCPF,1,1))
	Numero(2) = Cint(Mid(RecebeCPF,2,1))
	Numero(3) = Cint(Mid(RecebeCPF,3,1))
	Numero(4) = Cint(Mid(RecebeCPF,4,1))
	Numero(5) = Cint(Mid(RecebeCPF,5,1))
	Numero(6) = CInt(Mid(RecebeCPF,6,1))
	Numero(7) = Cint(Mid(RecebeCPF,7,1))
	Numero(8) = Cint(Mid(RecebeCPF,8,1))
	Numero(9) = Cint(Mid(RecebeCPF,9,1))
	Numero(10) = Cint(Mid(RecebeCPF,10,1))
	Numero(11) = Cint(Mid(RecebeCPF,11,1))

	Soma = 10 * Numero(1) + 9 * Numero(2) + 8 * Numero(3) + 7 * Numero(4) + 6 * Numero(5) + 5 * Numero(6) + 4 * Numero(7) + 3 * Numero(8) + 2 * Numero(9)
	Soma = Soma - (11 * (Int(Soma / 11)))

	IF Soma = 0 or Soma = 1 Then
	Resultado1 = 0
	Else
	Resultado1 = 11 - soma
	End IF

	IF Resultado1 = Numero(10) Then
	   Soma = Numero(1) * 11 + Numero(2) * 10 + Numero(3) * 9 + Numero(4) * 8 + Numero(5) * 7 + Numero(6) * 6 + Numero(7) * 5 + Numero(8) * 4 + Numero(9) * 3 + Numero(10) * 2
	   Soma = Soma -(11 * (Int(Soma / 11)))

	   IF Soma = 0 or Soma = 1 Then
	      Resultado2 = 0
	   Else
	      Resultado2 = 11 - Soma
	   End IF

	   IF Resultado2 = Numero(11) Then
	      Vb_Valido = True
	   Else
	      Vb_Valido = False
	   End IF
	Else
	   Vb_Valido = False
	End IF
	End IF

	CalculaCPF = Vb_Valido
	'CalculaCPF = Resultado1  & Resultado2 & Vb_Valido
End Function



function StringToArray(pStrList)
	dim lista(), i
	dim strList, item
	strList = trim(pStrList) 
	pItem = trim(pItem)
	StringToArray = ""
	
	i = 0	
	do while strList <> ""
		pos = InStr(1, strList, "|")
		if pos = 0 then 
			item = strList
			strList = ""
		else
			item = mid(strList,1, pos-1)
			strList = mid(strList,pos+1, len(strList))
		end if
		
		i = i + 1
		Redim Preserve lista(i)
		lista(i-1) = item
		StringToArray = lista
	loop
end function

function formataNota(pNota, pDecimais)
	formataNota = pNota
	
	if not IsNull(pNota) then
		pNota = replace(pNota,".",",")
		if pNota <> "SN" and _
			pNota <> "" and _
			pNota <> "-1" and _
			IsNumeric(pNota) then
			formataNota = FormatNumber(pNota,pDecimais)
		end if
	end if
end function



'Fun��o que verifica se o campo esta vazio ou nulo
' Retorno  booleano
Function Vazio(strCampo, nome)
	Vazio = true
	if strCampo = "" or isnull(strCampo)then
		mensagem " O campo " & nome & " est� vazio."
		Vazio = false
	end if
end Function


'Fun��o que verifica se o campo � numerico
' Retorno  booleano
Function Numerico(strCampoNumerico, nome1)
	Numerico = true
	if Vazio(strCampoNumerico, nome1) then
		if not isnumeric(strCampoNumerico) then	
			mensagem " O campo " & nome1 & " � num�rico."
			Numerico = false
		end if
	end if
end Function

Function QueryDB (qry)
	dim conMSSQL, rsTemp
	set conMSSQL = Server.CreateObject("ADODB.Connection")
	conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
	conMSSQL.Open
	

	if Instr (qry,"update") <> 0 or Instr (qry,"insert") <> 0 or Instr (qry,"delete") <> 0 then
		conMSSQL.Execute qry
	else
		set rsTemp = server.CreateObject("ADODB.Recordset")
		set lista = Server.CreateObject("Scripting.Dictionary")
		
		rsTemp.open qry, conMSSQL, 1, 1
		while not rsTemp.eof
			set vObj = Server.CreateObject("Scripting.Dictionary")		
			for each x in rsTemp.Fields
				vObj.Add x.name,x.value
			next
			lista.Add rsTemp.AbsolutePosition, vObj
			rsTemp.MoveNext()
			set vObj = nothing
		wend
		rsTemp.close
		set rsTemp = nothing
		QueryDB = lista.Items		
	end if
		
	conMSSQL.close
	set conMSSQL = nothing
End Function

function Ceil( Number )
  Ceil = Int( Number )
  if Ceil <> Number then
    Ceil = Ceil + 1
  end if
end function

function Floor( Number )
    Floor = Int( Number )
end function

function NomeProgramaAsp()
	dim strPrograma
	
	strPrograma = ""
	i = Len(Trim(Request.ServerVariables("PATH_INFO")))
	
	while i > 0
		strPrograma = Mid(Trim(Request.ServerVariables("PATH_INFO")),i,1) & strPrograma
		Request.ServerVariables("PATH_INFO")
		i = i - 1
		
		if Mid(Trim(Request.ServerVariables("PATH_INFO")),i,1) = "/" then i = 0
	wend
	
	NomeProgramaAsp = strPrograma
end function

function ExisteTituloPorNomePrograma(pNomePrograma, pTitulo)
	ExisteTituloPorNomePrograma = pTitulo
	if pNomePrograma <> "" then
		dim conMSSQL, rsTemp, cmdSQL
		set conMSSQL = Server.CreateObject("ADODB.Connection")
		conMSSQL.ConnectionString = Application("conMSSQL_ConnectionString")
		conMSSQL.Open

		set cmdSQL = Server.CreateObject("ADODB.Command")
		cmdSQL.ActiveConnection = conMSSQL
		cmdSQL.Parameters.Append cmdSQL.CreateParameter("@programa", adVarChar, adParamInput, 60)
		cmdSQL.Parameters("@programa") = pNomePrograma

		cmdSQL.CommandText = _
		" select menu_programa.id_programa, menu_titulo.titulo " &_
		" from menu_titulo with(nolock)  " &_
		" inner join menu_programa  with(nolock) on menu_titulo.programa = menu_programa.programa " &_
		" where menu_programa.programa = ? "	
		set rsTemp = cmdSQL.Execute()
		if not rsTemp.eof then 
			ExisteTituloPorNomePrograma = UCase(rsTemp("titulo"))
		end if
		rsTemp.close
		set rsTemp = nothing
		set cmdSQL = nothing
		conMSSQL.close
		set conMSSQL = nothing
	end if
end function

Function BooleanToBit(pDados)
	if pDados then 
		BooleanToBit = 1
	else
		BooleanToBit = 0
	end if
End Function

Function UndefinedToBranco(pDados)
	UndefinedToBranco = pDados
	if pDados = "undefined" or pDados = "null" or pDados = null then UndefinedToBranco = ""
End Function

function trataCharPost(auxUsuario)
	
	auxUsuario = replace(auxUsuario,"á"		,"�")
		auxUsuario = replace(auxUsuario,"� "		,"�")
		auxUsuario = replace(auxUsuario,"ã"		,"�")
		auxUsuario = replace(auxUsuario,"â"		,"�")
		auxUsuario = replace(auxUsuario,"ä"		,"�")
		auxUsuario = replace(auxUsuario,"é"		,"�")
		auxUsuario = replace(auxUsuario,"è"		,"�")
		auxUsuario = replace(auxUsuario,"ê"		,"�")
		auxUsuario = replace(auxUsuario,"ë"		,"�")
		auxUsuario = replace(auxUsuario,"í"		,"�") 
		auxUsuario = replace(auxUsuario,"ì"		,"�")
		auxUsuario = replace(auxUsuario,"î"		,"�")
		auxUsuario = replace(auxUsuario,"ï"		,"�")
		auxUsuario = replace(auxUsuario,"ó"		,"�")
		auxUsuario = replace(auxUsuario,"ò"		,"�")
		auxUsuario = replace(auxUsuario,"õ"		,"�")
		auxUsuario = replace(auxUsuario,"ô"		,"�")
		auxUsuario = replace(auxUsuario,"ö"		,"�")
		auxUsuario = replace(auxUsuario,"ú"		,"�")
		auxUsuario = replace(auxUsuario,"ù"		,"�")
		auxUsuario = replace(auxUsuario,"û"		,"�")
		auxUsuario = replace(auxUsuario,"ü"		,"�")
		auxUsuario = replace(auxUsuario,"ç"		,"�")
		auxUsuario = replace(auxUsuario,"ñ"		,"�")
		auxUsuario = replace(auxUsuario,"ý"		,"�")
		auxUsuario = replace(auxUsuario,"ÿ"		,"�")
		auxUsuario = replace(auxUsuario,"Ç"		,"�")
		auxUsuario = replace(auxUsuario,"Ã"		,"�")
		auxUsuario = replace(auxUsuario,"Õ"		,"�")
		auxUsuario = replace(auxUsuario,"Č"		,"�")
		auxUsuario = replace(auxUsuario,"É"		,"�")
		auxUsuario = replace(auxUsuario,"º"		,"�")
		auxUsuario = replace(auxUsuario,"°"		,"�")
		auxUsuario = replace(auxUsuario,"Í"			,"�")
		auxUsuario = replace(auxUsuario,"–"		,"-")
		auxUsuario = replace(auxUsuario,"À"		,"�")	
		auxUsuario = replace(auxUsuario,"�"		    ,"�")
		auxUsuario = replace(auxUsuario,"��"	    ,"�")
		auxUsuario = replace(auxUsuario,"��"	    ,"�")
		auxUsuario = replace(auxUsuario,"��"	    ,"�")
		auxUsuario = replace(auxUsuario,"��"	    ,"�")
		auxUsuario = replace(auxUsuario,"ª"	    ,"�")
		
	trataCharPost = auxUsuario
	
end function	


function DecodeUTF8(s)
  dim i
  dim c
  dim n

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191 
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s 
end function

function formataMin(hora)
	strhora = Split(hora ,":")
	strhoracerta = hora
	if Len(trim(strhora(1))) = 1 then
		certo = "0"&strhora(1)
		 strhoracerta = strhora(0) &":"& certo
	end if
	if Len(trim(strhora(1))) > 2 then
		 strhoracerta = mid(hora,1,len(hora) - 1)
	end if
	formataMin = strhoracerta
end function
Function AdicionaDiasUteis(Data, Dias)
	Dim i
	i = 1

	Do While i <= Dias
	Data = DateAdd("D",1,Data)
		If Weekday(Data) <> 1 And Weekday(Data) <> 7  Then
			i = i + 1	
		End If
	Loop
	AdicionaDiasUteis = Data
End Function


FUNCTION SortArray(varArray)
	For i = UBound(varArray) - 1 To 1 Step - 1
		MaxVal = varArray(i)
		MaxIndex = i

		For j = 0 To i
		If varArray(j) < MaxVal Then
		MaxVal = varArray(j)
		MaxIndex = j
		End If
		Next

		If MaxIndex < i Then
		varArray(MaxIndex) = varArray(i)
		varArray(i) = MaxVal
		End If
	Next 
END FUNCTION

Function BuscaSegundoDIaUtildoMes()
	strPrimeiraData = year(now)&"/"&month(now)&"/01"
	strSegundoDia =  AdicionaDiasUteis(strPrimeiraData,1)
	BuscaSegundoDIaUtildoMes = strSegundoDia
End Function 

Function QuantosDiasTemOMes(Mes,Ano)
  Select Case Mes
    Case 1,3,5,7,8,10,12: QuantosDiasTemOMes = 31
    Case 4,6,9,11: QuantosDiasTemOMes = 30
    Case Else
      If Ano Mod 4 = 0 And (Ano Mod 100 <> 0 Or Ano Mod 400 = 0) Then
        QuantosDiasTemOMes = 29
      Else
        QuantosDiasTemOMes = 28
      End If
  End Select
End Function

Function DigMes(Mes)
  Select Case Mes
    Case 1,2,3,4,5,6,7,8,9: DigMes = "0"
    Case 10,11,12: DigMes = ""
  End Select
End Function 

%>