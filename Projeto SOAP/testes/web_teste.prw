#Include 'Protheus.ch'

//*****************************************************************************
//O c�digo abaixo representa o arquivo ms01.APH, que cont�m a parte para Login de um usuario
<html>
	<h2 align="center"> Login </h2><hr>
	<form name="form1" method="post" action="w_ms02.apw">
	<p>Nome : <input name="txt_Nome" type="text" id="txt_Nome" size="25"></p>
	<p>Senha : <input name="txt_Senha" type="password" id="txt_Senha" size="3" maxlength="3"></p>
	<hr>	<p><input type="submit" value="Ok"></p></form>
</html>

//*****************************************************************************
//O c�digo abaixo representa o arquivo ms02.APH, que cont�m a parte do formul�rio
<html>
<head><title>ADVPL ASP</title></head>

//Codigo JavaScript no qual n�o permite que o formul�rio seja enviado sem que seus campos tenham sido preechidos.
<script language="javascript">
function envia()
{	var oFrm = document.forms[0];
	if ( oFrm.txt_Nome.value == "" || oFrm.txt_Pre.value == "" || oFrm.txt_Fone.value == "" || oFrm.txt_End.value == "" )
		{ alert( "Preencha Todos Os Dados Do Formul�rio" );	 return; }

		oFrm.action = "w_ms03.apw";
		oFrm.submit();
}
</script><body>
<h2 align="center"> Formul�rio</h2>
<hr>
<p>Bem Vindo <%=HttpSession->Usuario%></p>
<form name="form" method="post" action="">		<p>Nome : <input name="txt_Nome" type="text" id="txt_Nome" size="25" value=""></p>		<p>Telefone : <input name="txt_Pre" type="text" id="txt_Pre" size="3"> - 					<input name="txt_Fone" type="text" id="txt_Fone" size="10"></p>		<p>Endere�o : <input name="txt_End" type="text" id="txt_End" size="25"></p>			<p><input type="button" value="Enviar" onClick="envia()"></p>	</form>	<hr></body></html>

//*****************************************************************************
//| O c�digo abaixo representa o arquivo ms03.APH, que cont�m uma tabela que exibe os dados
//| preenchidos no formul�rio, mais um contador do total de vezes que foi realizado esse formul�rio

<HTML><table width="200" border="1">
<tr>
	<td width="95">Nome</td>
	<td colspan="2"><%=HttpPost->txt_Nome%></td>
</tr>
<tr>
	<td width="95">Telefone</td>
	<td width="75"><%=HttpPost->txt_Pre %></td>
	<td width="75"><%=HttpPost->txt_Fone%></td>
</tr>
<tr>
	<td>Endere�o;</td>
	<td colspan="2"><%=HttpPost->txt_End%></td>
</tr>
<tr>
	<td width="95">Contador</td>
	<td colspan="2"><%=HttpSession->Contador%></td>
</tr>
</table>
<P><input name="Reset" type="reset" value="Voltar" onClick="window.location = 'w_ms02.apw'"></P>
</HTML>


//| O c�digo abaixo representa o arquivo ms01.PRW, que cont�m as fun��es escritas em ADVPL ASP
#INCLUDE "PROTHEUS.CH"
#DEFINE ID "Admin"
#DEFINE SENHA "123"
//*****************************************************************************
web function ms01()//A fun��o � executada quando � chamada atrav�s do browser.

return h_ms01()

//*****************************************************************************
web function ms02()	//Verifica se � a primeira vez q usu�rio faz login.

conout( ID, SENHA )

if empty( HttpSession->Usuario )		//Verifica se os campos foram preenchidos.

if empty( HttpPost->txt_Nome ) .And. empty( HttpPost->txt_Senha)

return "Nome e Senha n�o informados!!"

endif		//Verifica usu�rio e senha.

if HttpPost->txt_Nome != ID

return "Usu�rio Inv�lido!!"

endif

if HttpPost->txt_Senha != SENHA

return "Senha Inv�lida!!"

endif		//Seta o nome do usuario.

HttpSession->Usuario := HttpPost->txt_Nome

endif

return h_ms02()

//*****************************************************************************
web function ms03() 	//Verifica se a Sesssion j� foi iniciada.

if empty( HttpSession->Contador )

HttpSession->Contador := 1	//caso tenha sido, incrementa o contador.

else

HttpSession->Contador++

endif

return h_ms03()


