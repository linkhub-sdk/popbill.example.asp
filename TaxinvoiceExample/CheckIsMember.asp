<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1231212312"		' ����ڹ�ȣ ("-"����)
		
	On Error Resume Next

	Set result = m_TaxinvoiceService.CheckIsMember(testCorpNum,LinkID)

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
		Response.end
	End If

	On Error GoTo 0


%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>����ȸ���� ���� ���� Ȯ�� ���</legend>
				<ul>
					<li>Response.code : <%=CStr(result.code)%></li>
					<li>Response.message : <%=result.message%></li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>