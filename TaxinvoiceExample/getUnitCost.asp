<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ڼ��ݰ�꼭 ����ܰ��� Ȯ���մϴ�.
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "123456789"		 
	
	On Error Resume Next

	unitCost = m_TaxinvoiceService.GetUnitCost(testCorpNum)
	
	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0 
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>���ݰ�꼭 ����ܰ� Ȯ�� </legend>
				<ul>
					<% If code = 0 Then %>
						<li>����ܰ� : <%=unitCost%> </li>
					<% Else %>
						<li> Response.code : <%=code%></li>
						<li> Response.message : <%=message%></li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>