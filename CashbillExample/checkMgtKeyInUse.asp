<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	mgtKey = "20150129-04"		'����������ȣ
		
	On Error Resume Next

	Presponse = m_CashbillService.CheckMgtKeyInUse(testCorpNum, mgtKey)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
	Else	
		If Presponse = True Then
			code = 1
			message = "�����"
		Else
			code = 0 
			message = "�̻����"
		End If
	End If 

	On Error GoTo 0

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>����������ȣ ��뿩�� Ȯ��</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>