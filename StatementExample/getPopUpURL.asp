<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1���� ���ڸ��� ���� �˾� URL�� ��ȯ�մϴ�.
	' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	'�˺� ȸ�� ���̵�
	userID = "testkorea"

	'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	itemCode = "121"

	'����������ȣ
	mgtKey = "20150201-01"

	On Error Resume Next

	url = m_StatementService.GetPopUpURL(testCorpNum, itemCode, mgtKey, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend> ���ڸ��� �˾� URL ��û</legend>
				<% If code = 0 Then %>
					<ul>
						<li>URL : <%=CStr(url)%> </li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>