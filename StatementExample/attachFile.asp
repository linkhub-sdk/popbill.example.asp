<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ڸ����� ÷�������� ����մϴ�.
	' - ÷������ ����� ���ڸ����� [�ӽ�����] ������ ��쿡�� �����մϴ�.
	' - ÷�������� �ִ� 5������ ����� �� �ֽ��ϴ�.
	' - https://docs.popbill.com/statement/asp/api#AttachFile
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	testCorpNum = "1234567890"
	
	'�˺� ȸ�� ���̵�
	userID = "testkorea"

	'���� �����ڵ� - 121(�ŷ�����), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
	itemCode = "121"

	'����������ȣ
	mgtKey = "20211201-001"

	'÷�� ���� ���
	filePath = "C:\popbill.example.asp\Popbill\�ΰ�.gif"

	On Error Resume Next

	Set result = m_StatementService.AttachFile(testCorpNum, itemCode, mgtKey, filePath, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = result.code
		message = result.message
	End If

	On Error GoTo 0

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>���ڸ��� ����÷��</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>