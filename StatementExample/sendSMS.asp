<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	userID = "testkorea"				'�˺� ȸ�� ���̵�
	itemCode = "121"					'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	mgtKey = "20150130-01"				'����������ȣ
	sender = "07075103710"				'�߽��ڹ�ȣ
	receiver = "010111222"			'�����ڹ�ȣ
	contents = "���ڸ��� �˸��������� �׽�Ʈ�Դϴ�."		'�޽�������

	On Error Resume Next

	Set Presponse = m_StatementService.SendSMS(testCorpNum, itemCode, mgtKey, sender, receiver, contents, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�˸����� ������</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>