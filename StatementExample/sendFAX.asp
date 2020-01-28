<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ڸ����� �ѽ������մϴ�.
	' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
	' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [�ѽ�] > [���۳���]
	'   �޴����� ���۰���� Ȯ���� �� �ֽ��ϴ�.
	' - https://docs.popbill.com/statement/asp/api#SendFAX
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	testCorpNum = "1234567890"
	
	'�˺� ȸ�� ���̵�
	userID = "testkorea"

	'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
	itemCode = "121"

	'������ȣ
	mgtKey = "20190103-001"

	'�߽Ź�ȣ
	sender = "07043042991"

	'�����ѽ���ȣ
	receiver = "070111222"

	On Error Resume Next

	Set result = m_StatementService.SendFAX(testCorpNum, itemCode, mgtKey, sender, receiver, userID)

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
				<legend>���ڸ��� �ѽ� ����</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>