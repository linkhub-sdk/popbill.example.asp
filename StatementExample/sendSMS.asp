<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˸����ڸ� �����մϴ�. (�ܹ�/SMS- �ѱ� �ִ� 45��)
	' - �˸����� ���۽� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
	' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [���۳���] �ǿ���
	'   ���۰���� Ȯ���� �� �ֽ��ϴ�.
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

	'���Ź�ȣ
	receiver = "010111222"

	'�޽��� ����, 90byte�ʰ��� ���̰� �����Ǿ� ���۵�
	contents = "���ڸ��� �˸��������� �׽�Ʈ�Դϴ�."

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