<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	userID = "testkorea"			'�˺� ȸ�� ���̵�
	reserveDT = "20150128200000"    '�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������

	Set msgList = CreateObject("Scripting.Dictionary")

	Set message = New Messages
	message.sender = "07075100000"				'�߽��ڹ�ȣ
	message.receiver = "01011112222"			'�����ڹ�ȣ
	message.receivername = " �������̸�"		'�����ڸ�
	message.content = "�˺� �޽��� API �׽�Ʈ��" '�޽��� ����(�ܹ��޽����� ���, 90byte �ʰ��� ������ ���̰� �����Ǿ� ���۵˴ϴ�)

	msgList.Add 0, message
	
	receiptNum = m_MessageService.SendSMS(testCorpNum, "","",msgList, reserveDT, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�ܹ� ���ڸ޽��� 1�� ���� </legend>
				<% If code = 0 Then %>
					<ul>
						<li>ReceiptNum(������ȣ) : <%=receiptNum%> </li>
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