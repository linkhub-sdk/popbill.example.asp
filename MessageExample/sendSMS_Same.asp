<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	'�˺� ȸ�� ���̵�
	userID = "testkorea"

	'������ ���ۿ���
	adsYN = False							

	'�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
	reserveDT = ""
	
	'�߽Ź�ȣ
	senderNum = "07043042991"

	'�޽��� ����, 90byte�ʰ��� ���̰� �����Ǿ� ���۵�
	content = "�����޽��� ����"
	
	'���������迭, �ִ� 1000��
	Set msgList = CreateObject("Scripting.Dictionary")

	For i = 0 To 99
		Set message = New Messages

		'���Ź�ȣ
		message.receiver = "000111222"

		'�����ڸ�
		message.receivername = " �������̸�"+CStr(i)

		msgList.Add i, message
	Next
	
	'���ۿ�û��ȣ (�˺� ȸ���� ���ߺ� ��ȣ �Ҵ�)
	'����,����,'-','_' ����, �ִ� 36��
	requestNum = ""	

	On Error Resume Next

	receiptNum = m_MessageService.SendSMS(testCorpNum, senderNum, content, msgList, reserveDT, adsYN, requestNum, userID)

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
				<legend>�ܹ� ���ڸ޽��� �������� </legend>
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