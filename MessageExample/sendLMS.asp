<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	userID = "testkorea"					'�˺� ȸ�� ���̵�
	adsYN = False							'������ ���ۿ���
'	reserveDT = "20150128200000"    '�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������

	Set msgList = CreateObject("Scripting.Dictionary")

	Set message = New Messages
	message.sender = "07075100000"
	message.senderName = "�߽��ڸ�"
	message.receiver = "010111222"
	message.receivername = " �������̸�"
	message.content = "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. This is Message �޽��� �׽�Ʈ��"
	message.subject = "�幮 �����Դϴ�"

	msgList.Add 0, message
	
	On Error Resume Next
	
	receiptNum = m_MessageService.SendLMS(testCorpNum, "", "", "","", msgList, reserveDT, adsYN, userID)

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
				<legend>�幮 ���ڸ޽��� 1�� ���� </legend>
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