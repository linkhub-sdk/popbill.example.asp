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
'	reserveDT = "20150128200000"    '�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
	
	Set msgList = CreateObject("Scripting.Dictionary")

	Set message = New Messages
	message.sender = "07075103710"
	message.receiver = "01043245117"
	message.receivername = " �������̸�"
	message.content = "MMS �޽��� �׽�Ʈ��"
	message.subject = "MMS �޽��� �����Դϴ�"

	msgList.Add 0, message
	
	FilePaths = Array("C:\popbill.example.asp\test.jpg")

	On Error Resume Next

	receiptNum = m_MessageService.SendMMS(testCorpNum,"","","", msgList, FilePaths, reserveDT, userID)

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
				<legend>MMS ���ڸ޽��� 1�� ���� </legend>
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