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

	'�޽��� ����
	subject = "�������� �޽��� ����"

	'�޽��� �פ��� 
	content = "�������� �޽��� ����"
	
	Set msgList = CreateObject("Scripting.Dictionary")
	
	'������������ �迭, �ִ� 1000��
	For i =0 To 99

		Set message = New Messages
		'���Ź�ȣ
		message.receiver = "000111222"

		'�����ڸ�
		message.receivername = " �������̸�"+CStr(i)
		msgList.Add i, message
	Next
		
	'����޽��� �̹�������, 300Kbyte JPEG ���� ���۰���
	FilePaths = Array("C:\popbill.example.asp\test.jpg")

	On Error Resume Next

	receiptNum = m_MessageService.SendMMS(testCorpNum, senderNum, subject, content, msgList, FilePaths, reserveDT, adsYN, userID)

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
				<legend>MMS ���ڸ޽��� ����</legend>
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