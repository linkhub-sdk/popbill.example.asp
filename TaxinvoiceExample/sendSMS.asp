<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'ȸ�� ����ڹ�ȣ, "-" ����
	testUserID = "testkorea"    'ȸ�� ���̵�
	KeyType= "SELL"             '�������� SELL(����), BUY(����), TRUSTEE(����Ź)
	MgtKey = "20150121-17"      '����������ȣ 
	Sender = "01011112222"		'�߽Ź�ȣ
	Receiver = "01043245117"    '���Ź�ȣ
	Contents = "���� �׽�Ʈ�Դϴ� 90Bytes�� �ʰ��ѳ����� ���۵��� �ʽ��ϴ�" 
	
	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.SendSMS(testCorpNum, KeyType, MgtKey, Sender, Receiver, Contents, testUserID)
	
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
				<legend>���� ������</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>