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
	' - https://docs.popbill.com/cashbill/asp/api#SendSMS
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	 

	'�˺� ȸ�� ���̵�
	userID = "testkorea"		 

	'������ȣ	
	mgtKey = "20211201-001"		 

	'�߽Ź�ȣ	
	sender = "07043242991"	
	
	'���Ź�ȣ
	receiver = "010111222"		 

	'�޽��� ����, 90byte�� �ʰ��� ������ ���̰� �����Ǿ� ���۵˴ϴ�.
	contents = "���ݿ����� �˸����� �׽�Ʈ�Դϴ�"  

	On Error Resume Next 

	Set Presponse = m_CashbillService.SendSMS(testCorpNum, mgtKey, Sender, Receiver, Contents, UserID)

	If Err.Number <> 0 then
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
				<legend>�˸����� ������ </legend>
				<ul>
					<li>Response.code : <%=code%></li>
					<li>Response.message : <%=message%></li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>