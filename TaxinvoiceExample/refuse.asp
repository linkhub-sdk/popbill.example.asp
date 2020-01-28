<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���޹޴��ڿ��� ��û���� ������ ���ݰ�꼭�� [�ź�]ó�� �մϴ�.
	' - ���ݰ�꼭�� ������ȣ�� �����ϱ� ���ؼ��� ���� (Delete API) ��
	'   ȣ���Ͽ� [����] ó���ؾ� �մϴ�.
	' - https://docs.popbill.com/taxinvoice/asp/api#Refuse
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	' �˺�ȸ�� ���̵�
	testUserID = "testkorea"
	
	' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
	KeyType = "SELL"

	' ������ȣ 
	MgtKey = "20190103-001"

	' �޸�
	Memo = "������ ��û �ź� �޸�"

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Refuse(testCorpNum, KeyType ,MgtKey, Memo, testUserID)
	
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
				<legend>�������û �ź�</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>