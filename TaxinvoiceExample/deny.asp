<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���࿹�� ���ݰ�꼭�� [�ź�]ó�� �մϴ�.
	' - [�ź�]ó���� ���ݰ�꼭�� ����(Delete API)�ϸ� ��ϵ� ����������ȣ��
	'   ������ �� �ֽ��ϴ�.
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
	testCorpNum = "1231212312"	  

	' �˺�ȸ�� ���̵�
	testUserID = "userid"		  

	' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
	KeyType= "BUY"				  

	' ����������ȣ 
	MgtKey = "20150122-23"        

	' �޸�
	Memo = "���࿹���ź� �޸�"    

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Deny(testCorpNum, KeyType ,MgtKey, Memo ,testUserID)
	
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
				<legend>���ݰ�꼭 ���࿹�� �ź�</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>