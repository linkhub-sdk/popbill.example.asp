<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1231212312"	'ȸ�� ����ڹ�ȣ, "-" ����
	testUserID = "userid"    'ȸ�� ���̵�
	KeyType= "BUY"             '�������� SELL(����), BUY(����), TRUSTEE(����Ź)
	MgtKey = "20150121-15"      '����������ȣ 
	Memo = "������ ��û ��� �޸�"     '�޸�

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.CancelRequest(testCorpNum, KeyType ,MgtKey, Memo, testUserID)
	
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
				<legend>���ݰ�꼭 ��)�����û ó��</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>