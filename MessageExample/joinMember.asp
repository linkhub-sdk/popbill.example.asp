<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	Set joinInfo = New JoinForm

	joinInfo.LinkID = "TESTER"		   '�������̵� 
	joinInfo.CorpNum = "1234567890"    '����ڹ�ȣ, "-"���� 10�ڸ�
	joinInfo.CEOName = "��ǥ�ڼ���"	   '��ǥ�ڼ���
	joinInfo.CorpName =  "��ȣ"		   '��ȣ��
	joinInfo.Addr =   "�ּ�"		   '�ּ�
	joinInfo.ZipCode =  "500-100"	   '�����ȣ
	joinInfo.BizType =  "����"		   '����
	joinInfo.BizClass = "����"		   '����
	joinInfo.ID =  "userid"		       '���̵� (6�� �̻� 20�� �̸�)
	joinInfo.PWD =  "1234567890"       '��й�ȣ (6�� �̻� 20�� �̸�)
	joinInfo.ContactName = "����ڸ�"    '����ڸ�
	joinInfo.ContactTEL = "02-999-9999"   '����ڿ���ó
	joinInfo.ContactHP = "010-1234-5678"	'����� �޴�����ȣ
	joinInfo.ContactFAX = "02-999-9999"		'�ѽ���ȣ
	joinInfo.ContactEmail = "test@test.com" '����� �̸���

	On Error Resume Next

	Set Presponse = m_MessageService.JoinMember(joinInfo)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message =Presponse.message
	End If

	On Error GoTo 0

%>

	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>����ȸ�� ���Կ�û</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>