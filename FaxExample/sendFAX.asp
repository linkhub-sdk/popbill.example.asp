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

	'�߽��� ��ȣ
	sendNum = "07043042992"	

	'���ۿ���ð� yyyyMMddHHmmss,  ����ó���� �������
	reserveDT = ""	
	
	'������ ���� 
	Dim receivers(0)
	Set receivers(0) = New FaxReceiver

	'���Ź�ȣ
	receivers(0).receiverNum = "07043042999"

	'�����ڸ�
	receivers(0).receiverName = "������ ��Ī"

	'�ѽ������� ����
	FilePaths = Array("C:\popbill.example.asp\���ѹα����.doc")


	'�����ѽ� ���ۿ���
	adsYN = False

	'�ѽ�����
	title = "ASP  �ѽ� ���� �׽�Ʈ"

	On Error Resume Next

	url = m_FaxService.SendFAX(testCorpNum , sendNum, receivers, FilePaths, reserveDT , userID, adsYN, title )

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�ѽ� ����</legend>
				<ul>
					<% If code = 0 Then %>
						<li>recepitNum (������ȣ) : <%=url%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>