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
	sendNum = "07075103710"			'�߽��� ��ȣ
	senderName = "�߽��ڸ�"			'�߽��ڸ�
'	reserveDT = "20150123200000"	'���ۿ���ð� yyyyMMddHHmmss, reserveDT���� null ��� �������
	
	'�����ڸ��, �ִ� 1000��
	Dim receivers(1)
	Set receivers(0) = New FaxReceiver
	receivers(0).receiverNum = "010111222"
	receivers(0).receiverName = "������ ��Ī"

	Set receivers(1) = New FaxReceiver
	receivers(1).receiverNum = "00011112222"
	receivers(1).receiverName = "������ ��Ī"

	FilePaths = Array("C:\popbill.example.asp\���ѹα����.doc")

	On Error Resume Next

	url = m_FaxService.SendFAX(testCorpNum, sendNum, senderName, receivers, FilePaths,  reserveDT , userID )

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
						<li>recepitNum : <%=url%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>