<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �ѽ��� �������մϴ�.
	' - �����Ϸ� 180���� ������� ���� �Ǹ� ������ �����մϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'�˺� ȸ�� ���̵�
	userID = "testkorea"			
	
	'�ѽ� ������ȣ 
	receiptNum = "017021616254800001"
	
	'�߽��� ��ȣ
	sendNum = "070111222"		
	
	sendName = "�߽��ڸ�9999"

	'���ۿ���ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
	reserveDT = ""	
	
	'���������� �������������� ������ ���
	ReDim receivers(-1)


	'���������� ������������ �ٸ� ��� �Ʒ� �ڵ� ����	
'	Dim receivers(0)
'	Set receivers(0) = New FaxReceiver
	
	'���Ź�ȣ
'	receivers(0).receiverNum = "07066666"

	'�����ڸ�
'	receivers(0).receiverName = "������ ��Ī"

	On Error Resume Next

	url = m_FaxService.ResendFAX(testCorpNum, receiptNum, sendNum, senderName, receivers, reserveDT , userID)

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
				<legend>�ѽ� ������</legend>
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