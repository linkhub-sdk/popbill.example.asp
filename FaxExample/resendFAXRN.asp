<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ۿ�û��ȣ(requestNum)�� �Ҵ��� �ѽ��� �������մϴ�.
    ' - �����Ϸκ��� 180���� ����� ��� �������� �� �����ϴ�.
	' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
	' - �ѽ����� ���� �������� �ȳ� : http://blog.linkhub.co.kr/2561
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'�˺� ȸ�� ���̵�
	userID = "testkorea"			
	
	'���� �ѽ� ���۽� �Ҵ��� ���ۿ�û��ȣ(requestNum)
	orgRequestNum = "1"
	
	'�߽��� ��ȣ
	sendNum = "07043042991"
	
	'�߽��ڸ�
	sendName = "�߽��ڸ�"

	'���ۿ���ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
	reserveDT = ""	

	'�ѽ� ����
	title = "�ѽ� ������"
	
	'���������� �������������� ������ ���
	ReDim receivers(-1)

	'���������� ������������ �ٸ� ��� �Ʒ� �ڵ� ����	
'	Dim receivers(0)
'	Set receivers(0) = New FaxReceiver
	
	'���Ź�ȣ
'	receivers(0).receiverNum = "07066666"

	'�����ڸ�
'	receivers(0).receiverName = "������ ��Ī"

	'���ۿ�û��ȣ (�˺� ȸ���� ���ߺ� ��ȣ �Ҵ�)
	'����,����,'-','_' ����, �ִ� 36��
	requestNum = ""

	On Error Resume Next

	url = m_FaxService.ResendFAXRN(testCorpNum, orgRequestNum, sendNum, senderName, receivers, reserveDT , userID, title, requestNum)

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