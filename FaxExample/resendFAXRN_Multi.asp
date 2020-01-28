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
    ' - �����Ϸκ��� 60���� ����� ��� �������� �� �����ϴ�.
	' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
	' - https://docs.popbill.com/fax/asp/api#ResendFAXRN
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

	'�ѽ�����
	title = "�ѽ� ���� ������"
	
	'���������� �������������� ������ ���
'	ReDim receivers(-1)
	

	'���������� ������������ �ٸ� ��� �Ʒ� �ڵ� ����	
	Dim receivers(1)
	Set receivers(0) = New FaxReceiver
	receivers(0).receiverNum = "010111222"
	receivers(0).receiverName = "������ ��Ī"

	Set receivers(1) = New FaxReceiver
	receivers(1).receiverNum = "000111222"
	receivers(1).receiverName = "������ ��Ī"
	
	'���ۿ�û��ȣ (�˺� ȸ���� ���ߺ� ��ȣ �Ҵ�)
	'����,����,'-','_' ����, �ִ� 36��
	requestNum = ""		

	On Error Resume Next

	url = m_FaxService.ResendFAXRN(testCorpNum, orgRequestNum, sendNum, sendName, receivers, reserveDT, userID, title, requestNum)

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