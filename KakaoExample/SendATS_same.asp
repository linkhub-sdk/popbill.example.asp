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
	testUserID = "testkorea"					

	'�˸��� ���ø� �ڵ� - ���ø� ��� ��ȸ (ListATSTemplate API)�� ��ȯ�׸� Ȯ��
	templateCode = "018020000001"

	'�˺��� ���� ��ϵ� �߽Ź�ȣ
	senderNum = "07043042993"

	'�˸��� ����, �ִ� 400��
	content = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�."

	'��ü���� ����
	altContent = "��ü���� �޽��� ����"

	'��ü���� �������� ����-������, A-��ü���ڳ��� ����, C-�˸��峻�� ����
	altSendType = "C"

	'�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
	reserveDT = ""

	Set receiverList = CreateObject("Scripting.Dictionary")

	'�������� �迭, �ִ� 1000��
	For i =0 To 9
		Set rcvInfo = New KakaoReceiver

		'�����ڹ�ȣ
		rcvInfo.rcv = "01011222"+ CStr(i)			

		'�����ڸ�
		rcvInfo.rcvnm = " �������̸�"

		receiverList.Add i, rcvInfo
	Next 
	
	On Error Resume Next
	
	receiptNum = m_KakaoService.SendATS(testCorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, testUserID)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�˸��� ���ϳ��� �뷮����</legend>
				<% If code = 0 Then %>
					<ul>
						<li>ReceiptNum(������ȣ) : <%=receiptNum%> </li>
					</ul>
				<% Else %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<% End If %>
			</fieldset>
		 </div>
	</body>
</html>