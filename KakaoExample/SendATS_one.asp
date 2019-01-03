<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' �˸��� ������ ��û�մϴ�.
    ' ������ ���ε� ���ø��� ����� �˸��� ���۳���(content)�� �ٸ� ��� ���۽��� ó���˴ϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'�˺� ȸ�� ���̵�
	testUserID = "testkorea"					

	'�˸��� ���ø� �ڵ� - ���ø� ��� ��ȸ (ListATSTemplate API)�� ��ȯ�׸� Ȯ��
	templateCode = "018080000079"

	'�˺��� ���� ��ϵ� �߽Ź�ȣ
	senderNum = "	07043042992"

	'�˸��� ����, �ִ� 1000��
	content = "[�׽�Ʈ] �׽�Ʈ ���ø��Դϴ�.dfdfdf"

	'��ü���� ����
	altContent = "��ü���� �޽��� ����"

	'��ü���� �������� ����-������, A-��ü���ڳ��� ����, C-�˸��峻�� ����
	altSendType = "C"

	'�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
	reserveDT = "20180315200000"

	Set receiverList = CreateObject("Scripting.Dictionary")

	'�޽��� ��������
	Set rcvInfo = New KakaoReceiver

	'�����ڹ�ȣ
	rcvInfo.rcv = "01011222"			

	'�����ڸ�
	rcvInfo.rcvnm = " �������̸�"		

	receiverList.Add 0, rcvInfo
	
	'���ۿ�û��ȣ (�˺� ȸ���� ���ߺ� ��ȣ �Ҵ�)
	'����,����,'-','_' ����, �ִ� 36��
	requestNum = ""		

	On Error Resume Next
	
	receiptNum = m_KakaoService.SendATS(testCorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, requestNum, testUserID)

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
				<legend>�˸��� 1�� ����</legend>
				<% If code = 0 Then %>
					<ul>
						<li>ReceiptNum(������ȣ) : <%=receiptNum%> </li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>