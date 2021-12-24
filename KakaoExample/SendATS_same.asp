<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' [��������] �˸��� ������ ��û�մϴ�.
    ' - ������ ���ε� ���ø��� ����� �˸��� ���۳���(content)�� �ٸ� ��� ���۽��� ó���˴ϴ�
	' - https://docs.popbill.com/kakao/asp/api#SendATS
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'�˺� ȸ�� ���̵�
	testUserID = "testkorea"					

	'�˸��� ���ø� �ڵ� - ���ø� ��� ��ȸ (ListATSTemplate API)�� ��ȯ�׸� Ȯ��
	templateCode = "019020000163"

	'�˺��� ���� ��ϵ� �߽Ź�ȣ
	senderNum = "01043245117"

	'�˸��� ����, �ִ� 1000��
	content = "[ �˺� ]" & vbCrLf
	content = content + "��û�Ͻ� #{���ø��ڵ�}�� ���� �ɻ簡 �Ϸ�Ǿ� ���� ó���Ǿ����ϴ�." & vbCrLf
	content = content + "�ش� ���ø����� ���� �����մϴ�." & vbCrLf & vbCrLf
	content = content + "���ǻ��� �����ø� ��Ʈ�ʼ��ͷ� ���ϰ� �����ֽñ� �ٶ��ϴ�. " & vbCrLf & vbCrLf
	content = content + "�˺� ��Ʈ�ʼ��� : 1600-8536" & vbCrLf
	content = content + "support@linkhub.co.kr"

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

		'��Ʈ�� ����Ű, ������ ������ �޸�, �̻��� ����ó��
		rcvInfo.interOPRefKey = "20211201-" +CStr(i)

		receiverList.Add i, rcvInfo
	Next 
	
	
	
	'���ۿ�û��ȣ (�˺� ȸ���� ���ߺ� ��ȣ �Ҵ�)
	'����,����,'-','_' ����, �ִ� 36��
	requestNum = ""	

	' �˸��� ��ư������ ���ø� ��û�� ������ ��ư������ �����ϰ� �����ϴ� ��� btnList�� ���� �ϰ� �Լ�ȣ��.
	Set btnList = CreateObject("Scripting.Dictionary")
	
	'�˸��� ��ư URL�� #{���ø�����}�� �����Ѱ�� ���ø����� ������ �����Ͽ� ��ư���� ����
	'Set btnInfo = New KakaoButton
	'btnInfo.n = "���ø� �ȳ�"			
	'btnInfo.t = "WL"		
	'btnInfo.u1 = "https://www.popbil.com"
	'btnInfo.u2 = "http://www.llinkhub.co.kr"
	'btnList.Add 0, btnInfo

	On Error Resume Next

	receiptNum = m_KakaoService.SendATS(testCorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, requestNum, testUserID, btnList)

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