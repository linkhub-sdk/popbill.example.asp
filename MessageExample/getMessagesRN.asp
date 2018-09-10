<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �������ۿ�û�� �Ҵ��� ���ۿ�û��ȣ(requestNum)�� ���ۻ��¸� Ȯ���մϴ�
	'  - �����׸� ���� �ڼ��� ������ "[���� API �����Ŵ���] >
	'  3.3.2. GetMessagesRN (���۳��� Ȯ�� - ��û��ȣ �Ҵ�)�� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	'�˺� ȸ�� ���̵�
	userID = "testkorea"

	' �������� ��û �� �Ҵ��� ���ۿ�û��ȣ(requestNum)
	requestNum = "20180905"
	
	On Error Resume Next

	Set result = m_MessageService.GetMessagesRN(testCorpNum, requestNum, UserID)
	
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
				<legend>���ڸ޽��� ���۰�� Ȯ��</legend>
				<ul>
					<% If code = 0 Then
						For i=0 To result.Count-1 
					%>
						<fieldset class="fieldset2">
							<legend>���ڸ޽��� ���۰�� [<%=i+1%>]</legend>
							<ul>
								<li>state (���ۻ��� �ڵ�) : <%=result.Item(i).state%> </li>
								<li>result (���۰�� �ڵ�) : <%=result.Item(i).result%> </li>
								<li>subject (�޽��� ����) : <%=result.Item(i).subject%> </li>
								<li>content (�޽��� ����) : <%=result.Item(i).content%> </li>
								<li>type (�޽��� ����) : <%=result.Item(i).msgType%> </li>
								<li>sendnum (�߽Ź�ȣ) : <%=result.Item(i).sendnum%> </li>
								<li>senderName (�߽��ڸ�) : <%=result.Item(i).senderName%> </li>
								<li>receiveNum (���Ź�ȣ) : <%=result.Item(i).receiveNum%> </li>
								<li>receiveName (�����ڸ�) : <%=result.Item(i).receiveName%> </li>
								<li>receiptDT (�����Ͻ�) : <%=result.Item(i).receiptDT%> </li>
								<li>sendDT (�����Ͻ�) : <%=result.Item(i).sendDT%> </li>
								<li>resultDT (���۰�� �����Ͻ�) : <%=result.Item(i).resultDT%> </li>
								<li>reserveDT (�����Ͻ�) : <%=result.Item(i).reserveDT%> </li>
								<li>tranNet (����ó�� �̵���Ż��) : <%=result.Item(i).tranNet%> </li>
							</ul>
						</fieldset>
					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>