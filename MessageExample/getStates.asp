<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���� ���۳��� ��������� Ȯ���մϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	'�˺� ȸ�� ���̵�
	userID = "testkorea"

	'���� ���۽� �߱޹��� ������ȣ �迭(ReceiptNumList)
	Dim ReceiptNumList(2)
	ReceiptNumList(0) = "018041717000000018"
	ReceiptNumList(1) = "018041717000000019"
	
	On Error Resume Next

	Set result = m_MessageService.GetStates(testCorpNum, ReceiptNumList, UserID)
	
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
				<legend>���ڸ޽��� ������� Ȯ��</legend>
				<ul>
					<% If code = 0 Then
							For i=0 To result.Count-1 
					%>
						<fieldset class="fieldset2">
							<legend>���ڸ޽��� ���۰�� [<%=i+1%>]</legend>
							<ul>
								<li>rNum (������ȣ) : <%=result.Item(i).rNum%> </li>
								<li>sn (�Ϸù�ȣ) : <%=result.Item(i).sn%> </li>
								<li>stat (���� �����ڵ�) : <%=result.Item(i).stat%> </li>
								<li>rlt (���� ����ڵ�) : <%=result.Item(i).rlt%> </li>
								<li>sDT (�����Ͻ�) : <%=result.Item(i).sDT%> </li>
								<li>rDT (����ڵ� �����Ͻ�) : <%=result.Item(i).rDT%> </li>
								<li>net (���� �̵���Ż��) : <%=result.Item(i).net%> </li>
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