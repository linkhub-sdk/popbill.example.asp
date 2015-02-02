<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	userID = "testkorea"				'�˺� ȸ�� ���̵�
	itemCode = "121"					'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	mgtKey = "20150201-01"				'����������ȣ

	On Error Resume Next

	Set result = m_StatementService.GetLogs(testCorpNum, itemCode, mgtKey, userID)

	If Err.Number <> 0 Then
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
				<legend>���ڸ��� �̷� Ȯ��</legend>
				<ul>
					<% If code = 0 Then 
						For i=0 To result.Count-1%>
						<fieldset class="fieldset2">
						<legend> ���ڸ��� �̷�����[<%=i+1%>]</legend>
							<ul>
								<li>docLogType : <%=result.Item(i).docLogType%> </li>
								<li>log : <%=result.Item(i).log%> </li>
								<li>procType : <%=result.Item(i).procType%> </li>
								<li>procCorpName : <%=result.Item(i).procCorpName%> </li>
								<li>procMemo : <%=result.Item(i).procMemo%> </li>
								<li>regDT : <%=result.Item(i).regDT%> </li>
								<li>ip : <%=result.Item(i).ip%> </li>
							</ul>
						</fieldset>
					<%
						Next
						Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>