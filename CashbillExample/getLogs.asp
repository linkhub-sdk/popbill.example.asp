<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ݿ����� ���� �����̷��� Ȯ���մϴ�.
	' - https://docs.popbill.com/cashbill/asp/api#GetLogs
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	 

	'�˺� ȸ�� ���̵�
	userID = "testkorea"		 

	'������ȣ
	mgtKey = "20190103-001"		 

	On Error Resume Next
	
	Set Presponse = m_CashbillService.GetLogs(testCorpNum, mgtKey, userID)

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
				<legend>���ݿ����� �̷� Ȯ��</legend>
				<ul>
					<% If code = 0 Then
						For i=0 To Presponse.Count-1
					%>
						<fieldset class="fieldset2">
							 <ul>
								<li>docLogType(�α�Ÿ��) : <%=Presponse.Item(i).docLogType%></li>
								<li>log(�̷�����) : <%=Presponse.Item(i).log%></li>
								<li>procType(ó������) : <%=Presponse.Item(i).procType%></li>
								<li>procMemo(ó���޸�) : <%=Presponse.Item(i).procMemo%></li>
								<li>regDT(����Ͻ�) : <%=Presponse.Item(i).regDT%></li>
								<li>ip(������) : <%=Presponse.Item(i).ip%></li>
							</ul>
						</fieldset>
					<%	
						Next
						Else
					%>
						<li>Response.code : <%=code%></li>
						<li>Response.message : <%=message%><li>
					<%
						End If
					%>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>