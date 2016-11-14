<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1���� ���ڸ��� ����/��� ������ Ȯ���մϴ�.
	' - �����׸� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] > 3.3.1.
	'   GetInfo (���� Ȯ��)"�� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	testCorpNum = "1234567890"	
	
	'�˺� ȸ�� ���̵�
	userID = "testkorea"				

	'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	itemCode = "121"					

	'����������ȣ
	mgtKey = "20161114-10"				

	On Error Resume Next

	Set result = m_StatementService.GetInfo(testCorpNum, itemCode, mgtKey, userID)

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
				<legend>���ڸ��� ����/��� ����Ȯ��</legend>
				<ul>
					<% If code = 0 Then %>
						<li>itemKey : <%=result.itemKey%> </li>
						<li>stateCode : <%=result.stateCode%> </li>
						<li>taxType : <%=result.taxType%> </li>
						<li>purposeType : <%=result.purposeType%> </li>
						<li>writeDate : <%=result.writeDate%> </li>
						<li>senderCorpName : <%=result.senderCorpName%> </li>
						<li>senderCorpNum : <%=result.senderCorpNum%> </li>
						<li>senderPrintYN : <%=result.senderPrintYN%> </li>
						<li>receiverCorpName : <%=result.receiverCorpName%> </li>
						<li>receiverCorpNum : <%=result.receiverCorpNum%> </li>
						<li>receiverPrintYN : <%=result.receiverPrintYN%> </li>
						<li>supplyCostTotal : <%=result.supplyCostTotal%> </li>
						<li>taxTotal : <%=result.taxTotal%> </li>
						<li>issueDT : <%=result.issueDT%> </li>
						<li>stateDT : <%=result.stateDT%> </li>
						<li>openYN : <%=result.openYN%> </li>
						<li>stateMemo : <%=result.stateMemo%> </li>
						<li>regDT : <%=result.regDT%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>