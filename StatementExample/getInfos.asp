<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"
	userID = "testkorea"
	itemCode = "121"					'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)

	Dim mgtKeyList(2)
	mgtKeyList(0) = "20150202-03"
	mgtKeyList(1) = "20150202-04"

	On Error Resume Next

	Set result = m_StatementService.GetInfos(testCorpNum, itemCode, mgtKeyList, userID)

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
				<legend>���ڸ��� ���� �뷮 Ȯ��</legend>
				<ul>
					<% If code = 0 Then 
						For i=0 To result.Count-1 %>

						<fieldset class="fieldset2">
							<legend> ���ڸ��� ��ȸ��� [<%=i+1%>] </legend>
							<ul>
								<li>itemKey : <%=result.Item(i).itemKey%> </li>
								<li>stateCode : <%=result.Item(i).stateCode%> </li>
								<li>taxType : <%=result.Item(i).taxType%> </li>
								<li>purposeType : <%=result.Item(i).purposeType%> </li>
								<li>writeDate : <%=result.Item(i).writeDate%> </li>
								<li>senderCorpName : <%=result.Item(i).senderCorpName%> </li>
								<li>senderCorpNum : <%=result.Item(i).senderCorpNum%> </li>
								<li>receiverCorpName : <%=result.Item(i).receiverCorpName%> </li>
								<li>receiverCorpNum : <%=result.Item(i).receiverCorpNum%> </li>
								<li>supplyCostTotal : <%=result.Item(i).supplyCostTotal%> </li>
								<li>taxTotal : <%=result.Item(i).taxTotal%> </li>
								<li>issueDT : <%=result.Item(i).issueDT%> </li>
								<li>stateDT : <%=result.Item(i).stateDT%> </li>
								<li>openYN : <%=result.Item(i).openYN%> </li>
								<li>stateMemo : <%=result.Item(i).stateMemo%> </li>
								<li>regDT : <%=result.Item(i).regDT%> </li>
							</ul>
						</fieldset>
					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>