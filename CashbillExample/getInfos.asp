<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �ټ����� ���ݿ����� ����/��� ������ Ȯ���մϴ�. (�ִ� 1000��)
	' - �����׸� ���� �ڼ��� ������ "[���ݿ����� API �����Ŵ���] > 4.2.
	'   ���ݿ����� �������� ����"�� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************
	
	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	 

	'�˺� ȸ�� ���̵�
	userID = "testkorea"		 

	'��ȸ�� ���ݿ����� ����������ȣ �迭, �ִ� 1000��
	Dim mgtKeyList(3) 
	MgtKeyList(0) = "20190103-001"
	MgtKeyList(1) = "20190103-002"
	MgtKeyList(2) = "20190103-003"

	On Error Resume Next
	
	Set Presponse = m_CashbillService.GetInfos(testCorpNum, mgtKeyList, userID)

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
				<legend>���ݿ����� ���� �뷮 Ȯ��</legend>
				<ul>
					<% If code = 0 Then 
						For i=0 To Presponse.Count-1 %>
						<fieldset class="fieldset2">
							<legend> ���ݿ����� ��ȸ ��� [<%=i+1%>]</legend>
							<ul>
								<li>itemKey (���ݿ����� ������Ű) : <%=Presponse.Item(i).itemKey%></li>
								<li>confirmNum (����û ���ι�ȣ) : <%=Presponse.Item(i).confirmNum%></li>
								<li>mgtKey (����������ȣ) : <%=Presponse.Item(i).mgtKey%></li>
								<li>tradeDate (�ŷ�����) : <%=Presponse.Item(i).tradeDate%></li>
								<li>issueDT (�����Ͻ�) : <%=Presponse.Item(i).issueDT%></li>
								<li>regDT (����Ͻ�) : <%=Presponse.Item(i).regDT%></li>
								<li>taxationType (��������) : <%=Presponse.Item(i).taxationType%></li>
								<li>totalAmount (�ŷ��ݾ�) : <%=Presponse.Item(i).totalAmount%></li>
								<li>tradeUsage (�ŷ�����) : <%=Presponse.Item(i).tradeUsage%></li>
								<li>tradeOpt (�ŷ�����) : <%=Presponse.Item(i).tradeOpt%></li>
								<li>tradeType (��������) : <%=Presponse.Item(i).tradeType%></li>
								<li>stateCode (�����ڵ�) : <%=Presponse.Item(i).stateCode%></li>
								<li>stateDT (���º����Ͻ�) : <%=Presponse.Item(i).stateDT%></li>
								<li>stateMemo (���¸޸�) : <%=Presponse.Item(i).stateMemo%></li>
								<li>identityNum (�ŷ�ó �ĺ���ȣ) : <%=Presponse.Item(i).identityNum%></li>
								<li>itemName (��ǰ��) : <%=Presponse.Item(i).itemName%></li>
								<li>customerName (����) : <%=Presponse.Item(i).customerName%></li>
								<li>ntssendDT (����û �����Ͻ�) : <%=Presponse.Item(i).ntssendDT%></li>
								<li>ntsresultDT (����û ó����� �����Ͻ�) : <%=Presponse.Item(i).ntsResultDT%></li>
								<li>ntsresultCode (����û ó����� �����ڵ�) : <%=Presponse.Item(i).ntsResultCode%></li>
								<li>ntsresultMessage (����û ó����� �޽���) : <%=Presponse.Item(i).ntsResultMessage%></li>
								<li>orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) : <%=Presponse.Item(i).orgConfirmNum%></li>
								<li>orgTradeDate (���� ���ݿ����� �ŷ�����) : <%=Presponse.Item(i).orgTradeDate%></li>
								<li>printYN (�μ⿩��) : <%=Presponse.Item(i).printYN%></li>
							</ul>
							</fieldset>
					<%	Next
						Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If%> 
					
				</ul>
			</fieldset>
		 </div>
	</body>
</html>