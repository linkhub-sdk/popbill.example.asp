<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� ���ݿ����� ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 6����)
	' - https://docs.popbill.com/cashbill/asp/api#Search
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	


	'�˻����� ����, R-�������, T-�ŷ�����, I-��������
	DType = "T"						
	
	'��������, yyyyMMdd
	SDate = "20211201"				

	'��������, yyyyMMdd
	EDate = "20211230"				

	' ���ۻ��°� �迭, �̱����� ��ü��ȸ, �������°� 3�ڸ� �迭, 2,3��° �ڸ� ���ϵ�ī�� ��밡��
	Dim State(3)
	State(0) = "2**"
	State(1) = "3**"
	State(2) = "4**"
	
	'��������, N-�Ϲ����ݿ�����, C-������ݿ�����
	Dim TradeType(2)			
	TradeType(0) = "N"
	TradeType(1) = "C"

	'�ŷ�����, P-�ҵ������, C-����������
	Dim TradeUsage(2)		
	TradeUsage(0) = "P"
	TradeUsage(1) = "C"

	'�ŷ�����, N-�Ϲ�, B-��������, T-���߱���
	Dim TradeOpt(3)		
	TradeOpt(0) = "N"
	TradeOpt(1) = "B"
	TradeOpt(2) = "T"

	'�������� �迭, T-����,  N-�����
	Dim TaxationType(2)		
	TaxationType(0) = "T"
	TaxationType(1) = "N"


	'���Ĺ���, A-��������, D-��������
	Order = "D"			

	'��������ȣ
	Page = 1				

	'�������� �˻�����, �ִ� 1000
	PerPage = 20		

	'�ĺ���ȣ ����, ����ó���� ��ü��ȸ
	QString = ""		

	On Error Resume Next
	
	Set SearchResult = m_CashbillService.Search(testCorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TradeOpt, TaxationType, Order, Page, PerPage, QString)

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
				<legend>���ݿ����� �����ȸ</legend>
					<ul>
						<li> code (���� �ڵ�) : <%=SearchResult.code%></li>
						<li> message (���� �޽���) : <%=SearchResult.message%></li>
						<li> total (�� �˻���� �Ǽ�) : <%=SearchResult.total%></li>
						<li> pageNum (������ ��ȣ) : <%=SearchResult.pageNum%></li>
						<li> perPage (�������� �˻�����) : <%=SearchResult.perPage%></li>
						<li> pageCount (������ ����) : <%=SearchResult.pageCount%></li>
					</ul>
					<% If code = 0 Then 
						For i=0 To UBound(SearchResult.list)-1 %>
						<fieldset class="fieldset2">
							<legend> ���ݿ����� ��ȸ ��� [<%= i+1 %> / <%= SearchResult.total %>]</legend>
							<ul>
								<li>itemKey (���ݿ����� ������Ű) : <%=SearchResult.list(i).itemKey%></li>
								<li>confirmNum (����û ���ι�ȣ) : <%=SearchResult.list(i).confirmNum%></li>
								<li>mgtKey (������ȣ) : <%=SearchResult.list(i).mgtKey%></li>
								<li>tradeDate (�ŷ�����) : <%=SearchResult.list(i).tradeDate%></li>
								<li>issueDT (�����Ͻ�) : <%=SearchResult.list(i).issueDT%></li>
								<li>regDT (����Ͻ�) : <%=SearchResult.list(i).regDT%></li>
								<li>taxationType (��������) : <%=SearchResult.list(i).taxationType%></li>
								<li>totalAmount (�ŷ��ݾ�) : <%=SearchResult.list(i).totalAmount%></li>
								<li>tradeUsage (�ŷ�����) : <%=SearchResult.list(i).tradeUsage%></li>
								<li>tradeOpt (�ŷ�����) : <%=SearchResult.list(i).tradeOpt%></li>
								<li>tradeType (��������) : <%=SearchResult.list(i).tradeType%></li>
								<li>stateCode (�����ڵ�) : <%=SearchResult.list(i).stateCode%></li>
								<li>stateDT (���º����Ͻ�) : <%=SearchResult.list(i).stateDT%></li>
								<li>stateMemo (���¸޸�) : <%=SearchResult.list(i).stateMemo%></li>
								<li>identityNum (�ŷ�ó �ĺ���ȣ) : <%=SearchResult.list(i).identityNum%></li>
								<li>itemName (��ǰ��) : <%=SearchResult.list(i).itemName%></li>
								<li>customerName (����) : <%=SearchResult.list(i).customerName%></li>
								<li>ntssendDT (����û �����Ͻ�) : <%=SearchResult.list(i).ntssendDT%></li>
								<li>ntsresultDT (����û ó����� �����Ͻ�) : <%=SearchResult.list(i).ntsResultDT%></li>
								<li>ntsresultCode (����û ó����� �����ڵ�) : <%=SearchResult.list(i).ntsResultCode%></li>
								<li>ntsresultMessage (����û ó����� �޽���) : <%=SearchResult.list(i).ntsResultMessage%></li>
								<li>orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) : <%=SearchResult.list(i).orgConfirmNum%></li>
								<li>orgTradeDate (���� ���ݿ����� �ŷ�����) : <%=SearchResult.list(i).orgTradeDate%></li>
								<li>printYN (�μ⿩��) : <%=SearchResult.list(i).printYN%></li>
							</ul>
						</fieldset>
					<%	Next
						Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If%> 
				</fieldset>
		 </div>
	</body>
</html>