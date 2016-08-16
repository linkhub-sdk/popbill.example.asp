<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	DType = "T"							'�˻����� ����, R-�������, T-�ŷ�����, I-��������
	SDate = "20160701"				'��������, yyyyMMdd
	EDate = "20160831"				'��������, yyyyMMdd

	' ���ۻ��°� �迭, �̱����� ��ü��ȸ, �������°� 3�ڸ� �迭, 2,3��° �ڸ� ���ϵ�ī�� ��밡��
	Dim State(3)
	State(0) = "2**"
	State(1) = "3**"
	State(2) = "4**"
	
	'���ݿ����� ����, N-�Ϲ����ݿ�����, C-������ݿ�����
	Dim TradeType(2)			
	TradeType(0) = "N"
	TradeType(1) = "C"

	'�ŷ��뵵 �迭, P-�ҵ������, C-����������
	Dim TradeUsage(2)		
	TradeUsage(0) = "P"
	TradeUsage(1) = "C"

	'�������� �迭, T-����,  N-�����
	Dim TaxationType(2)		
	TaxationType(0) = "T"
	TaxationType(1) = "N"

	Order = "D"			'���Ĺ���, A-��������, D-��������
	Page = 1				'��������ȣ
	PerPage = 20		'�������� �˻�����, �ִ� 1000

	QString = ""		'�ĺ���ȣ ����, ����ó���� ��ü��ȸ

	On Error Resume Next
	
	Set SearchResult = m_CashbillService.Search(testCorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TaxationType, Order, Page, PerPage, QString)

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
						<li> code : <%=SearchResult.code%></li>
						<li> total : <%=SearchResult.total%></li>
						<li> pageNum : <%=SearchResult.pageNum%></li>
						<li> perPage : <%=SearchResult.perPage%></li>
						<li> pageCount : <%=SearchResult.pageCount%></li>
						<li> message : <%=SearchResult.message%></li>
					</ul>
					<% If code = 0 Then 
						For i=0 To UBound(SearchResult.list)-1 %>
						<fieldset class="fieldset2">
							<legend> ���ݿ����� ��ȸ ��� [<%=i+1%> / <%=UBound(SearchResult.list)%>]</legend>
							<ul>
								<li>itemKey : <%=SearchResult.list(i).itemKey%></li>
								<li>mgtKey : <%=SearchResult.list(i).mgtKey%></li>
								<li>tradeDate : <%=SearchResult.list(i).tradeDate%></li>
								<li>issueDT : <%=SearchResult.list(i).issueDT%></li>
								<li>customerName : <%=SearchResult.list(i).customerName%></li>
								<li>itemName : <%=SearchResult.list(i).itemName%></li>
								<li>identityNum : <%=SearchResult.list(i).identityNum%></li>
								<li>taxactionType : <%=SearchResult.list(i).taxationType%></li>
								<li>totalAmount : <%=SearchResult.list(i).totalAmount%></li>
								<li>tradeUsage : <%=SearchResult.list(i).tradeUsage%></li>
								<li>tradeType : <%=SearchResult.list(i).tradeType%></li>
								<li>stateCode : <%=SearchResult.list(i).stateCode%></li>
								<li>stateDT : <%=SearchResult.list(i).stateDT%></li>
								<li>printYN : <%=SearchResult.list(i).printYN%></li>
								<li>confirmNum : <%=SearchResult.list(i).confirmNum%></li>
								<li>orgTradeDate : <%=SearchResult.list(i).orgTradeDate%></li>
								<li>orgConfirmNum : <%=SearchResult.list(i).orgConfirmNum%></li>
								<li>ntssendDT : <%=SearchResult.list(i).ntssendDT%></li>
								<li>ntsPresponse : <%=SearchResult.list(i).ntsResult%></li>
								<li>ntsPresponseDT : <%=SearchResult.list(i).ntsResultDT%></li>
								<li>ntsPresponseCode : <%=SearchResult.list(i).ntsResultCode%></li>
								<li>ntsPresponseMessage : <%=SearchResult.list(i).ntsResultMessage%></li>
								<li>regDT : <%=SearchResult.list(i).regDT%></li>
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