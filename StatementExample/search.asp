<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� ���ڸ��� ����� ��ȸ�մϴ�.
	' - https://docs.popbill.com/statement/asp/api#Search
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ
	testCorpNum = "1234567890"

	'�˻����� ����, R-�������, W-�ۼ�����, I-��������
	DType = "W"

	'��������, yyyyMMdd
	SDate = "20190901"				

	'��������, yyyyMMdd
	EDate = "20191231"				

	' ���ۻ��°� �迭, �̱����� ��ü��ȸ, �������°� 3�ڸ� �迭, 2,3��° �ڸ� ���ϵ�ī�� ��밡��
	Dim State(2)
	State(0) = "2**"
	State(1) = "3**"

	'���� �����ڵ�迭 
	' - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	Dim ItemCode(6)
	ItemCode(0) = "121"
	ItemCode(1) = "122"
	ItemCode(2) = "123"
	ItemCode(3) = "124"
	ItemCode(4) = "125"
	ItemCode(5) = "126"
	
	'���Ĺ���, A-��������, D-��������
	Order = "D"			

	'������ ��ȣ
	Page = 1				

	'�������� �˻�����
	PerPage = 20		

	'�ŷ�ó����, �ŷ�ó��ȣ �Ǵ� ����ڵ�Ϲ�ȣ ����, �̱����� ��ü��ȸ
	SQuery = ""	

	On Error Resume Next

	Set result = m_StatementService.Search(testCorpNum, DType, SDate, EDate, State, ItemCode, Order, Page, PerPage, SQuery)

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
				<legend>���ڸ��� �����ȸ</legend>
					<% If code = 0 Then %>
					<ul>
						<li> code(���� �����ڵ�) : <%=result.code%></li>
						<li> total(�� �˻���� �Ǽ�) : <%=result.total%></li>
						<li> pageNum(������ ��ȣ) : <%=result.pageNum%></li>
						<li> perPage(�������� �˻�����) : <%=result.perPage%></li>
						<li> pageCount(������ ����) : <%=result.pageCount%></li>
						<li> message(���� �޽���) : <%=result.message%></li>
					</ul>
					
					<% For i=0 To UBound(result.list)-1 %>

						<fieldset class="fieldset2">
							<legend> ���ڸ��� ��ȸ��� [ <%=i+1%> / <%=UBound(result.list)%> ] </legend>
							<ul>
								<li> itemKey(������Ű) : <%=result.list(i).itemKey%></li>
								<li> itemCode(���������ڵ�) : <%=result.list(i).itemCode%></li>
								<li> stateCode(�����ڵ�) : <%=result.list(i).stateCode%></li>
								<li> taxType(��������) : <%=result.list(i).taxType%></li>
								<li> purposeType(����/û��) : <%=result.list(i).purposeType%></li>
								<li> writeDate(�ۼ�����) : <%=result.list(i).writeDate%></li>
								<li> senderCorpName(�߽��� ��ȣ) : <%=result.list(i).senderCorpName%></li>
								<li> senderCorpNum(�߽��� ����ڹ�ȣ) : <%=result.list(i).senderCorpNum%></li>
								<li> senderPrintYN(�߽��� �μ⿩��) : <%=result.list(i).senderPrintYN%></li>
								<li> receiverCorpName(������ ��ȣ) : <%=result.list(i).receiverCorpName%></li>
								<li> receiverCorpNum(������ ����ڹ�ȣ) : <%=result.list(i).receiverCorpNum%></li>
								<li> receiverPrintYN(������ �μ⿩��) : <%=result.list(i).receiverPrintYN%></li>
								<li> supplyCostTotal(���ް��� �հ�) : <%=result.list(i).supplyCostTotal%></li>
								<li> taxTotal(���� �հ�) : <%=result.list(i).taxTotal%></li>
								<li> issueDT(�����Ͻ�) : <%=result.list(i).issueDT%></li>
								<li> stateDT(���� �����Ͻ�) : <%=result.list(i).stateDT%></li>
								<li> openYN(���� ���� ����) : <%=result.list(i).openYN%></li>
								<li> openDT(���� �Ͻ�) : <%=result.list(i).openDT%></li>
								<li> stateMemo(���¸޸�) : <%=result.list(i).stateMemo%></li>
								<li> regDT(����Ͻ�) : <%=result.list(i).regDT%></li>
							</ul>
						</fieldset>
					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
			</fieldset>
		 </div>
	</body>
</html>