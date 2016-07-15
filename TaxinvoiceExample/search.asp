<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "4108600477"		' [�ʼ�] �˺�ȸ�� ����ڹ�ȣ, "-" ����
	UserID = "innoposttest"
	KeyType= "SELL"						' [�ʼ�] �������� SELL(����), BUY(����), TRUSTEE(����Ź)
	DType = "R"								' [�ʼ�] �˻����� ����, R-�������, W-�ۼ�����, I-��������
	SDate = "20160501"					' [�ʼ�] ��������, yyyyMMdd
	EDate = "20160731"					' [�ʼ�] ��������, yyyyMMdd
	
	' ���ۻ��°� �迭, �̱����� ��ü��ȸ, �������°� 3�ڸ� �迭, 2,3��° �ڸ� ���ϵ�ī�� ��밡��
	Dim State(2)
	State(0) = "2**"
	State(1) = "3**"

	
	' �������� �迭, N-�Ϲݼ��ݰ�꼭, M-�������ݰ�꼭  �� ���ù迭
	Dim TIType(2)
	TIType(0) = "N"
	TIType(1) = "M"

	' �������� �迭, T-����, N-�鼼, Z-���� �� ���� �迭
	Dim TaxType(3)
	TaxType(0) = "T"
	TaxType(1) = "N"
	TaxType(2) = "Z"

	LateOnly = null		' �������࿩��,  null- ��ü��ȸ, False-�������� ��ȸ, True-��������� ��ȸ

	Order = "D"			' ���Ĺ���, A-��������, D-��������
	Page = 1				' ������ ��ȣ
	PerPage = 100		' �������� �˻�����, �ִ� 1000

	'��������ȣ ���������, S-����, B-����, T-��Ź
	TaxRegIDType = "S"

	'��������ȣ ����, ����-��ü��ȸ, 0-��������ȣ ����, 1-��������ȣ ����
	TaxRegIDYN = ""
	
	'��������ȣ, �޸�(",")�� �����Ͽ� ���� ex) "1234,0001"
	TaxRegID = ""

	Set result = m_TaxinvoiceService.Search(testCorpNum, KeyType, DType, SDate, EDate, State, TIType, TaxType, LateOnly, Order, Page, PerPage, TaxRegIDType, TaxRegIDYN, TaxRegID, UsreID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>���ݰ�꼭 �����ȸ ���</legend>
						<ul>
							<li> code : <%=result.code%></li>
							<li> total : <%=result.total%></li>
							<li> pageNum : <%=result.pageNum%></li>
							<li> perPage : <%=result.perPage%></li>
							<li> pageCount : <%=result.pageCount%></li>
							<li> message : <%=result.message%></li>
						</ul>

				<%
					If code = 0 Then
						For i=0 To UBound(result.list) -1
				%>
							<fieldset class="fieldset2">					
								<legend>  Search.List [ <%=i+1%> / <%=UBound(result.list)%> ]</legend>
									<ul>
										<li> itemKey :  <%=result.list(i).itemKey%> </li>
										<li> stateCode :  <%=result.list(i).stateCode%> </li>
										<li> taxType :  <%=result.list(i).taxType%> </li>
										<li> purposeType :  <%=result.list(i).purposeType%> </li>
										<li> issueType :  <%=result.list(i).issueType %> </li>
										<li> writeDate :  <%=result.list(i).writeDate%> </li>
										<li> invoicerCorpName :  <%=result.list(i).invoicerCorpName%> </li>
										<li> invoicerCorpNum :  <%=result.list(i).invoicerCorpNum%> </li>
										<li> invoicerMgtKey :  <%=result.list(i).invoicerMgtKey%> </li>
										<li> invoicerPrintYN :  <%=result.list(i).invoicerPrintYN%> </li>
										<li> invoiceeCorpName :  <%=result.list(i).invoiceeCorpName%> </li>
										<li> invoiceeCorpNum :  <%=result.list(i).invoiceeCorpNum%> </li>
										<li> invoiceeMgtKey :  <%=result.list(i).invoiceeMgtKey%> </li>
										<li> invoiceePrintYN :  <%=result.list(i).invoiceePrintYN%> </li>
										<li> trusteeCorpName :  <%=result.list(i).trusteeCorpName%> </li>
										<li> trusteeCorpNum :  <%=result.list(i).trusteeCorpName%> </li>
										<li> trusteeMgtKey :  <%=result.list(i).trusteeMgtKey%> </li> 
										<li> trusteePrintYN :  <%=result.list(i).trusteePrintYN%> </li>
										<li> supplyCostTotal :  <%=result.list(i).supplyCostTotal%> </li>
										<li> taxTotal :  <%=result.list(i).taxTotal%> </li>
										<li> issueDT :  <%=result.list(i).issueDT%> </li>
										<li> preIssueDT :  <%=result.list(i).preIssueDT%> </li>
										<li> stateDT :  <%=result.list(i).stateDT%> </li>
										<li> openYN :  <%=result.list(i).openYN%> </li>
										<li> openDT :  <%=result.list(i).openDT%> </li>
										<li> ntsresult :  <%=result.list(i).ntsresult%> </li>
										<li> ntsconfirmNum :  <%=result.list(i).ntsconfirmNum %> </li>
										<li> ntssendDT :  <%=result.list(i).ntssendDT%> </li>
										<li> ntsresultDT :  <%=result.list(i).ntsresultDT%> </li>
										<li> ntssendErrCode :  <%=result.list(i).ntssendErrCode%> </li>
										<li> stateMemo :  <%=result.list(i).stateMemo%> </li>
										<li> regDT :  <%=result.list(i).regDT%> </li>
										<li> lateIssueYN :  <%=result.list(i).lateIssueYN%> </li>
									</ul>
								</fieldset>
				<%
						Next
					Else
				%>
				</fieldset>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>
