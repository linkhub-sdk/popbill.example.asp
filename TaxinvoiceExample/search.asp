<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� ���ݰ�꼭 ����� ��ȸ�մϴ�.
	' - �����׸� ���� �ڼ��� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] >
	'   4.2. (����)��꼭 �������� ����" �� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
	testCorpNum = "1234567890"
	
	' �˺�ȸ�� ���̵�
	UserID = "testkorea"

	' [�ʼ�] �������� SELL(����), BUY(����), TRUSTEE(����Ź)
	KeyType = "SELL"

	' [�ʼ�] �˻����� ����, R-�������, W-�ۼ�����, I-��������
	DType = "W"
	
	' [�ʼ�] ��������, yyyyMMdd
	SDate = "20181201"

	' [�ʼ�] ��������, yyyyMMdd
	EDate = "20190103"
	
	' ���ۻ��°� �迭, �̱����� ��ü��ȸ, �������°� 3�ڸ� �迭, 2,3��° �ڸ� ���ϵ�ī�� ��밡��
	Dim State(2)
	State(0) = "3**"
	State(1) = "6**"

	
	' �������� �迭, N-�Ϲݼ��ݰ�꼭, M-�������ݰ�꼭  �� ���ù迭
	Dim TIType(2)
	TIType(0) = "N"
	TIType(1) = "M"

	' �������� �迭, T-����, N-�鼼, Z-���� �� ���� �迭
	Dim TaxType(3)
	TaxType(0) = "T"
	TaxType(1) = "N"
	TaxType(2) = "Z"

	' �������� �迭, T-����, N-�鼼, Z-���� �� ���� �迭
	Dim IssueType(3)
	IssueType(0) = "N"
	IssueType(1) = "R"
	IssueType(2) = "T"

	' �������࿩��,  null- ��ü��ȸ, False-�������� ��ȸ, True-��������� ��ȸ
	LateOnly = null		

	' ���Ĺ���, A-��������, D-��������
	Order = "D"

	' ������ ��ȣ
	Page = 1

	' �������� �˻�����, �ִ� 1000
	PerPage = 5

	'��������ȣ ���������, S-����, B-����, T-��Ź
	TaxRegIDType = "S"

	'��������ȣ ����, ����-��ü��ȸ, 0-��������ȣ ����, 1-��������ȣ ����
	TaxRegIDYN = ""
	
	'��������ȣ, �޸�(",")�� �����Ͽ� ���� ex) "1234,0001"
	TaxRegID = ""

	'�ŷ�ó ����, �ŷ�ó ��ȣ �Ǵ� ����ڵ�Ϲ�ȣ ����, ����ó���� ��ü��ȸ
	QString = ""

	'�������� ��ȸ����, ����-��ü��ȸ, 0-�Ϲݹ��� ��ȸ, 1-�������� ��ȸ
	InterOPYN = ""

	On Error Resume Next

	Set result = m_TaxinvoiceService.Search(testCorpNum, KeyType, DType, SDate, EDate, State, _ 
						TIType, TaxType, IssueType, LateOnly, Order, Page, PerPage, TaxRegIDType, TaxRegIDYN, _
						TaxRegID, QString, InterOPYN, UsreID)

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
				<%
					If code = 0 Then
				%>
						<legend>���ݰ�꼭 �����ȸ</legend>
						<ul>
							<li> code (�����ڵ�) : <%=result.code%></li>
							<li> message (����޽���) : <%=result.message%></li>
							<li> total (�� �˻���� �Ǽ�) : <%=result.total%></li>
							<li> pageNum (������ ��ȣ) : <%=result.pageNum%></li>
							<li> perPage (�������� ��ϰ���) : <%=result.perPage%></li>
							<li> pageCount (������ ����) : <%=result.pageCount%></li>
						</ul>
						<%
							For i=0 To UBound(result.list) -1
						%>
							<fieldset class="fieldset2">					
								<legend>  ���ݰ�꼭 ����/������� [ <%=i+1%> / <%=UBound(result.list)%> ]</legend>
									<ul>
										<li> itemKey (���ݰ�꼭 ������Ű) :  <%=result.list(i).itemKey%> </li>
										<li> stateCode (�����ڵ�) :  <%=result.list(i).stateCode%> </li>
										<li> taxType (��������) :  <%=result.list(i).taxType%> </li>
										<li> purposeType (����/û��) :  <%=result.list(i).purposeType%> </li>
										<li> issueType (��������) :  <%=result.list(i).issueType %> </li>
										<li> writeDate (�ۼ�����) :  <%=result.list(i).writeDate%> </li>

										<li> invoicerCorpName (������ ��ȣ) :  <%=result.list(i).invoicerCorpName%> </li>
										<li> invoicerCorpNum (������ ����ڹ�ȣ) :  <%=result.list(i).invoicerCorpNum%> </li>
										<li> invoicerMgtKey (������ ����������ȣ) :  <%=result.list(i).invoicerMgtKey%> </li>
										<li> invoicerPrintYN (������ �μ⿩��) :  <%=result.list(i).invoicerPrintYN%> </li>
										
										<li> invoiceeCorpName (���޹޴��� ��ȣ) :  <%=result.list(i).invoiceeCorpName%> </li>
										<li> invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) :  <%=result.list(i).invoiceeCorpNum%> </li>
										<li> invoiceeMgtKey (���޹޴��� ����������ȣ) :  <%=result.list(i).invoiceeMgtKey%> </li>
										<li> invoiceePrintYN (���޹޴��� �μ⿩��) :  <%=result.list(i).invoiceePrintYN%> </li>
										<li> closeDownState (���޹޴��� ���������) :  <%=result.list(i).closeDownState%> </li>
										<li> closeDownStateDate (���޹޴��� ���������) :  <%=result.list(i).closeDownStateDate%> </li>

										<li> interOPYN (�������� ����) :  <%=result.list(i).interOPYN%> </li>
										<li> supplyCostTotal (���ް��� �հ�) :  <%=result.list(i).supplyCostTotal%> </li>
										<li> taxTotal (���� �հ�) :  <%=result.list(i).taxTotal%> </li>
										<li> issueDT (�����Ͻ�) :  <%=result.list(i).issueDT%> </li>

										<li> stateDT (���� �����Ͻ�) :  <%=result.list(i).stateDT%> </li>
										<li> openYN (���� ����) :  <%=result.list(i).openYN%> </li>
										<li> openDT (���� �Ͻ�) :  <%=result.list(i).openDT%> </li>
										<li> ntsresult (����û ���۰��) :  <%=result.list(i).ntsresult%> </li>
										<li> ntsconfirmNum (����û ���ι�ȣ) :  <%=result.list(i).ntsconfirmNum %> </li>
										<li> ntssendDT (����û �����Ͻ�) :  <%=result.list(i).ntssendDT%> </li>
										<li> ntsresultDT (����û ��� �����Ͻ�) :  <%=result.list(i).ntsresultDT%> </li>
										<li> ntssendErrCode (���۽��� �����ڵ�) :  <%=result.list(i).ntssendErrCode%> </li>

										<li> stateMemo (���¸޸�) :  <%=result.list(i).stateMemo%> </li>
										<li> regDT (����Ͻ�) :  <%=result.list(i).regDT%> </li>
										<li> lateIssueYN (�������� ����) :  <%=result.list(i).lateIssueYN%> </li>
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
