<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �뷮�� ���ݰ�꼭 ����/��� ������ Ȯ���մϴ�. (�ִ� 1000��)
	' - https://docs.popbill.com/taxinvoice/asp/api#GetInfos
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
	KeyType = "SELL"

	' �˺�ȸ�� ���̵�
	UserID = "testkorea"

	' ���ݰ�꼭 ������ȣ �迭, �ִ� 1000��
	Dim MgtKeyList(3) 
	MgtKeyList(0) = "20190103-001"
	MgtKeyList(1) = "20190103-002"
	MgtKeyList(2) = "20190103-003"
	
	On Error Resume Next

	Set result = m_TaxinvoiceService.GetInfos(testCorpNum, KeyType, MgtKeyList, UserID)

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
				<legend>���ݰ�꼭 ����/��� ���� Ȯ�� - �뷮</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend> TaxinvoiceResult : <%=i+1%> </legend>
									<ul>
										<li> itemKey (���ݰ�꼭 ������Ű) :  <%=result.Item(i).itemKey%> </li>
										<li> stateCode (�����ڵ�) :  <%=result.Item(i).stateCode%> </li>
										<li> taxType (��������) :  <%=result.Item(i).taxType%> </li>
										<li> purposeType (����/û��) :  <%=result.Item(i).purposeType%> </li>
										<li> issueType (��������) :  <%=result.Item(i).issueType %> </li>
										<li> writeDate (�ۼ�����) :  <%=result.Item(i).writeDate%> </li>

										<li> invoicerCorpName (������ ��ȣ) :  <%=result.Item(i).invoicerCorpName%> </li>
										<li> invoicerCorpNum (������ ����ڹ�ȣ) :  <%=result.Item(i).invoicerCorpNum%> </li>
										<li> invoicerMgtKey (������ ������ȣ) :  <%=result.Item(i).invoicerMgtKey%> </li>
										<li> invoicerPrintYN (������ �μ⿩��) :  <%=result.Item(i).invoicerPrintYN%> </li>

										<li> invoiceeCorpName (���޹޴��� ��ȣ) :  <%=result.Item(i).invoiceeCorpName%> </li>
										<li> invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) :  <%=result.Item(i).invoiceeCorpNum%> </li>
										<li> invoiceeMgtKey (���޹޴��� ������ȣ) :  <%=result.Item(i).invoiceeMgtKey%> </li>
										<li> invoiceePrintYN (���޹޴��� �μ⿩��) :  <%=result.Item(i).invoiceePrintYN%> </li>
										<li> closeDownState (���޹޴��� ���������) :  <%=result.Item(i).closeDownState%> </li>
										<li> closeDownStateDate (���޹޴��� ������Ͻ�) :  <%=result.Item(i).closeDownStateDate%> </li>
										<li> interOPYN (�������� ����) :  <%=result.Item(i).interOPYN%> </li>

										<li> supplyCostTotal (���ް��� �հ�) :  <%=result.Item(i).supplyCostTotal%> </li>
										<li> taxTotal (���� �հ�) :  <%=result.Item(i).taxTotal%> </li>
										<li> issueDT (���� �Ͻ�) :  <%=result.Item(i).issueDT%> </li>

										<li> stateDT (���� �����Ͻ�) :  <%=result.Item(i).stateDT%> </li>
										<li> openYN (���� ����) :  <%=result.Item(i).openYN%> </li>
										<li> openDT (���� �Ͻ�) :  <%=result.Item(i).openDT%> </li>
										<li> ntsresult (����û ���۰��) :  <%=result.Item(i).ntsresult%> </li>
										<li> ntsconfirmNum (����û ���ι�ȣ) :  <%=result.Item(i).ntsconfirmNum %> </li>
										<li> ntssendDT (����û �����Ͻ�) :  <%=result.Item(i).ntssendDT%> </li>
										<li> ntsresultDT (����û ��� �����Ͻ�) :  <%=result.Item(i).ntsresultDT%> </li>
										<li> ntssendErrCode (���۽��� ����) :  <%=result.Item(i).ntssendErrCode%> </li>
										<li> stateMemo (���¸޸�) :  <%=result.Item(i).stateMemo%> </li>
										<li> regDT (�ӽ����� �Ͻ�) :  <%=result.Item(i).regDT%> </li>
										<li> lateIssueYN (�������� ����) :  <%=result.Item(i).lateIssueYN%> </li>
									</ul>
								</fieldset>
				<%
						Next
					Else
				%>
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
