<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� ��������� ��ȸ�մϴ�.
	' - �����׸� ���� ������ "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���]
	'   > 3.3.1. Search (���� ��� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "6798700433"		

	'�˺�ȸ�� ���̵�
	UserID = ""
	
	'���� ��û(requestJob) �� ��ȯ���� �۾����̵�(jobID)
	JobID = "018100317000000001"

	'���ݿ����� �迭 N-�Ϲ����ݿ�����, C-������ݿ�����
	Dim TradeType(2) 
	TradeType(0) = "N"
	TradeType(1) = "C"

	'�ŷ��뵵 �迭, P-�ҵ������, C-����������
	Dim TradeUsage(2)
	TradeUsage(0) = "P"
	TradeUsage(1) = "C"

	'������ ��ȣ 
	Page  = 1

	'�������� ��ϰ���
	PerPage = 10

	'���Ĺ���, D-��������, A-��������
	Order = "D"

	On Error Resume Next

	Set result = m_HTCashbillService.Search(testCorpNum, JobID, TradeType, TradeUsage, Page, PerPage, Order, UserID)

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
				<legend>���� ��� ��ȸ</legend>
				<%
					If code = 0 Then
				%>
					<ul>
						<li> code (�����ڵ�) : <%=result.code%> </li>
						<li> message  (����޽���) : <%=result.message%> </li>
						<li> total (�� �˻���� �Ǽ�) : <%=result.total%> </li>
						<li> perPage (�������� �˻�����) : <%=result.perPage%> </li>
						<li> pageNum (������ ��ȣ) : <%=result.pageNum%> </li>
						<li> pageCount (������ ����) : <%=result.pageCount%> </li>
					</ul>

				<%
					For i=0 To UBound(result.list) -1 
				%>
					<fieldset class="fieldset2">					
						<legend>ListActiveJob [ <%=i+1%> / <%= UBound(result.list) %> ] </legend>
							<ul>										
								<li> ntsconfirmNum (����û���ι�ȣ) : <%= result.list(i).ntsconfirmNum %></li>
								<li> tradeDT (�ŷ�����) : <%= result.list(i).tradeDT %></li>
								<li> tradeDT (�ŷ��Ͻ�) : <%= result.list(i).tradeDT %></li>
								<li> tradeType (��������) : <%= result.list(i).tradeType %></li>
								<li> tradeUsage (�ŷ�����) : <%= result.list(i).tradeUsage %></li>
								<li> totalAmount (�ŷ��ݾ�) : <%= result.list(i).totalAmount %></li>
								<li> supplyCost (���ް���) : <%= result.list(i).supplyCost %></li>
								<li> tax (�ΰ���) : <%= result.list(i).tax %></li>
								<li> serviceFee (�����) : <%= result.list(i).serviceFee %></li>
								<li> invoiceType (����/����) : <%= result.list(i).invoiceType %></li>
								<li> franchiseCorpNum (������ ����ڹ�ȣ) : <%= result.list(i).franchiseCorpNum %></li>
								<li> franchiseCorpName (������ ��ȣ) : <%= result.list(i).franchiseCorpName %></li>
								<li> franchiseCorpType (������ ���������) : <%= result.list(i).franchiseCorpType %></li>
								<li> identityNum (�ĺ���ȣ) : <%= result.list(i).identityNum %></li>
								<li> identityNumType (�ĺ���ȣ����) : <%= result.list(i).identityNumType %></li>
								<li> customerName (����) : <%= result.list(i).customerName %></li>
								<li> cardOwnerName (ī������ڸ�) : <%= result.list(i).cardOwnerName %></li>
								<li> deductionType (��������) : <%= result.list(i).deductionType %></li>						
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

