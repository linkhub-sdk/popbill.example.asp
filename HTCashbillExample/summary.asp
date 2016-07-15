<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	UserID = "testkorea"				'�˺�ȸ�� ���̵�
	
	'���� ��û(requestJob) �� ��ȯ���� �۾����̵�(jobID)
	JobID = "016071511000000009"

	'���ݿ����� �迭 N-�Ϲ����ݿ�����, C-������ݿ�����
	Dim TradeType(2) 
	TradeType(0) = "N"
	TradeType(1) = ""

	'�ŷ��뵵 �迭, P-�ҵ������, C-����������
	Dim TradeUsage(2)
	TradeUsage(0) = "P"
	TradeUsage(1) = ""
	
	Set result = m_HTCashbillService.Summary(testCorpNum, JobID, TradeType, TradeUsage, UserID)

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
				<legend>���� ��� ��ȸ</legend>
				<%
					If code = 0 Then
				%>
					<ul>
						<li> count (���� ��� �Ǽ�) : <%=result.count%> </li>
						<li> supplyCostTotal (���ް��� �հ�) : <%=result.supplyCostTotal%> </li>
						<li> taxTotal (���� �հ�) : <%=result.taxTotal%> </li>
						<li> serviceFeeTotal (����� �հ�) : <%=result.serviceFeeTotal%> </li>
						<li> amountTotal (�հ� �ݾ�) : <%=result.amountTotal%> </li>
					</ul>
				<%
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