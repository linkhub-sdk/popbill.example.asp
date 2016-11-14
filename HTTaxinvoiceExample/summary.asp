<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� ���� ��� ��������� ��ȸ�մϴ�.
	' - �����׸� ���� ������ "[Ȩ�ý� ����(����)��꼭 ���� API �����Ŵ���]
	'   > 3.3.2. Summary (���� ��� ������� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "6798700433"	
	
	'�˺�ȸ�� ���̵�
	UserID = "testkorea_linkhub"	
	
	'���� ��û(requestJob) �� ��ȯ���� �۾����̵�(jobID)
	JobID = "016111416000000024"

	'�������� �迭, N-�Ϲ� ���ڼ��ݰ�꼭, M-���� ���ڼ��ݰ�꼭 
	Dim TIType(2) 
	TIType(0) = "N"
	TIType(1) = "M"

	'�������� �迭,  T-����, N-�鼼, Z-����
	Dim TaxType(3)
	TaxType(0) = "T"
	TaxType(1) = "N"
	TaxType(2) = "Z"
	
	'����/û�� �迭, R-����, C-û��, N-����
	Dim PurposeType(3)
	PurposeType(0) = "R"
	PurposeType(1) = "C"
	PurposeType(2) = "N"

	'������� ����, ����-��ü��ȸ, 0-��������ȣ ����, 1-��������ȣ ��ȸ
	TaxRegIDYN = ""

	'������� ����� ����, S-������, B-���޹޴���, T-��Ź��
	TaxRegIDType = "S"

	'��������ȣ, �޸�(",")�� �����Ͽ� ���� ex) 1234,1001
	TaxRegID = ""
	
	On Error Resume Next

	Set result = m_HTTaxinvoiceService.Summary(testCorpNum, JobID, TIType, TaxType,  _
							PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, UserID)

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
						<li> count (���� ��� �Ǽ�) : <%=result.count%> </li>
						<li> supplyCostTotal (���ް��� �հ�) : <%=result.supplyCostTotal%> </li>
						<li> taxTotal (���� �հ�) : <%=result.taxTotal%> </li>
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