<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���� �۾��� �Ϸ�� ������ �ŷ������� ��ȸ�մϴ�.
	' - https://docs.popbill.com/easyfinbank/asp/api#Search
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	

	'�˺�ȸ�� ���̵�
	UserID = "testkorea"
	
	'���� ��û(requestJob) �� ��ȯ���� �۾����̵�(jobID)
	JobID = "019123114000000010"

	'�ŷ����� �迭, I-�Ա�, O-���
	Dim TradeType(2) 
	TradeType(0) = "I"
	TradeType(1) = "O"

	'��ȸ �˻���, �Ա�/��ݾ�, �޸�, ���� like �˻�
	SearchString = ""

	'������ ��ȣ 
	Page  = 1

	'�������� ��ϰ���
	PerPage = 10

	'���Ĺ���, D-��������, A-��������
	Order = "D"

	On Error Resume Next

	Set result = m_EasyFinBankService.Search(testCorpNum, JobID, TradeType, SearchString, _	
								Page, PerPage, Order, UserID)

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
						<legend>�ŷ����� ���� [ <%=i+1%> / <%= UBound(result.list) %> ] </legend>
							<ul>								
								<li> tid (�ŷ����� ���̵�) : <%= result.list(i).tid %></li>
								<li> trdate (�ŷ�����) : <%= result.list(i).trdate %></li>
								<li> trserial (�ŷ����ں� �ŷ����� ����) : <%= result.list(i).trserial %></li>
								<li> trdt (�ŷ��Ͻ�) : <%= result.list(i).trdt %></li>
								<li> accIn (�Աݾ�) : <%= result.list(i).accIn %></li>
								<li> accOut (��ݾ�) : <%= result.list(i).accOut %></li>
								<li> balance (�ܾ�) : <%= result.list(i).balance %></li>
								<li> remark1 (���1) : <%= result.list(i).remark1 %></li>
								<li> remark2 (���2) : <%= result.list(i).remark2 %></li>
								<li> remark3 (���3) : <%= result.list(i).remark3 %></li>
								<li> memo (�޸�) : <%= result.list(i).memo %></li>
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

