
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� �ѽ����� ������ ��ȸ�մϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'��������, yyyyMMdd
	SDate = "20180601"					

	'��������, yyyyMMdd
	EDate = "20180730"					
	
	' ���ۻ��°� �迭, 1-���, 2-����, 3-����, 4-���
	Dim State(4)
	State(0) = "1"
	State(1) = "2"
	State(2) = "3"
	State(3) = "4"
	
	'�������� �˻�����
	ReserveYN = False
	
	'������ȸ ����
	SenderOnlyYN = False

	'���Ĺ���, A-��������, D-��������
	Order = "D"

	'������ ��ȣ
	Page = 1

	'�������� �˻�����
	PerPage = 20
	
	On Error Resume Next

	Set result = m_FaxService.Search(testCorpNum, SDate, EDate, State, ReserveYN, SenderOnlyYN, Order, Page, PerPage)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�ѽ����� ���۳��� ��ȸ </legend>
					<ul>
						<li> code : <%=result.code%></li>
						<li> total : <%=result.total%></li>
						<li> pageNum : <%=result.pageNum%></li>
						<li> perPage : <%=result.perPage%></li>
						<li> pageCount : <%=result.pageCount%></li>
						<li> message : <%=result.message%></li>
					</ul>
				<% If code = 0 Then 
						For i=0 To UBound(result.list)-1
				%>
					<fieldset class="fieldset2">
							<legend> �ѽ� ���۰�� [ <%=i+1%> /  <%=UBound(result.list)%> ] </legend>
							<ul>
								<li>state (���ۻ��� �ڵ�) : <%=result.list(i).state%> </li>
								<li>result (���۰�� �ڵ�) : <%=result.list(i).result%> </li>
								<li>sendNum (�߽Ź�ȣ) : <%=result.list(i).sendNum%> </li>
								<li>senderName (�߽��ڸ�) : <%=result.list(i).senderName%> </li>
								<li>receiveNum (���Ź�ȣ) : <%=result.list(i).receiveNum%> </li>
								<li>receiveName (�����ڸ�) : <%=result.list(i).receiveName%> </li>
								<li>title (�ѽ� ����) : <%=result.list(i).title %> </li>
								<li>sendPageCnt (��������) : <%=result.list(i).sendPageCnt%></li>
								<li>successPageCnt (���� ��������) : <%=result.list(i).successPageCnt%></li>
								<li>failPageCnt (���� ��������) : <%=result.list(i).failPageCnt%></li>
								<li>refundPageCnt (ȯ�� ��������) : <%=result.list(i).refundPageCnt%></li>
								<li>cancelPageCnt (��� ��������) : <%=result.list(i).cancelPageCnt%></li>
								<li>reserveDT (����ð�) : <%=result.list(i).reserveDT%></li>
								<li>sendDT (�߼۽ð�) : <%=result.list(i).sendDT%></li>
								<li>receiptDT (���� �����ð�) : <%=result.list(i).receiptDT%></li>
								<li>fileNames (�������ϸ� �迭) : <%=result.list(i).fileNames%></li>
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
				<%	End If	%>

			</fieldset>
		 </div>
	</body>
</html>