
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� �������۳��� ����� ��ȸ�մϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'��������
	SDate = "20151001"

	'��������
	EDate = "20160127"					
	
	'���ۻ��°� �迭, 1-���, 2-����, 3-����, 4-���
	Dim State(4)
	State(0) = "1"
	State(1) = "2"
	State(2) = "3"
	State(3) = "4"

	'�˻���� �迭, SMS., LMS, MMS
	Dim Item(3)
	Item(0) = "SMS"
	Item(1) = "LMS"
	Item(2) = "MMS"

	' �������ۿ���
	ReserveYN = False	

	' ������ȸ���� 
	SenderYN = False		

	' ���Ĺ���, D-��������, A-��������
	Order = "D"				

	' ������ ��ȣ 
	Page = 1					

	' �������� �˻����� 
	PerPage = 30			
	
	On Error Resume Next

	Set result = m_MessageService.Search(testCorpNum, SDate, EDate, Item, ReserveYN, SenderYN, Order, Page, PerPage)
	
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
				<legend>���ڸ޽��� ���۳��� ��ȸ </legend>
				<ul>
						<li> code : <%=result.code%></li>
						<li> total : <%=result.total%></li>
						<li> pageNum : <%=result.pageNum%></li>
						<li> perPage : <%=result.perPage%></li>
						<li> pageCount : <%=result.pageCount%></li>
						<li> message : <%=result.message%></li>
				</ul>
					<% If code = 0 Then
						For i=0 To UBound(result.list) -1
					%>

						<fieldset class="fieldset2">
							<legend> ���ڸ޽��� ���۰�� [ <%=i+1%> / <%= UBound(result.list)%> ] </legend>
							<ul>
								<li>state : <%=result.list(i).state%> </li>
								<li>resultDT : <%=result.list(i).resultDT%> </li>
								<li>sendResult : <%=result.list(i).sendResult%> </li>
								<li>subject : <%=result.list(i).subject%> </li>
								<li>content : <%=result.list(i).content%> </li>
								<li>type : <%=result.list(i).msgType%> </li>
								<li>sendnum: <%=result.list(i).sendnum%> </li>
								<li>senderName: <%=result.list(i).senderName%> </li>
								<li>receiveNum : <%=result.list(i).receiveNum%> </li>
								<li>receiveName : <%=result.list(i).receiveName%> </li>
								<li>reserveDT : <%=result.list(i).reserveDT%> </li>
								<li>sendDT : <%=result.list(i).sendDT%> </li>
								<li>tranNet : <%=result.list(i).tranNet%> </li>
								<li>receiptDT : <%=result.list(i).receiptDT%> </li>
							</ul>
						</fieldset>

					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>

			</fieldset>
		 </div>
	</body>
</html>