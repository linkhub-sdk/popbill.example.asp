
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
	' - �ִ� �˻��Ⱓ : 6���� �̳�
	' - https://docs.popbill.com/message/asp/api#Search
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'��������
	SDate = "20210501"

	'��������
	EDate = "20210601"					
	
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
	
	'��ȸ �˻���.
	'���� ���۽� �Է��� �߽��ڸ� �Ǵ� �����ڸ� ����.
	'��ȸ �˻�� ������ �߽��ڸ� �Ǵ� �����ڸ��� �˻��մϴ�.
	QString = ""
	
	' On Error Resume Next

	Set resultObj = m_MessageService.Search(testCorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)
	
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
				<legend>���ڸ޼��� ���۳��� ��ȸ </legend>
				<ul>
						<li> code (�����ڵ�) : <%=resultObj.code%></li>
						<li> total (�� �˻���� �Ǽ�) : <%=resultObj.total%></li>
						<li> pageNum (������ ��ȣ) : <%=resultObj.pageNum%></li>
						<li> perPage (�������� ��ϰ���) : <%=resultObj.perPage%></li>
						<li> pageCount (������ ����) : <%=resultObj.pageCount%></li>
						<li> message (����޽���) : <%=resultObj.message%></li>
				</ul>
					<% If code = 0 Then
						For i=0 To UBound(resultObj.list) -1
					%>

						<fieldset class="fieldset2">
							<legend> ���ڸ޽��� ���۰�� [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
							<ul>

								<li>state (���ۻ��� �ڵ�) : <%=resultObj.list(i).state%> </li>
								<li>result (���۰�� �ڵ�) : <%=resultObj.list(i).result%> </li>
								<li>subject (�޽��� ����) : <%=resultObj.list(i).subject%> </li>
								<li>content (�޽��� ����) : <%=resultObj.list(i).content%> </li>
								<li>type (�޽��� ����) : <%=resultObj.list(i).msgType%> </li>
								<li>sendnum (�߽Ź�ȣ) : <%=resultObj.list(i).sendnum%> </li>
								<li>senderName (�߽��ڸ�) : <%=resultObj.list(i).senderName%> </li>
								<li>receiveNum (���Ź�ȣ) : <%=resultObj.list(i).receiveNum%> </li>
								<li>receiveName (�����ڸ�) : <%=resultObj.list(i).receiveName%> </li>
								<li>receiptDT (�����Ͻ�) : <%=resultObj.list(i).receiptDT%> </li>
								<li>sendDT (�����Ͻ�) : <%=resultObj.list(i).sendDT%> </li>
								<li>resultDT (���۰�� �����Ͻ�) : <%=resultObj.list(i).resultDT%> </li>
								<li>reserveDT (�����Ͻ�) : <%=resultObj.list(i).reserveDT%> </li>
								<li>tranNet (����ó�� �̵���Ż��) : <%=resultObj.list(i).tranNet%> </li>
								<li>receiptNum (������ȣ) : <%=resultObj.list(i).receiptNum%> </li>
								<li>requestNum (��û��ȣ) : <%=resultObj.list(i).requestNum%> </li>
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