
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˻������� ����Ͽ� īī���� ���۳��� ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 2����)
	' - īī���� �����Ͻ÷κ��� 6���� �̳� �����Ǹ� ��ȸ�� �� �ֽ��ϴ�.
	' - https://docs.popbill.com/kakao/asp/api#Search
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'��������
	SDate = "20210601"

	'��������
	EDate = "20210624"					
	
	'���ۻ��°� �迭, 0-���۴��, 1-������, 2-���ۼ���, 3-��ü���� ����, 4-���۽���, 5-�������
	Dim State(6)
	State(0) = "0"
	State(1) = "1"
	State(2) = "2"
	State(3) = "3"
	State(4) = "4"
	State(5) = "5"

	'�˻���� �迭, ATS-�˸���, FTS-ģ���� �ؽ�Ʈ, FMS-ģ���� �̹���
	Dim Item(3)
	Item(0) = "ATS"
	Item(1) = "FTS"
	Item(2) = "FMS"

	' �������ۿ���, ����-��ü��ȸ, 1-�������۰� ��ȸ, 0-������۰� ��ȸ
	ReserveYN = ""	

	' ������ȸ���� (True-������ȸ / False-��ü��ȸ)
	SenderYN = False		

	' ���Ĺ���, D-��������, A-��������
	Order = "D"				

	' ������ ��ȣ 
	Page = 1					

	PerPage = 30			

	'��ȸ �˻���.
	'īī���� ���۽� �Է��� �߽��ڸ� �Ǵ� �����ڸ� ����.
	'��ȸ �˻�� ������ �߽��ڸ� �Ǵ� �����ڸ��� �˻��մϴ�.
	QString = ""

	On Error Resume Next

	Set resultObj = m_KakaoService.Search(testCorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)

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
				
					<legend>īī���� ���۳��� ��ȸ </legend>
					<ul>
					<% If code = 0 Then %>
							<li> code (�����ڵ�) : <%=resultObj.code%></li>
							<li> message (����޽���) : <%=resultObj.message%></li>
							<li> total (�� �˻���� �Ǽ�) : <%=resultObj.total%></li>
							<li> pageNum (������ ��ȣ) : <%=resultObj.pageNum%></li>
							<li> pageCount (������ ����) : <%=resultObj.pageCount%></li>
							<li> perPage (�������� �˻�����) : <%=resultObj.perPage%></li>
					</ul>
						<% 
							For i=0 To UBound(resultObj.list) -1
						%>
							<fieldset class="fieldset2">
								<legend> īī���� ���۰�� [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
								<ul>
									<li>state (���ۻ��� �ڵ�) : <%=resultObj.list(i).state%> </li>
									<li>sendDT (�����Ͻ�) : <%=resultObj.list(i).sendDT%> </li>
									<li>result (���۰�� �ڵ�) : <%=resultObj.list(i).result%> </li>
									<li>resultDT (���۰�� �����Ͻ�) : <%=resultObj.list(i).resultDT%> </li>
									<li>contentType (īī���� ����) : <%=resultObj.list(i).contentType%> </li>
									<li>receiveNum (���Ź�ȣ) : <%=resultObj.list(i).receiveNum%> </li>
									<li>receiveName (�����ڸ�) : <%=resultObj.list(i).receiveName%> </li>
									<li>content (�˸���/ģ���� ����) : <%=resultObj.list(i).content%> </li>
									<li>altContentType (��ü���� ����Ÿ��) : <%=resultObj.list(i).altContentType%> </li>
									<li>altSendDT (��ü���� �����Ͻ�) : <%=resultObj.list(i).altSendDT%> </li>
									<li>altResult (��ü���� ���۰�� �ڵ�) : <%=resultObj.list(i).altResult%> </li>
									<li>altResultDT (��ü���� ���۰�� �����Ͻ�) : <%=resultObj.list(i).altResultDT%> </li>
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