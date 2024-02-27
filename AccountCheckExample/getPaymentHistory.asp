<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
    	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    	<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    	<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"-->
<%
	'**************************************************************
	' ����ȸ���� ����Ʈ ���������� Ȯ���մϴ�.
	' - https://developers.popbill.com/reference/accountcheck/asp/api/point#GetPaymentHistory
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	CorpNum = "1234567890"

	' ��ȸ �Ⱓ�� �������� (���� : yyyyMMdd)
	SDate = "20230401"

	' ��ȸ �Ⱓ�� �������� (���� : yyyyMMdd)
	EDate = "20230530"

	' ��� ��������ȣ
	Page = 1

	' �������� ǥ���� ��� ����
	PerPage = 500

	'�˺�ȸ�� ���̵�
	UserID = "testkorea"

	On Error Resume Next

	Set result = m_AccountCheckService.GetPaymentHistory(CorpNum, SDate, EDate, Page, PerPage, UserID)

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
            	<legend>����ȸ�� ����Ʈ �������� Ȯ��</legend>
            	<%
                	If code = 0 Then
            	%>
            	<ul>
                	<li> code (ȯ�� ���� ����Ʈ) : <%=result.code%></li>
                	<li> total (ȯ�� ���� ����Ʈ) : <%=result.total%></li>
                	<li> perPage (ȯ�� ���� ����Ʈ) : <%=result.perPage%></li>
                	<li> pageNum (ȯ�� ���� ����Ʈ) : <%=result.pageNum%></li>
                	<li> pageCount (ȯ�� ���� ����Ʈ) : <%=result.pageCount%></li>
            	</ul>
            	<%
                	Dim i
                	For i = 0 To UBound(result.list) -1
            	%>
                	<fieldset class="fieldset2">
                    	<legend> PaymentHistory [ <%= i+1%> / <%=UBound(result.list)%>]</legend>
                    	<ul>
                        	<li>productType (���� ����) : <%= result.list(i).productType %></li>
                        	<li>productName (���� ��ǰ��) : <%= result.list(i).productName %></li>
                        	<li>settleType (��������) : <%= result.list(i).settleType %></li>
                        	<li>settlerName (����ڸ�) : <%= result.list(i).settlerName %></li>
                        	<li>settlerEmail (����ڸ���) : <%= result.list(i).settlerEmail %></li>
                        	<li>settleCost (�����ݾ�) : <%= result.list(i).settleCost %></li>
                        	<li>settlePoint (��������Ʈ) : <%= result.list(i).settlePoint %></li>
                        	<li>settleState (��������) : <%= result.list(i).settleState %></li>
                        	<li>regDT (����Ͻ� ) : <%= result.list(i).regDT %></li>
                        	<li>stateDT (�����Ͻ� ) : <%= result.list(i).stateDT %></li>
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
