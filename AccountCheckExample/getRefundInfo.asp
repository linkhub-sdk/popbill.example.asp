<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
    	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    	<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    	<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"-->
<%
	'**************************************************************
	' ȯ�� ��û ������ Ȯ���մϴ�.
	' - https://developers.popbill.com/reference/accountcheck/asp/api/point#GetRefundInfo
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	CorpNum = "1234567890"

	'ȯ�� �ڵ�
	refundCode = "023040000017"

	'�˺�ȸ�� ���̵�
	UserID = "testkorea"

	On Error Resume Next

	Set result = m_AccountCheckService.GetRefundInfo(CorpNum, refundCode, UserID)

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
            	<legend>ȯ�� ��û ���� ��ȸ</legend>
            	<%
                	If code = 0 Then
            	%>
                	<fieldset class="fieldset2">
                    	<legend> GetRefundInfo </legend>
                        	<ul>
                            	<li> reqDT (��û �Ͻ�) : <%=result.reqDT%></li>
                            	<li> requestPoint (ȯ�� ��û����Ʈ) : <%=result.requestPoint%></li>
                            	<li> accountBank (ȯ�Ұ��� �����) : <%=result.accountBank%></li>
                            	<li> accountNum (ȯ�Ұ��¹�ȣ) : <%=result.accountNum%></li>
                            	<li> accountName (ȯ�Ұ��� �����ָ�) : <%=result.accountName%></li>
                            	<li> state (����) : <%=result.state%></li>
                            	<li> reason (ȯ�һ���) : <%=result.reason%></li>
                        	</ul>
                    	</fieldset>
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
