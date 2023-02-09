<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' Ȩ�ý����� ������ ���� �˺��� ��ϵ� ���ݿ����� �ڷ���ȸ �μ������ ������ Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/cert#CheckDeptUser
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	 

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    On Error Resume Next

    Set Presponse = m_HTCashbillService.CheckDeptUser(testCorpNum, userID)

    If Err.Number <> 0 then
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
                <legend>�μ������ ������� Ȯ��</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>