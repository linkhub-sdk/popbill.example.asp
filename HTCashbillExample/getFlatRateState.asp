<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' Ȩ�ý����� ������ ���� ���¸� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/point#GetFlatRateState
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_HTCashbillService.GetFlatRateState ( testCorpNum, UserID )

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
                <legend> ������ ���� ���� Ȯ��</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> referenceID (����ڹ�ȣ) : <%=result.referenceID%></li>
                        <li> contractDT (������ ���� �����Ͻ�) : <%=result.contractDT%></li>
                        <li> useEndDate (������ ���� ������) : <%=result.useEndDate%></li>
                        <li> baseDate (�ڵ����� ������) : <%=result.baseDate%></li>
                        <li> state (������ ���� ����) : <%=result.state%></li>
                        <li> closeRequestYN (������ ���� ������û ����) : <%=result.closeRequestYN%></li>
                        <li> useRestrictYN (������ ���� ������� ����) : <%=result.useRestrictYN%></li>
                        <li> closeOnExpired (������ ���� ���� �� ��������) : <%=result.closeOnExpired%></li>
                        <li> unPaidYN (�̼��� ��������) : <%=result.unPaidYN%></li>
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
