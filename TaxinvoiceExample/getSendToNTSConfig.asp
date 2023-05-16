<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ���� ����û ���� �ɼ� ���� ���¸� Ȯ���մϴ�.
    ' - �˺� ����û ���� ��å [https://developers.popbill.com/guide/taxinvoice/asp/introduction/policy-of-send-to-nts]
    ' - ����û ���� �ɼ� ������ �˺� ����Ʈ [���ڼ��ݰ�꼭] > [ȯ�漳��] > [���ݰ�꼭 ����] �޴����� ������ �� ������, API�� ������ �Ұ��� �մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#GetSendToNTSConfig
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    On Error Resume Next

    ntsConfig = m_TaxinvoiceService.GetSendToNTSConfig(testCorpNum)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend> ����û ���� ���� Ȯ��</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ntsConfig : <%=ntsConfig%></li>
                        <li>(True)-���� ��� ���� (False)-���� �ڵ� ����</li>
                    </ul>
                <%	Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%	End If	%>
            </fieldset>
         </div>
    </body>
</html>