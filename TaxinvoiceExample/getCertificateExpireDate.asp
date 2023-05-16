<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺� ���������� ��ϵ� �������� �������� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/cert#GetCertificateExpireDate
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    On Error Resume Next

    certificateExpiration = m_TaxinvoiceService.GetCertificateExpireDate(testCorpNum)

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
                <legend>���������� �����Ͻ� Ȯ��</legend>
                    <ul>
                        <% If code = 0 Then %>
                            <li>�����Ͻ� : <%=certificateExpiration%> </li>
                        <% Else %>
                            <li>Response.code : <%=code%> </li>
                            <li>Response.message : <%=message%> </li>
                        <% End If%>
                    </ul>
            </fieldset>
         </div>
    </body>
</html>