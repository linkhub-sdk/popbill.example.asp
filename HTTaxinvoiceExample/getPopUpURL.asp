<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ������ ���ڼ��ݰ�꼭 1���� �󼼳����� �μ��ϴ� �������� URL�� ��ȯ�մϴ�.
    ' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/search#GetPopUpURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ����û ���ι�ȣ
    NTSConfirmNum = "20161031410002030000277c"

    On Error Resume Next

    url = m_HTTaxinvoiceService.GetPopUpURL(testCorpNum, NTSConfirmNum, userID)

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
                <legend>Ȩ�ý� ���ڼ��ݰ�꼭 ���� �˾� URL</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>URL : <%=url%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>