<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �α��� ���·� �˺� ����Ʈ�� ���ݿ����� ������ �޴��� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
    ' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/info#GetURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' TBOX(�ӽù�����), PBOX(���๮����), WRITE(���ݿ����� �ۼ�)
    TOGO = "PBOX"

    On Error Resume Next

    url = m_CashbillService.GetURL(CorpNum, UserID, TOGO)

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
                <legend>�˺� ���ݿ����� ������ URL</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>URL : <%=url%> </li>
                    <% Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>
