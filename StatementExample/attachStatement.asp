<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ϳ��� ���ڸ������� �ٸ� ���ڸ������� ÷���մϴ�.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#AttachStatement
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' ������ �ڵ� - 121(�ŷ�������), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
    itemCode = 121

    ' ����������ȣ
    mgtKey = "20220720-ASP-001"

    ' ÷���� ������ �ڵ�
    subItemCode = 121

    ' ÷���� ������ ������ȣ
    subMgtKey ="20220720-100"

    On Error Resume Next

    Set result = m_StatementService.AttachStatement(testCorpNum, itemCode, mgtKey, subItemCode, subMgtKey)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�ٸ� ���ڸ����� ÷��</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>