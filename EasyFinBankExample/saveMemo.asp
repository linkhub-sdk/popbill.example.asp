<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �� ���� �ŷ� ������ �޸� �����մϴ�.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/search#SaveMemo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' �ŷ����� ���̵�, Search API ��ȯ�׸� �� TID
    TID = "01912181100000000120191231000001"

    ' �޸�
    Memo = "asp �޸� �׽�Ʈ"

    On Error Resume Next

    Set Presponse = m_EasyFinBankService.SaveMemo(CorpNum, TID, Memo, UserID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
    End If

    On Error GoTo 0

%>

    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�ŷ����� �޸� ����</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>