<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ�� ����Ʈ ������ �Աݽ�û���� 1���� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/point#GetSettleResult
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    '�����ڵ�
    SettleCode = "202305120000000035"

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_KakaoService.GetSettleResult(CorpNum, SettleCode, UserID)

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
                <legend>����ȸ�� ������ �Աݽ�û ����Ȯ��</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> PaymentHistory </legend>
                        <ul>
                            <li>productType (���� ����) : <%= result.productType %></li>
                            <li>productName (���� ��ǰ��) : <%= result.productName %></li>
                            <li>settleType (��������) : <%= result.settleType %></li>
                            <li>settlerName (����ڸ�) : <%= result.settlerName %></li>
                            <li>settlerEmail (����ڸ���) : <%= result.settlerEmail %></li>
                            <li>settleCost (�����ݾ�) : <%= result.settleCost %></li>
                            <li>settlePoint (��������Ʈ) : <%= result.settlePoint %></li>
                            <li>settleState (��������) : <%= result.settleState %></li>
                            <li>regDT (����Ͻ� ) : <%= result.regDT %></li>
                            <li>stateDT (�����Ͻ� ) : <%= result.stateDT %></li>
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
