<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺� īī���� API ���� ���������� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/point#GetChargeInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �������� (ATS - �˸���, FTS - ģ���� �ؽ�Ʈ,  FMS-ģ���� �̹���)
    sendType = "ATS"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_KakaoService.GetChargeInfo ( testCorpNum, sendType, UserID )

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
                <legend><%=sendType%> �������� ��ȸ</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> unitCost (���۴ܰ�) : <%=result.unitCost%></li>
                        <li> chargeMethod (��������) : <%=result.chargeMethod%></li>
                        <li> rateSystem (��������) : <%=result.rateSystem%></li>
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
