<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' īī���� ���۴ܰ��� Ȯ���մϴ�.
    ' - https://docs.popbill.com/kakao/asp/api#GetUnitCost
    '**************************************************************

    '�˺� ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺� ȸ�� ���̵�
    userID = "testkorea"

    '�������� (ATS - �˸���, FTS - ģ���� �ؽ�Ʈ, FMS - ģ���� �̹���)
    sendType = "ATS"

    On Error Resume Next

    unitCost = m_KakaoService.GetUnitCost(testCorpNum, sendType)

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
                <legend>īī���� ���۴ܰ� Ȯ��</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li><%=sendType%> ���۴ܰ� : <%=CInt(unitCost)%> </li>
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