<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 문자 전송시 과금되는 포인트 단가를 확인합니다.
    ' - https://developers.popbill.com/reference/sms/asp/api/point#GetUnitCost
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    userID = "testkorea"

    ' 전송유형 (SMS - 단문, LMS - 장문, MMS - 포토)
    sendType = "SMS"

    On Error Resume Next

    unitCost = m_MessageService.GetUnitCost(testCorpNum, sendType)

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
                <legend>문자메시지 전송단가 확인</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li><%=sendType%> 전송단가 : <%=CInt(unitCost)%> </li>
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