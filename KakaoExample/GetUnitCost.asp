<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 카카오톡 전송단가를 확인합니다.
    ' - https://docs.popbill.com/kakao/asp/api#GetUnitCost
    '**************************************************************

    '팝빌 회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '팝빌 회원 아이디
    userID = "testkorea"

    '전송유형 (ATS - 알림톡, FTS - 친구톡 텍스트, FMS - 친구톡 이미지)
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
                <legend>카카오톡 전송단가 확인</legend>
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