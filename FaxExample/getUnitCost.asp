<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팩스 전송시 과금되는 포인트 단가를 확인합니다.
    ' - https://developers.popbill.com/reference/fax/asp/api/point#GetUnitCost
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 수신번호 유형 : "일반" / "지능" 중 택 1
    ' └ 일반망 : 지능망을 제외한 번호
    ' └ 지능망 : 030*, 050*, 070*, 080*, 대표번호
    receiveNumType = "지능"

    On Error Resume Next

    unitCost = m_FaxService.GetUnitCost(testCorpNum, receiveNumType)

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
                <legend>팩스 전송 단가 확인 </legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>전송 단가 : <%=unitCost%> </li>
                    <% Else %>
                        <li> Response.code : <%=code%></li>
                        <li> Response.message : <%=message%></li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>