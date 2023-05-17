<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌 팩스 API 서비스 과금정보를 확인합니다.
    ' - https://developers.popbill.com/reference/fax/asp/api/point#GetChargeInfo
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 수신번호 유형 : "일반" / "지능" 중 택 1
    ' └ 일반망 : 지능망을 제외한 번호
    ' └ 지능망 : 030*, 050*, 070*, 080*, 대표번호
    receiveNumType = "지능"

    On Error Resume Next

    Set result = m_FaxService.GetChargeInfo ( testCorpNum, UserID, receiveNumType )

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
                <legend> 과금정보 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> unitCost (전송단가) : <%=result.unitCost%></li>
                        <li> chargeMethod (과금유형) : <%=result.chargeMethod%></li>
                        <li> rateSystem (과금제도) : <%=result.rateSystem%></li>
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