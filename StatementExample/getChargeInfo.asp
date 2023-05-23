<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌 전자명세서 API 서비스 과금정보를 확인합니다.
    ' - https://developers.popbill.com/reference/statement/asp/api/point#GetChargeInfo
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
    itemCode = "121"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_StatementService.GetChargeInfo(CorpNum, ItemCode, UserID)

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
                <legend>과금정보 확인</legend>
                <%
                    If code = 0 Then
                %>
                        <ul>
                            <li> unitCost (발행단가) : <%=result.unitCost%></li>
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