<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 홈택스연동 정액제 서비스 상태를 확인합니다.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/point#GetFlatRateState
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_HTCashbillService.GetFlatRateState ( testCorpNum, UserID )

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
                <legend> 정액제 서비스 상태 확인</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> referenceID (사업자번호) : <%=result.referenceID%></li>
                        <li> contractDT (정액제 서비스 시작일시) : <%=result.contractDT%></li>
                        <li> useEndDate (정액제 서비스 종료일) : <%=result.useEndDate%></li>
                        <li> baseDate (자동연장 결제일) : <%=result.baseDate%></li>
                        <li> state (정액제 서비스 상태) : <%=result.state%></li>
                        <li> closeRequestYN (정액제 서비스 해지신청 여부) : <%=result.closeRequestYN%></li>
                        <li> useRestrictYN (정액제 서비스 사용제한 여부) : <%=result.useRestrictYN%></li>
                        <li> closeOnExpired (정액제 서비스 만료 시 해지여부) : <%=result.closeOnExpired%></li>
                        <li> unPaidYN (미수금 보유여부) : <%=result.unPaidYN%></li>
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