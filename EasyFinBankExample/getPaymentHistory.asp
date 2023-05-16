<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 포인트 결제내역을 확인합니다.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/point#GetPaymentHistory
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 조회 기간의 시작일자 (형식 : yyyyMMdd)
    SDate = "20230501"

    ' 조회 기간의 종료일자 (형식 : yyyyMMdd)
    EDate = "20230530"

    ' 목록 페이지번호
    Page = 1

    ' 페이지당 표시할 목록 개수
    PerPage = 500

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_EasyFinBankService.GetPaymentHistory(testCorpNum, SDate, EDate, Page, PerPage, UserID)

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
                <legend>연동회원 포인트 결제내역 확인</legend>
                <%
                    If code = 0 Then
                %>
                <ul>
                    <li> code (환불 가능 포인트) : <%=result.code%></li>
                    <li> total (환불 가능 포인트) : <%=result.total%></li>
                    <li> perPage (환불 가능 포인트) : <%=result.perPage%></li>
                    <li> pageNum (환불 가능 포인트) : <%=result.pageNum%></li>
                    <li> pageCount (환불 가능 포인트) : <%=result.pageCount%></li>
                </ul>
                <%
                    Dim i
                    For i = 0 To UBound(result.list) -1
                %>
                    <fieldset class="fieldset2">
                        <legend> PaymentHistory [ <%= i+1%> / <%=UBound(result.list)%>]</legend>
                        <ul>
                            <li>productType : <%= result.list(i).productType %></li>
                            <li>productName : <%= result.list(i).productName %></li>
                            <li>settleType : <%= result.list(i).settleType %></li>
                            <li>settlerName : <%= result.list(i).settlerName %></li>
                            <li>settlerEmail : <%= result.list(i).settlerEmail %></li>
                            <li>settleCost : <%= result.list(i).settleCost %></li>
                            <li>settlePoint : <%= result.list(i).settlePoint %></li>
                            <li>settleState : <%= result.list(i).settleState %></li>
                            <li>regDT : <%= result.list(i).regDT %></li>
                            <li>stateDTz : <%= result.list(i).stateDTz %></li>
                        </ul>
                    </fieldset>
                <%
                    Next
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