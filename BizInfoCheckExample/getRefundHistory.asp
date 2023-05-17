<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 포인트 환불신청내역을 확인합니다.
    ' - https://developers.popbill.com/reference/bizinfocheck/asp/api/point#GetRefundHistory
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '목록 페이지번호
    Page = 1

    ' 페이지당 표시할 목록개수
    PerPage = 500

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_BizInfoCheckService.GetRefundHistory(testCorpNum, Page, PerPage, UserID)

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
            <legend>연동회원 포인트 환불내역 확인</legend>
            <%
                If code = 0 Then
            %>

            <ul>
                <li> code (응답 코드) : <%=result.code%></li>
                <li> total (총 검색결과 건수) : <%=result.total%></li>
                <li> perPage (페이지당 검색개수) : <%=result.perPage%></li>
                <li> pageNum (페이지 번호) : <%=result.pageNum%></li>
                <li> pageCount (페이지 개수) : <%=result.pageCount%></li>
            </ul>
                <%
                    Dim i
                    For i = 0 To UBound(result.list) - 1
                %>
                <fieldset class="fieldset2">
                    <legend> RefundHistory  [ <%= i+1%> / <%=UBound(result.list)%>]</legend>
                    <ul>
                        <li> reqDT (신청 일시) : <%=result.list(i).reqDT%></li>
                        <li> requestPoint (환불 신청포인트) : <%=result.list(i).requestPoint%></li>
                        <li> accountBank (환불계좌 은행명) : <%=result.list(i).accountBank%></li>
                        <li> accountNum (환불계좌번호) : <%=result.list(i).accountNum%></li>
                        <li> accountName (환불계좌 예금주명) : <%=result.list(i).accountName%></li>
                        <li> state (상태) : <%=result.list(i).state%></li>
                        <li> reason (환불사유) : <%=result.list(i).reason%></li>
                    </ul>
                </fieldset>
                <%
                    Next
                %>
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