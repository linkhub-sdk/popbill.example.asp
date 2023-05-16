<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 포인트 사용내역을 확인합니다.
    ' - https://developers.popbill.com/reference/accountcheck/asp/api/point#GetUseHistory
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '조회 기간의 시작일자
    SDate = "20230501"

    '조회 기간의 종료일자
    EDate = "20230530"

    '목록 페이지번호
    Page = 1

    '페이지당 표시할 목록 개수
    PerPage = 500

    '거래일자를 기준으로 하는 목록 정렬 방향
    Order = "D"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set useHistoryResult = m_AccountCheckService.GetUseHistory(testCorpNum, SDate,EDate,Page,PerPage,Order, UserID)

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
                <legend>환불 가능 포인트 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> UseHistoryResult </legend>
                            <ul>
                                <li> code (응답코드) : <%=useHistoryResult.code%></li>
                                <li> total (총 검색결과 건수) : <%=useHistoryResult.total%></li>
                                <li> perPage (페이지당 검색개수) : <%=useHistoryResult.perPage%></li>
                                <li> pageNum (페이지 번호) : <%=useHistoryResult.pageNum%></li>
                                <li> pageCount (페이지 개수) : <%=useHistoryResult.pageCount%></li>
                                <%
                                    Dim i
                                    For i = 0 to UBound(useHistoryResult.list)-1

                                %>
                                <li> itemCode (서비스코드) : <%=useHistoryResult.list(i).itemCode%></li>
                                <li> txType (포인트 증감 유형) : <%=useHistoryResult.list(i).txType%></li>
                                <li> txPoint (증감 포인트) : <%=useHistoryResult.list(i).txPoint%></li>
                                <li> balance (잔여 포인트) : <%=useHistoryResult.list(i).balance%></li>
                                <li> txDT (포인트 증감 일시) : <%=useHistoryResult.list(i).txDT%></li>
                                <li> userID (담당자 아이디) : <%=useHistoryResult.list(i).userID%></li>
                                <li> userName (담당자명) : <%=useHistoryResult.list(i).userName%></li>
                                <%
                                Next
                                %>
                            </ul>
                        </fieldset>
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