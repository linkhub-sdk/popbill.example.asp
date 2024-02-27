<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 수집 상태 확인(GetJobState API) 함수를 통해 상태 정보가 확인된 작업아이디를 활용하여 수집된 현금영수증 매입/매출 내역의 요약 정보를 조회합니다.
    ' - 요약 정보 : 현금영수증 수집 건수, 공급가액 합계, 세액 합계, 봉사료 합계, 합계 금액
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/search#Summary
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    '수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
    JobID = "016111417000000002"

    ' 문서형태 배열 ("N" 와 "C" 중 선택, 다중 선택 가능)
    ' └ N = 일반 현금영수증 , C = 취소현금영수증
    ' - 미입력 시 전체조회
    Dim TradeType(2)
    TradeType(0) = "N"
    TradeType(1) = "C"

    ' 거래구분 배열 ("P" 와 "C" 중 선택, 다중 선택 가능)
    ' └ P = 소득공제용 , C = 지출증빙용
    ' - 미입력 시 전체조회
    Dim TradeUsage(2)
    TradeUsage(0) = "P"
    TradeUsage(1) = "C"

    On Error Resume Next

    Set result = m_HTCashbillService.Summary(CorpNum, JobID, TradeType, TradeUsage, UserID)

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
                <legend>수집 결과 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> count (수집 결과 건수) : <%=result.count%> </li>
                        <li> supplyCostTotal (공급가액 합계) : <%=result.supplyCostTotal%> </li>
                        <li> taxTotal (세액 합계) : <%=result.taxTotal%> </li>
                        <li> serviceFeeTotal (봉사료 합계) : <%=result.serviceFeeTotal%> </li>
                        <li> amountTotal (합계 금액) : <%=result.amountTotal%> </li>
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
