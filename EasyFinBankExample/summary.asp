<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 수집 상태 확인(GetJobState API) 함수를 상태 정보가 확인된 작업아이디를 활용하여 계좌 거래내역의 요약 정보를 조회합니다.
    ' - 요약 정보는 입·출 금액 합계, 입·출 거래 건수를 가리킵니다.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/search#Summary
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    '수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
    JobID = "019123114000000010"

    ' 거래유형 배열 ("I" 와 "O" 중 선택, 다중 선택 가능)
    ' └ I = 입금 , O = 출금
    ' - 미입력 시 전체조회
    Dim TradeType(2)
    TradeType(0) = "I"
    TradeType(1) = "O"

    ' "입·출금액" / "메모" / "비고" 중 검색하고자 하는 값 입력
    ' - 메모 = 거래내역 메모저장(SaveMemo)을 사용하여 저장한 값
    ' - 비고 = EasyFinBankSearchDetail의 remark1, remark2, remark3 값
    ' - 미입력시 전체조회
    SearchString = ""

    On Error Resume Next

    Set result = m_EasyFinBankService.Summary(testCorpNum, JobID, TradeType, SearchString, UserID)

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
                <legend>수집 결과 요약정보 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> count (수집 결과 건수) : <%=result.count%> </li>
                        <li> cntAccIn (입금거래 건수) : <%=result.cntAccIn%> </li>
                        <li> cntAccOut (출금거래 건수) : <%=result.cntAccOut%> </li>
                        <li> totalAccIn (입금액 합계) : <%=result.totalAccIn%> </li>
                        <li> totalAccOut (출금액 합계) : <%=result.totalAccOut%> </li>
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