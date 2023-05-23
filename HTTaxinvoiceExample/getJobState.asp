<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 수집 요청(RequestJob API) 함수를 통해 반환 받은 작업 아이디의 상태를 확인합니다.
    ' - 수집 결과 조회(Search API) 함수 또는 수집 결과 요약 정보 조회(Summary API) 함수를 사용하기 전에
    '   수집 작업의 진행 상태, 수집 작업의 성공 여부를 확인해야 합니다.
    ' - 작업 상태(jobState) = 3(완료)이고 수집 결과 코드(errorCode) = 1(수집성공)이면
    '   수집 결과 내역 조회(Search) 또는 수집 결과 요약 정보 조회(Summary) 를 해야합니다.
    ' - 작업 상태(jobState)가 3(완료)이지만 수집 결과 코드(errorCode)가 1(수집성공)이 아닌 경우에는
    '   오류메시지(errorReason)로 수집 실패에 대한 원인을 파악할 수 있습니다.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/job#GetJobState
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 수집요청시 반환받은 작업아이디(jobID)
    JobID = "016111416000000024"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.GetJobState(CorpNum, JobID, UserID)

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
                <legend>수집 상태 확인</legend>
                <%
                    If code = 0 Then
                %>
                        <ul>
                            <li> jobID (작업아이디) : <%=result.jobID%></li>
                            <li> jobState (수집상태) : <%=result.jobState%></li>
                            <li> queryType (수집유형) : <%=result.queryType%></li>
                            <li> queryDateType (일자유형) : <%=result.queryDateType%></li>
                            <li> queryStDate (시작일자) : <%=result.queryStDate%></li>
                            <li> queryEnDate (종료일자) : <%=result.queryEnDate%></li>
                            <li> errorCode (오류코드) : <%=result.errorCode%></li>
                            <li> errorReason (오류메시지) : <%=result.errorReason%></li>
                            <li> jobStartDT (작업 시작일시) : <%=result.jobStartDT%></li>
                            <li> jobEndDT (작업 종료일시) : <%=result.jobEndDT%></li>
                            <li> collectCount (수집개수) : <%=result.collectCount%></li>
                            <li> regDT (수집 요청일시) : <%=result.regDT%></li>
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