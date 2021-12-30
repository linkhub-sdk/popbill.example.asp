<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 수집 요청 상태를 확인합니다.
    ' - https://docs.popbill.com/easyfinbank/asp/api#GetJobState
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    ' 수집요청시 반환받은 작업아이디(jobID)
    JobID = "019123114000000010"	

    ' 팝빌회원 아이디
    UserID = "testkorea"	
    
    On Error Resume Next

    Set result = m_EasyFinBankService.GetJobState(testCorpNum, JobID, UserID)

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
                            <li> startDate (시작일자) : <%=result.startDate%></li>
                            <li> endDate (종료일자) : <%=result.endDate%></li>
                            <li> errorCode (오류코드) : <%=result.errorCode%></li>
                            <li> errorReason (오류메시지) : <%=result.errorReason%></li>
                            <li> jobStartDT (작업 시작일시) : <%=result.jobStartDT%></li>
                            <li> jobEndDT (작업 종료일시) : <%=result.jobEndDT%></li>
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
