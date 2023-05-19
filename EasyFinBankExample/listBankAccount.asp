<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 팝빌에 등록된 계좌정보 목록을 반환합니다.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/manage#ListBankAccount
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_EasyFinBankService.ListBankAccount(testCorpNum, UserID)

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
                <legend>계좌 목록</legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count-1
                %>
                            <fieldset class="fieldset2">
                                <legend>ListBankAccount [ <%=i+1%> / <%=result.Count%> ] </legend>
                                    <ul>
                                        <li>accountNumber (계좌번호) : <%=result.Item(i).accountNumber%></li>
                                        <li>bankCode (기관코드) : <%=result.Item(i).bankCode%></li>
                                        <li>accountName (계좌 별칭) : <%=result.Item(i).accountName%></li>
                                        <li>accountType (계좌유형) : <%=result.Item(i).accountType%></li>
                                        <li>state (정액제 상태) : <%=result.Item(i).state%></li>
                                        <li>regDT (등록일시) : <%=result.Item(i).regDT%></li>
                                        <li>contractDT (정액제 서비스 시작일시) : <%=result.Item(i).contractDT %> </li>
                                        <li>useEndDate (정액제 서비스 종료일) : <%=result.Item(i).useEndDate %> </li>
                                        <li>baseDate (자동연장 결제일) : <%=result.Item(i).baseDate %> </li>
                                        <li>contractState (정액제 서비스 상태) : <%=result.Item(i).contractState%> </li>
                                        <li>closeRequestYN (정액제 서비스 해지신청 여부) : <%=result.Item(i).closeRequestYN%> </li>
                                        <li>useRestrictYN (정액제 서비스 사용제한 여부) : <%=result.Item(i).useRestrictYN%> </li>
                                        <li>closeOnExpired (정액제 서비스 만료 시 해지 여부) : <%=result.Item(i).closeOnExpired %> </li>
                                        <li>unPaidYN (미수금 보유 여부) : <%=result.Item(i).unPaidYN %> </li>
                                        <li>memo (메모) : <%=result.Item(i).memo%></li>

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
