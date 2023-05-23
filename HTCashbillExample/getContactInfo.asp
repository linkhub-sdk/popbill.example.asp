<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/member#GetContactInfo
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 확인할 담당자 아이디
    contactID = "testID"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set conInfo = m_HTCashbillService.GetContactInfo(CorpNum, contactID ,UserID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>담당자 정보 확인</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li> id(아이디) : <%=conInfo.id%></li>
                        <li> personName(담당자 성명) : <%=conInfo.personName%></li>
                        <li> email(담당자 이메일) : <%=conInfo.email%></li>
                        <li> tel(담당자 연락처) : <%=conInfo.tel%></li>
                        <li> regDT(등록일시) : <%=conInfo.regDT%></li>
                        <li> SearchRole(담당자 조회권한) : <%=conInfo.SearchRole%></li>
                        <li> mgrYN(관리자 여부) : <%=conInfo.mgrYN%></li>
                        <li> state(상태) : <%=conInfo.state%></li>
                    </ul>
                <%	Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%	End If	%>
            </fieldset>
        </div>
    </body>
</html>