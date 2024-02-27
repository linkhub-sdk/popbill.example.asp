<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 연동회원의 회사정보를 확인합니다.
    ' - https://developers.popbill.com/reference/bizinfocheck/asp/api/member#GetCorpInfo
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    '팝빌회원 아이디
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_BizInfoCheckService.GetCorpInfo(CorpNum, UserID)

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
                <legend>회사정보 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> CorpInfo </legend>
                            <ul>
                                <li> ceoname (대표자명) : <%=result.ceoname%></li>
                                <li> corpName (상호) : <%=result.corpName%></li>
                                <li> addr (주소) : <%=result.addr%></li>
                                <li> bizType (업태) : <%=result.bizType%></li>
                                <li> bizClass (종목) : <%=result.bizClass%></li>
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
