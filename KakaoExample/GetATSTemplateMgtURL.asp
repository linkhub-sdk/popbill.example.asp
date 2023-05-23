<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' 알림톡 템플릿을 신청하고 승인심사 결과를 확인하며 등록 내역을 확인하는 알림톡 템플릿 관리 페이지 팝업 URL을 반환합니다.
        ' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
        ' - 승인된 알림톡 템플릿은 수정이 불가하고, 변경이 필요한 경우 새롭게 템플릿 신청을 해야합니다.
        ' - https://developers.popbill.com/reference/kakaotalk/asp/api/template#GetATSTemplateMgtURL
        '**************************************************************

        ' 팝빌회원 사업자번호, "-" 제외
        CorpNum = "1234567890"

        ' 팝빌회원 아이디
        UserID = "testkorea"

        On Error Resume Next

        url = m_KakaoService.GetATSTemplateMgtURL(CorpNum, UserID)

        If Err.Number <> 0 then
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
                <legend>알림톡 템플릿 관리 팝업 URL</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>URL : <%=url%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>