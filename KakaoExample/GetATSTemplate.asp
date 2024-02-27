<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' 승인된 알림톡 템플릿 정보를 확인합니다.
        ' - https://developers.popbill.com/reference/kakaotalk/asp/api/template#GetATSTemplate
        '**************************************************************

        ' 팝빌회원 사업자번호, "-" 제외
        CorpNum = "1234567890"

        ' 템플릿 코드
        templateCode = "021120000347"

        ' 팝빌회원 아이디
        UserID = "testkorea"

        On Error Resume Next

        Set resultObj = m_KakaoService.GetATSTemplate(CorpNum, templateCode, UserID)

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
                <legend>알림톡 템플릿 정보 확인 </legend>
                    <%
                        If code = 0 Then
                    %>
                        <fieldset class="fieldset2">
                            <legend>  템플릿 정보 </legend>
                            <ul>
                                <li> templateCode (템플릿 코드) : <%=resultObj.templateCode%></li>
                                <li> templateName (템플릿 제목) : <%=resultObj.templateName%></li>
                                <li> template (템플릿 내용) : <%=resultObj.template%></li>
                                <li> plusFriendID (검색용 아이디) : <%=resultObj.plusFriendID%></li>
                                <li> ads (광고메시지 내용) : <%=resultObj.ads%></li>
                                <li> appendix (부가메시지 내용) : <%=resultObj.appendix%></li>
                                <li> secureYN (보안템플릿 여부) : <%=resultObj.secureYN%></li>
                                <li> state (템플릿 상태) : <%=resultObj.state%></li>
                                <li> stateDT (템플릿 상태 일시) : <%=resultObj.stateDT%></li>
                            </ul>
                        <%
                            For i=0 To UBound(resultObj.btns) -1
                        %>
                                <fieldset class="fieldset3">
                                    <legend> 버튼정보 [ <%=i+1%> / <%= UBound(resultObj.btns)%> ] </legend>
                                    <ul>
                                        <li>n : <%=resultObj.btns(i).n%> </li>
                                        <li>t : <%=resultObj.btns(i).t%> </li>
                                        <li>u1 : <%=resultObj.btns(i).u1%> </li>
                                        <li>u2 : <%=resultObj.btns(i).u2%> </li>
                                    </ul>
                            </fieldset>
                        <%
                                Next
                        %>
                        </fieldset>
                        <%
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
            </fieldset>
        </div>
    </body>
</html>
