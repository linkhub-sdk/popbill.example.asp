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
        ' - https://docs.popbill.com/kakao/asp/api#GetATSTemplate
        '**************************************************************

        ' 팝빌회원 사업자번호, "-" 제외
        testCorpNum = "1234567890"		

        ' 템플릿 코드
        templateCode = "021120000347"

        ' 팝빌회원 아이디
        UserID = "testkorea"

        On Error Resume Next

        Set resultObj = m_KakaoService.GetATSTemplate(testCorpNum, templateCode, UserID)

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
                                <li> templateCode : <%=resultObj.templateCode%></li>
                                <li> templateName : <%=resultObj.templateName%></li>
                                <li> template : <%=resultObj.template%></li>
                                <li> plusFriendID : <%=resultObj.plusFriendID%></li>
                                <li> ads : <%=resultObj.ads%></li>
                                <li> appendix : <%=resultObj.appendix%></li>
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