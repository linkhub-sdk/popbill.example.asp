
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 승인된 알림톡 템플릿 목록을 확인합니다.
    ' - https://docs.popbill.com/kakao/asp/api#ListATSTemplate
    '**************************************************************

    '팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    On Error Resume Next

    Set resultObj = m_KakaoService.ListATSTemplate(testCorpNum)
    
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
                <legend>알림톡 템플릿 목록 조회 </legend>
                    <% 
                        If code = 0 Then
                            For i=0 To resultObj.Count-1 
                    %>
                        <fieldset class="fieldset2">
                            <legend>  템플릿 정보 [ <%=i+1%> / <%= resultObj.Count %> ] </legend>
                            <ul>
                                <li> templateCode : <%=resultObj(i).templateCode%></li>
                                <li> templateName : <%=resultObj(i).templateName%></li>
                                <li> template : <%=resultObj(i).template%></li>
                                <li> plusFriendID : <%=resultObj(i).plusFriendID%></li>
                                <li> ads : <%=resultObj(i).ads%></li>
                                <li> appendix : <%=resultObj(i).appendix%></li>
                            </ul>
                        <%
                            For j=0 To UBound(resultObj(i).btns) -1
                        %>
                                <fieldset class="fieldset3">
                                    <legend> 버튼정보 [ <%=j+1%> / <%= UBound(resultObj(i).btns)%> ] </legend>
                                    <ul>
                                        <li>n : <%=resultObj(i).btns(j).n%> </li>
                                        <li>t : <%=resultObj(i).btns(j).t%> </li>
                                        <li>u1 : <%=resultObj(i).btns(j).u1%> </li>
                                        <li>u2 : <%=resultObj(i).btns(j).u2%> </li>
                                    </ul>
                            </fieldset>
                        <% 
                                Next
                        %>
                        </fieldset>
                        <%
                            Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>

            </fieldset>
         </div>
    </body>
</html>