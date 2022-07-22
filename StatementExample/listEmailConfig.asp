<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 전자명세서 관련 메일 항목에 대한 발송설정을 확인합니다.
    ' - https://docs.popbill.com/statement/asp/api#ListEmailConfig
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    ' 팝빌회원 아이디
    UserID = "testkorea"					
    
    On Error Resume Next

    Set emailObj = m_StatementService.listEmailConfig(testCorpNum, UserID)

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
                <legend>알림메일 전송목록 조회</legend>
                        <ul>
                        <%
                            If code = 0 Then
                            For i=0 To emailObj.Count-1
                        %>
                            <% If emailObj.Item(i).emailType = "SMT_ISSUE" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (공급받는자에게 전자명세서가 발행 되었음을 알려주는 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "SMT_ACCEPT" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (공급자에게 전자명세서가 승인 되었음을 알려주는 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "SMT_DENY" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (공급자에게 전자명세서가 거부 되었음을 알려주는 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "SMT_CANCEL" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (공급받는자에게 전자명세서가 취소 되었음을 알려주는 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "SMT_CANCEL_ISSUE" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (공급받는자에게 전자명세서가 발행취소 되었음을 알려주는 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                        <%
                            Next
                            Else
                        %>
                        </ul>
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
