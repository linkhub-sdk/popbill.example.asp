<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 세금계산서 관련 메일 항목에 대한 발송설정을 확인합니다.
    ' - https://docs.popbill.com/taxinvoice/asp/api#ListEmailConfig
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    ' 팝빌회원 아이디
    UserID = "testkorea"					
    
    On Error Resume Next

    Set emailObj = m_TaxinvoiceService.listEmailConfig(testCorpNum, UserID)

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
                            <% If emailObj.Item(i).emailType = "TAX_ISSUE" Then %>
                                    <li>[정발행] <%= emailObj.Item(i).emailType %>(공급받는자에게 전자세금계산서 발행 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_ISSUE_INVOICER" Then %>
                                    <li>[정발행] <%= emailObj.Item(i).emailType %>(공급자에게 전자세금계산서 발행 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CHECK" Then %>
                                    <li>[정발행] <%= emailObj.Item(i).emailType %>(공급자에게 전자세금계산서 수신확인 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CANCEL_ISSUE" Then %>
                                    <li>[정발행] <%= emailObj.Item(i).emailType %>(공급받는자에게 전자세금계산서 발행취소 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_REQUEST" Then %>
                                    <li>[역발행] <%= emailObj.Item(i).emailType %>(공급자에게 세금계산서를 발행요청 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CANCEL_REQUEST" Then %>
                                    <li>[역발행] <%= emailObj.Item(i).emailType %>(공급받는자에게 세금계산서 취소 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_REFUSE" Then %>
                                    <li>[역발행] <%= emailObj.Item(i).emailType %>(공급받는자에게 세금계산서 거부 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_ISSUE" Then %>
                                    <li>[위수탁발행] <%= emailObj.Item(i).emailType %>(수탁자에게 전자세금계산서 발행 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_ISSUE_TRUSTEE" Then %>
                                    <li>[위수탁발행] <%= emailObj.Item(i).emailType %>(수탁자에게 전자세금계산서 발행 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_ISSUE_INVOICER" Then %>
                                    <li>[위수탁발행] <%= emailObj.Item(i).emailType %>(공급자에게 전자세금계산서 발행 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_CANCEL_ISSUE" Then %>
                                    <li>[위수탁발행] <%= emailObj.Item(i).emailType %>(공급받는자에게 전자세금계산서 발행취소 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_CANCEL_ISSUE_INVOICER" Then %>
                                    <li>[위수탁발행] <%= emailObj.Item(i).emailType %>(공급자에게 전자세금계산서 발행취소 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CLOSEDOWN" Then %>
                                    <li>[처리결과 ]<%= emailObj.Item(i).emailType %>(거래처의 휴폐업 여부 확인 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_NTSFAIL_INVOICER" Then %>
                                    <li>[처리결과] <%= emailObj.Item(i).emailType %>(전자세금계산서 국세청 전송실패 안내 메일 전송 여부) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "ETC_CERT_EXPIRATION" Then %>
                                    <li>[처리결과] <%= emailObj.Item(i).emailType %>(팝빌에 등록된 인증서의 만료예정을 안내하는 메일) : <%= emailObj.Item(i).sendYN %></li>
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
