<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 세금계산서의 상태에 대한 변경이력을 확인합니다.
    ' - https://docs.popbill.com/taxinvoice/asp/api#GetLogs
    '**************************************************************

    '  팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"	

    ' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
    KeyType= "SELL"             

    ' 문서번호 
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_TaxinvoiceService.GetLogs(testCorpNum, KeyType, MgtKey)
    
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
                <legend> 문서이력확인 </legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count -1 %>
                         <fieldset class="fieldset2">
                            <ul>
                                <li> DocLogType(로그타입) :  <%=result.Item(i).DocLogType%> </li>
                                <li> Log(이력정보) : <%=result.Item(i).Log %> </li>
                                <li> ProcType(처리형태) : <%=result.Item(i).ProcType%> </li>
                                <li> ProcCorpName(처리회사명) : <%=result.Item(i).ProcCorpName%></li>
                                <li> procContactName(처리담당자) : <%=result.Item(i).procContactName%></li>
                                <li> ProcMemo(처리메모) : <%=result.Item(i).ProcMemo %></li>
                                <li> regDT(등록일시) : <%=result.Item(i).regDT %></li>
                                <li> ip(아이피) : <%=result.Item(i).ip %></li>
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