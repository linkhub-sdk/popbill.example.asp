<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 파트너가 할당한 전송요청 번호를 통해 팩스 전송상태 및 결과를 확인합니다.
    ' - https://docs.popbill.com/fax/asp/api#GetFaxDetailRN
    '**************************************************************
    
    ' 팝빌회원 사업자번호, "-" 제외
    testCorpNum = "1234567890"		

    ' 팝빌회원 아이디
    userID = "testkorea"					

    ' 팩스전송 요청시 할당한 전송요청번호
    requestNum = "" 
 
    On Error Resume Next

    Set result = m_FaxService.GetFaxDetailRN(testCorpNum, requestNum, userID)
    
    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>팩스전송 전송결과 확인 </legend>
                <% If code = 0 Then 
                        For i=0 To result.Count-1
                %>
                    <fieldset class="fieldset2">
                            <legend> 팩스 전송결과 [<%=i+1%>] </legend>
                            <ul>
                                <li>state (전송상태 코드) : <%=result.Item(i).state%> </li>
                                <li>result (전송결과 코드) : <%=result.Item(i).result%> </li>
                                <li>sendNum (발신번호) : <%=result.Item(i).sendNum%> </li>
                                <li>senderName (발신자명) : <%=result.Item(i).senderName%> </li>
                                <li>receiveNum (수신번호) : <%=result.Item(i).receiveNum%> </li>
                                <li>receiveNumType (수신번호 유형) : <%=result.Item(i).receiveNumType%> </li>
                                <li>receiveName (수신자명) : <%=result.Item(i).receiveName%> </li>
                                <li>title (팩스 제목) : <%=result.Item(i).title %> </li>
                                <li>sendPageCnt (페이지수) : <%=result.Item(i).sendPageCnt%></li>
                                <li>successPageCnt (성공 페이지수) : <%=result.Item(i).successPageCnt%></li>
                                <li>failPageCnt (실패 페이지수) : <%=result.Item(i).failPageCnt%></li>
                                <li>cancelPageCnt (취소 페이지수) : <%=result.Item(i).cancelPageCnt%></li>
                                <li>reserveDT (예약시간) : <%=result.Item(i).reserveDT%></li>
                                <li>sendDT (발송시간) : <%=result.Item(i).sendDT%></li>
                                <li>receiptDT (전송 접수시간) : <%=result.Item(i).receiptDT%></li>
                                <li>fileNames (전송파일명 배열) : <%=result.Item(i).fileNames%></li>
                                <li>receiptNum (접수번호) : <%=result.Item(i).receiptNum%> </li>
                                <li>requestNum (요청번호) : <%=result.Item(i).requestNum%> </li>
                                <li>interOPRefKey (파트너 지정키) : <%=result.Item(i).interOPRefKey%> </li>
                                <li>chargePageCnt (과금 페이지수) : <%=result.Item(i).chargePageCnt%> </li>
                                <li>refundPageCnt (환불 페이지수) : <%=result.Item(i).refundPageCnt%></li>
                                <li>tiffFileSize (변환파일용량 (단위 : byte)) : <%=result.Item(i).tiffFileSize%> </li>							
                            </ul>
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
                <%	End If	%>

            </fieldset>
         </div>
    </body>
</html>