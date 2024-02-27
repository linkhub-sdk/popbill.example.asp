<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>휴폐업조회 API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' 다수건의 사업자번호에 대한 휴폐업정보를 확인합니다. (최대 1,000건)
        ' - https://developers.popbill.com/reference/closedown/asp/api/check#CheckCorpNums
        '**************************************************************

        ' 팝빌회원 사업자번호
        UserCorpNum = "1234567890"

        ' 조회할 사업자번호 배열, 최대 1000건
        Dim CorpNumList(3)
        CorpNumList(0) = "1234567890"
        CorpNumList(1) = "6798700433"
        CorpNumList(2) = "110-04-45791"

        On Error Resume Next

        Set result = m_ClosedownService.checkCorpNums(UserCorpnum, CorpNumList)

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
                <legend>휴폐업조회 - 대량</legend>
                <br/>
                <p class="info">> state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업</p>
                <p class="info">> taxType (사업 유형) : null-알수없음, 10-일반과세자, 20-면세과세자, 30-간이과세자, 31-간이과세자(세금계산서 발급사업자), 40-비영리법인, 국가기관</p>

                <br/>
            <%
                If Not IsEmpty(result) Then
                    For i=0 To result.Count-1
            %>
                    <fieldset class="fieldset2">
                        <legend>휴폐업정보 [<%=i+1 %>]</legend>
                        <ul>
                                <li>사업자번호 (corpNum) : <%= result.Item(i).corpNum%></li>
                                <li>휴폐업상태 (state) : <%= result.Item(i).state%></li>
                                <li>사업자유형 (taxType) : <%= result.Item(i).taxType%></li>
                                <li>휴폐업일자 (stateDate) : <%= result.Item(i).stateDate%></li>
                                <li>과세유형 전환일자 (typeDate) : <%= result.Item(i).typeDate%></li>
                                <li>국세청 확인일자 (checkDate) : <%= result.Item(i).checkDate%></li>
                        </ul>
                    </fieldset>
            <%
                    Next
                End If
                If Not IsEmpty(code) then
            %>
                <fieldset class="fieldset2">
                    <legend>휴폐업조회 - 단건</legend>
                    <ul>
                        <li>Response.code : <%= code %> </li>
                        <li>Response.message : <%= message %></li>
                    </ul>
                </fieldset>
            <%
                End If
            %>

            </fieldset>

        <script type ="text/javascript">
             window.onload=function(){
                 document.getElementById('CorpNum').focus();
             }

             function search(){
                document.getElementById('corpnum_form').submit();
             }
         </script>
    </body>
</html>
