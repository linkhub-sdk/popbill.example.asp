<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>팝빌 SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' 수집 상태 확인(GetJobState API) 함수를 통해 상태 정보가 확인된 작업아이디를 활용하여 수집된 전자세금계산서 매입/매출 내역의 요약 정보를 조회합니다.
    ' - 요약 정보 : 전자세금계산서 수집 건수, 공급가액 합계, 세액 합계, 합계 금액
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/search#Summary
    '**************************************************************

    ' 팝빌회원 사업자번호, "-" 제외
    CorpNum = "1234567890"

    ' 팝빌회원 아이디
    UserID = "testkorea"

    ' 수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
    JobID = "019102415000000014"

    ' 문서형태 배열 ("N" 와 "M" 중 선택, 다중 선택 가능)
    ' └ N = 일반 , M = 수정
    ' - 미입력 시 전체조회
    Dim TIType(2)
    TIType(0) = "N"
    TIType(1) = "M"

    ' 과세형태 배열 ("T" , "N" , "Z" 중 선택, 다중 선택 가능)
    ' └ T = 과세, N = 면세, Z = 영세
    ' - 미입력 시 전체조회
    Dim TaxType(3)
    TaxType(0) = "T"
    TaxType(1) = "N"
    TaxType(2) = "Z"

    ' 발행목적 배열 ("R" , "C", "N" 중 선택, 다중 선택 가능)
    ' └ R = 영수, C = 청구, N = 없음
    ' - 미입력 시 전체조회
    Dim PurposeType(3)
    PurposeType(0) = "R"
    PurposeType(1) = "C"
    PurposeType(2) = "N"

    ' 종사업장번호 유무 (null , "0" , "1" 중 택 1)
    ' - null = 전체 , 0 = 없음, 1 = 있음
    TaxRegIDYN = ""

    ' 종사업장번호의 주체 ("S" , "B" , "T" 중 택 1)
    ' └ S = 공급자 , B = 공급받는자 , T = 수탁자
    ' - 미입력시 전체조회
    TaxRegIDType = "S"

    ' 종사업장번호
    ' 다수기재시 콤마(",")로 구분하여 구성 ex ) "0001,0002"
    ' - 미입력시 전체조회
    TaxRegID = ""

    ' 거래처 상호 / 사업자번호 (사업자) / 주민등록번호 (개인) / "9999999999999" (외국인) 중 검색하고자 하는 정보 입력
    ' - 사업자번호 / 주민등록번호는 하이픈('-')을 제외한 숫자만 입력
    ' - 미입력시 전체조회
    SearchString = ""

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.Summary(CorpNum, JobID, TIType, TaxType,  _
                            PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, UserID, SearchString)

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
                <legend>수집 결과 조회</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> count (수집 결과 건수) : <%=result.count%> </li>
                        <li> supplyCostTotal (공급가액 합계) : <%=result.supplyCostTotal%> </li>
                        <li> taxTotal (세액 합계) : <%=result.taxTotal%> </li>
                        <li> amountTotal (합계 금액) : <%=result.amountTotal%> </li>
                    </ul>
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
