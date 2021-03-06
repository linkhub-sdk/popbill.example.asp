<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
    '**************************************************************
    ' 팝빌 전자세금계산서 API ASP SDK Example
    '
    ' - 업데이트 일자 : 2021-06-01
    ' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
    ' - 연동 기술지원 이메일 : code@linkhub.co.kr
    '
    ' <테스트 연동개발 준비사항>
    ' 1) 23, 26번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
    '    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
    ' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
    ' 3) 전자세금계산서 발행을 위해 공인인증서를 등록합니다.
    '    - 팝빌사이트 로그인 > [전자세금계산서] > [환경설정]
    '      > [공인인증서 관리]
    '    - 공인인증서 등록 팝업 URL (GetTaxCertURL API)을 이용하여 등록
    '**************************************************************

    ' 링크아이디 
    LinkID = "TESTER"

    ' 비밀키
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_TaxinvoiceService = new TaxinvoiceService
    
    ' 세금계산서 API 서비스 모듈 초기화
    m_TaxinvoiceService.Initialize LinkID, SecretKey

    ' 연동환경 설정값, 개발용(True),  상업용(False)
    m_TaxinvoiceService.IsTest = True

    ' 인증토큰 IP제한기능 사용여부, 권장(True)
    m_TaxinvoiceService.IPRestrictOnOff = True
    
    ' 팝빌 API 서비스 고정 IP 사용여부(GA), Ture-사용, False-미사용, 기본값(False)
    m_TaxinvoiceService.UseStaticIP = False
    
    ' 로컬시스템 시간 사용여부 True-사용(기본값-권장), false-미사용
    m_TaxinvoiceService.UseLocalTimeYN = True
%>