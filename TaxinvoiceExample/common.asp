<!--#include virtual="/Popbill/Popbill.asp"-->
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
    '**************************************************************
    ' 팝빌 전자세금계산서 API ASP SDK Example
    '
    ' ASP SDK 연동환경 설정방법 안내 : https://developers.popbill.com/guide/taxinvoice/asp/getting-started/environment-set-up
    ' - 업데이트 일자 : 2022-07-20
    ' - 연동 기술지원 연락처 : 1600-9854
    ' - 연동 기술지원 이메일 : code@linkhubcorp.com
    '
    ' <테스트 연동개발 준비사항>
    ' 1) 22, 25번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
    '    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
    ' 2) 전자세금계산서 발행을 위해 공인인증서를 등록합니다.
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

    ' 연동환경 설정값, True-개발용, false-상업용
    m_TaxinvoiceService.IsTest = True

    ' 인증토큰 발급 IP 제한 On/Off, True-사용, false-미사용, 기본값(True)
    m_TaxinvoiceService.IPRestrictOnOff = True

    ' 팝빌 API 서비스 고정 IP 사용여부, True-사용, false-미사용, 기본값(false)
    m_TaxinvoiceService.UseStaticIP = False

    ' 로컬시스템 시간 사용여부 Ture-사용, False-미사용, 기본값(True)
    m_TaxinvoiceService.UseLocalTimeYN = True
%>