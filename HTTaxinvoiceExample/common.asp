<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%

    '**************************************************************'
    ' 팝빌 홈택스 전자세금계산서 연동 API ASP SDK Example
    '
    ' ASP SDK 연동환경 설정방법 안내 : https://docs.popbill.com/htcashbill/tutorial/asp
    ' - 업데이트 일자 : 2022-07-20
    ' - 연동 기술지원 연락처 : 1600-9854
    ' - 연동 기술지원 이메일 : code@linkhubcorp.com
    '
    ' <테스트 연동개발 준비사항>
    ' 1) 23, 26번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
    '    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
    ' 3) 홈택스 인증 처리를 합니다. (부서사용자등록 / 공인인증서 등록)
    '     - 팝빌로그인 > [홈택스연동] > [환경설정] > [인증 관리] 메뉴
    '     - 홈택스연동 인증 관리 팝업 URL(GetCertificatePopUpURL API) 반환된 URL을 이용하여
    '       홈택스 인증 처리를 합니다.
    '**************************************************************
    
    '링크아이디 
    LinkID = "TESTER"

    '비밀키
    SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_HTTaxinvoiceService = new HTTaxinvoiceService

    m_HTTaxinvoiceService.Initialize LinkID, SecretKey

    ' 연동환경 설정값, True-개발용, false-상업용
    m_HTTaxinvoiceService.IsTest = True

    ' 인증토큰 발급 IP 제한 On/Off, True-사용, false-미사용, 기본값(True)
    m_HTTaxinvoiceService.IPRestrictOnOff = True
    
    ' 팝빌 API 서비스 고정 IP 사용여부, True-사용, false-미사용, 기본값(false)
    m_HTTaxinvoiceService.UseStaticIP = False
    
    ' 로컬시스템 시간 사용여부 Ture-사용, False-미사용, 기본값(True)
    m_HTTaxinvoiceService.UseLocalTimeYN = True
%>