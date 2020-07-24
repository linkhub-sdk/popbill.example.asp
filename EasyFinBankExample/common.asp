<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/EasyFinBankService.asp"-->
<%

	'**************************************************************'
	' 팝빌 계좌조회 API ASP SDK Example
	'
	' - 업데이트 일자 : 2020-07-24
	' - 기술지원 연락처 : 1600-9854 / 070-4304-2991
	' - 기술지원 이메일 : code@linkhub.co.kr
	'
	'**************************************************************
	
	'링크아이디 
	LinkID = "TESTER"

	'비밀키
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_EasyFinBankService = new EasyFinBankService
	m_EasyFinBankService.Initialize LinkID, SecretKey

	'연동환경 설정값, 개발용(True), 상업용(False)
	m_EasyFinBankService.IsTest = True

	' 인증토큰 IP제한기능 사용여부, 권장(True)
	m_EasyFinBankService.IPRestrictOnOff = True

	' 팝빌 API 서비스 고정 IP 사용여부(GA), Ture-사용, False-미사용, 기본값(False)
	m_EasyFinBankService.UseStaticIP = False
%>