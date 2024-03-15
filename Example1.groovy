import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.apache.commons.lang.RandomStringUtils as RandomStringUtils
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords

def excel = ExcelKeywords.getWorkbook(excelPath)

def sheet2 = ExcelKeywords.getExcelSheet(excel, sheetName)

def nMatricula = 'TATT' + RandomStringUtils.randomNumeric(3)

ExcelKeywords.setValueToCellByAddress(sheet2, 'C8', nMatricula)

ExcelKeywords.saveWorkbook(excelPath, excel)

WebUI.callTestCase(findTestCase('Transversal/TVL - AgenteNavegacaoLogin'), [('userAgente') : 'xxxx\r\n', ('password') : 'xxx'], 
    FailureHandling.STOP_ON_FAILURE)

WebUI.callTestCase(findTestCase('GEN/GEN - AcederUltimaEscala'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.verifyElementText(findTestObject('Page_JUL/operacoesSeccao'), 'Operações')

WebUI.click(findTestObject('Page_JUL/xxx'))

WebUI.verifyElementText(findTestObject('Page_JUL/xxxTitulo'), 'xxx')

WebUI.click(findTestObject('Page_JUL/oDCriada'))

WebUI.delay(6)

WebUI.click(findTestObject('Page_JUL/oDEquipamentoTransporte'))

WebUI.click(findTestObject('Page_JUL/oDUploadDoc'))

WebUI.delay(6)

WebUI.sendKeys(findTestObject('Page_JUL/oDUploadAnexo'), excelAnexo)

WebUI.click(findTestObject('Page_JUL/fecharSidebar'))

WebUI.delay(7)

WebUI.click(findTestObject('Page_JUL/oDMenu'))

WebUI.click(findTestObject('Page_JUL/genDGCGravar'))

WebUI.delay(7)

WebUI.closeBrowser()

