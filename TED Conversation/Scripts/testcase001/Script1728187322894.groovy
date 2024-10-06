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

WebUI.openBrowser('')

WebUI.navigateToUrl('https://www.telkomsel.com/enterprise')

WebUI.click(findTestObject('Object Repository/Page_Homepage  Enterprise/bublechat'))

WebUI.setText(findTestObject('Object Repository/Page_Homepage  Enterprise/input_Name'), 'testing')

WebUI.setText(findTestObject('Object Repository/Page_Homepage  Enterprise/input_Email'), 'test@gmail.com')

WebUI.click(findTestObject('Object Repository/Page_Homepage  Enterprise/chechkbox-input'))

WebUI.click(findTestObject('Object Repository/Page_Homepage  Enterprise/button_Mulai Chat'))

WebUI.setText(findTestObject('Page_Homepage  Enterprise/sendquestions'), 'saya mau tau iot')

WebUI.sendKeys(findTestObject('Page_Homepage  Enterprise/sendquestions'), Keys.chord(Keys.ENTER))

