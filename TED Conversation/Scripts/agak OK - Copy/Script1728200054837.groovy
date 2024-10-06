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
import com.kms.katalon.core.webui.common.WebUiCommonHelper as WebUiCommonHelper
import org.openqa.selenium.WebElement as WebElement
import com.kms.katalon.core.testobject.ConditionType as ConditionType
import org.openqa.selenium.Keys as Keys

// Buka browser dan navigasi ke halaman yang dituju
WebUI.openBrowser('')

WebUI.navigateToUrl('https://www.telkomsel.com/enterprise')

WebUI.maximizeWindow()

// Klik bublechat dan lakukan input data (nama, email, dll.)
WebUI.click(findTestObject('Object Repository/Page_Homepage  Enterprise/bublechat'))

WebUI.setText(findTestObject('Object Repository/Page_Homepage  Enterprise/input_Name'), 'testing')

WebUI.setText(findTestObject('Object Repository/Page_Homepage  Enterprise/input_Email'), 'test@gmail.com')

WebUI.click(findTestObject('Object Repository/Page_Homepage  Enterprise/chechkbox-input'))

WebUI.delay(5)

WebUI.click(findTestObject('Object Repository/Page_Homepage  Enterprise/button_Mulai Chat'))

// Kirim pertanyaan ke chatbot
WebUI.setText(findTestObject('Page_Homepage  Enterprise/sendquestions'), 'apa saja product iklan yang dimiliki?')

WebUI.sendKeys(findTestObject('Page_Homepage  Enterprise/sendquestions'), Keys.chord(Keys.ENTER))

WebUI.delay(50)

// Tunggu hingga elemen respons chatbot selesai dirender
int timeout = 70

// Tunggu hingga iframe muncul dan beralih ke iframe
WebUI.waitForElementVisible(findTestObject('Page_Homepage  Enterprise/iframe_HUBUNGI KAMI_ted-iframe'), timeout)

WebUI.switchToFrame(findTestObject('Page_Homepage  Enterprise/iframe_HUBUNGI KAMI_ted-iframe'), timeout)

// XPath untuk mengambil elemen respons chatbot terbaru
TestObject dynamicResponseObject = new TestObject().addProperty('xpath', ConditionType.EQUALS, '(//div[contains(@class, \'px-4 pt-6 pb-2 rounded-2xl text-white relative\')])[last()]')

// Tunggu sampai elemen terlihat
println('Menunggu elemen respons chatbot...')

//WebUI.waitForElementVisible(dynamicResponseObject, timeout)
WebUI.delay(timeout)

println('Elemen respons chatbot terlihat.')

// Cari elemen terakhir yang sesuai dengan XPath dinamis
WebElement lastResponse = WebUiCommonHelper.findWebElement(dynamicResponseObject, timeout)

// Cek apakah elemen respons terakhir ditemukan
if (lastResponse != null) {
    // Ambil teks dari elemen terakhir
    String lastResponseText = lastResponse.getText()

    // Tampilkan teks respons terakhir di console
    println('Response terakhir dari chatbot: ' + lastResponseText)
} else {
    println('Tidak ada response yang ditemukan.')
}

// Kembali ke konten utama
WebUI.switchToDefaultContent()

