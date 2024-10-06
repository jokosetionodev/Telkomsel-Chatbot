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
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.nio.file.Files as Files
import java.nio.file.Paths as Paths
import java.util.UUID as UUID

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

// Membaca data pertanyaan dari Test Data
TestData questionsData = findTestData('Data Files/questions/questions' // Ganti dengan nama Test Data yang sesuai
    )

int questionCount = questionsData.getRowNumbers()

// Membuat workbook baru untuk menyimpan respons
String uuid = UUID.randomUUID().toString()

String outputExcelFilePath = ('C:\\Users\\User\\Documents\\hasilexcel\\response_' + uuid) + '.xlsx'

Workbook outputWorkbook = new XSSFWorkbook()

Sheet outputSheet = outputWorkbook.createSheet('Responses')

// Membuat header untuk output
Row headerRow = outputSheet.createRow(0)

headerRow.createCell(0).setCellValue('Pertanyaan')

headerRow.createCell(1).setCellValue('Response')

// Mengirim pertanyaan satu per satu
for (int i = 1; i <= questionCount; i++) {
    // Ambil pertanyaan dari Test Data
    String question = questionsData.getValue(1, i // Asumsi pertanyaan ada di kolom pertama
        )

    // Kirim pertanyaan ke chatbot
    WebUI.setText(findTestObject('Page_Homepage  Enterprise/sendquestions'), question)

    WebUI.sendKeys(findTestObject('Page_Homepage  Enterprise/sendquestions'), Keys.chord(Keys.ENTER))

    // Tunggu selama 13 detik sebelum mengambil screenshot
    WebUI.delay(13)

    // Ambil screenshot
    String screenshotPath = ((('C:\\Users\\User\\Documents\\hasilexcel\\screenshot_' + uuid) + '_') + i) + '.png' // Ganti dengan path yang diinginkan

    WebUI.takeScreenshot(screenshotPath)

    println('Screenshot diambil dan disimpan di: ' + screenshotPath)

 

    // Tunggu hingga elemen respons chatbot selesai dirender
    int timeout = 2

    WebUI.waitForElementVisible(findTestObject('Page_Homepage  Enterprise/iframe_HUBUNGI KAMI_ted-iframe'), timeout)

    WebUI.switchToFrame(findTestObject('Page_Homepage  Enterprise/iframe_HUBUNGI KAMI_ted-iframe'), timeout)

    // XPath untuk mengambil elemen respons chatbot terbaru
    TestObject dynamicResponseObject = new TestObject().addProperty('xpath', ConditionType.EQUALS, '(//div[contains(@class, \'px-4 pt-6 pb-2 rounded-2xl text-white relative\')])[last()]')

    println('Menunggu elemen respons chatbot...')

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

        // Menyimpan pertanyaan dan respons terakhir ke sheet output
        Row outputRow = outputSheet.createRow(i)

        outputRow.createCell(0).setCellValue(question)

        outputRow.createCell(1).setCellValue(lastResponseText)
    } else {
        println('Tidak ada response yang ditemukan untuk pertanyaan: ' + question)
    }
    
    // Kembali ke konten utama
    WebUI.switchToDefaultContent()
}

// Menyimpan file Excel output
FileOutputStream fileOut = new FileOutputStream(outputExcelFilePath)

outputWorkbook.write(fileOut)

fileOut.close()

outputWorkbook.close()

println('Semua pertanyaan dan respons terakhir telah disimpan ke ' + outputExcelFilePath)

