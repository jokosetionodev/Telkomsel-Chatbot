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
import org.apache.commons.text.similarity.CosineSimilarity // Pastikan Anda menambahkan dependensi ini
import org.apache.commons.collections4.MapUtils // Tambahkan dependensi ini untuk memudahkan penghitungan frekuensi kata
import org.apache.commons.collections4.MultiValuedMap
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap

// Fungsi untuk menghitung cosine similarity
double calculateCosineSimilarity(String text1, String text2) {
	Map<String, Integer> freqMap1 = getWordFrequency(text1)
	Map<String, Integer> freqMap2 = getWordFrequency(text2)

	Set<String> uniqueWords = new HashSet<>(freqMap1.keySet())
	uniqueWords.addAll(freqMap2.keySet())

	double dotProduct = 0
	double magnitude1 = 0
	double magnitude2 = 0

	for (String word : uniqueWords) {
		int freq1 = freqMap1.getOrDefault(word, 0)
		int freq2 = freqMap2.getOrDefault(word, 0)

		dotProduct += freq1 * freq2
		magnitude1 += freq1 * freq1
		magnitude2 += freq2 * freq2
	}

	if (magnitude1 == 0 || magnitude2 == 0) {
		return 0 // untuk menghindari pembagian dengan nol
	}

	return dotProduct / (Math.sqrt(magnitude1) * Math.sqrt(magnitude2))
}

// Fungsi untuk mendapatkan frekuensi kata
Map<String, Integer> getWordFrequency(String text) {
	Map<String, Integer> frequencyMap = new HashMap<>()
	String[] words = text.toLowerCase().split("\\W+")
	
	for (String word : words) {
		frequencyMap.put(word, frequencyMap.getOrDefault(word, 0) + 1)
	}

	return frequencyMap
}

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
TestData questionsData = findTestData('Data Files/questions/questions') // Ganti dengan nama Test Data yang sesuai

int questionCount = questionsData.getRowNumbers()

// Temukan indeks kolom untuk "Quest" dan "Expected"
int questionColumnIndex = -1
int expectedColumnIndex = -1

// Mendapatkan semua nama kolom
List<String> columnNames = questionsData.getColumnNames()

for (int colIndex = 0; colIndex < columnNames.size(); colIndex++) {
	if (columnNames[colIndex].equalsIgnoreCase("Quest")) {
		questionColumnIndex = colIndex + 1 // Indeks di Katalon dimulai dari 1
	} else if (columnNames[colIndex].equalsIgnoreCase("Expected")) {
		expectedColumnIndex = colIndex + 1 // Indeks di Katalon dimulai dari 1
	}
}

if (questionColumnIndex == -1 || expectedColumnIndex == -1) {
	throw new RuntimeException("Kolom 'Quest' atau 'Expected' tidak ditemukan dalam Test Data.")
}

// Membuat workbook baru untuk menyimpan respons
String uuid = UUID.randomUUID().toString()

String outputExcelFilePath = ('C:\\Users\\User\\Documents\\hasilexcel\\response_' + uuid) + '.xlsx'

Workbook outputWorkbook = new XSSFWorkbook()

Sheet outputSheet = outputWorkbook.createSheet('Responses')

// Membuat header untuk output
Row headerRow = outputSheet.createRow(0)

headerRow.createCell(0).setCellValue('Pertanyaan')
headerRow.createCell(1).setCellValue('Response')
headerRow.createCell(2).setCellValue('Expected Answer')
headerRow.createCell(3).setCellValue('Similarity')

// Mengirim pertanyaan satu per satu
for (int i = 1; i <= questionCount; i++) {
	// Ambil pertanyaan dan jawaban yang diharapkan dari Test Data
	String question = questionsData.getValue(questionColumnIndex, i) // Ambil dari kolom "Quest"
	String expectedAnswer = questionsData.getValue(expectedColumnIndex, i) // Ambil dari kolom "Expected"

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
	int timeout = 5

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

		// Hitung similarity antara response terakhir dan expected answer
		double similarityScore = calculateCosineSimilarity(lastResponseText, expectedAnswer)

		// Tampilkan similarity score di console
		println('Similarity Score: ' + similarityScore)

		// Menyimpan pertanyaan, respons terakhir, jawaban yang diharapkan, dan similarity score ke sheet output
		Row outputRow = outputSheet.createRow(i)

		outputRow.createCell(0).setCellValue(question)
		outputRow.createCell(1).setCellValue(lastResponseText)
		outputRow.createCell(2).setCellValue(expectedAnswer)
		outputRow.createCell(3).setCellValue(similarityScore)
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
