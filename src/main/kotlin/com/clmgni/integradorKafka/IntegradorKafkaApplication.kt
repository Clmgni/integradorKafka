package com.clmgni.integradorKafka

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.openxml4j.opc.PackageAccess
import org.apache.poi.poifs.crypt.EncryptionInfo
import org.apache.poi.poifs.crypt.EncryptionMode
import org.apache.poi.poifs.crypt.Encryptor
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.*


@SpringBootApplication
class IntegradorKafkaApplication

fun main(args: Array<String>) {
	runApplication<IntegradorKafkaApplication>(*args)

	//Cria Planilha
	println("Gera planilha...")
	writeToExcelFile("e:/planilhateste.xlsx","Informacoes")

	//Le planilha
	println("Le planilha...")
	readFromExcelFile("e:/planilhateste.xlsx")

}

/**
 * GERACAO DE PLANILHA EXCEL
 */
fun writeToExcelFile(filepath: String, dados: Any) {
	//Instancia Excel workbook:

	val xlWb = XSSFWorkbook()
	Biff8EncryptionKey.setCurrentUserPassword("userpass")

	//Instancia aba da planilha Excel
	val xlWs = xlWb.createSheet("Transferencias")
	val xlWs2 = xlWb.createSheet("Opcoes")

	//Dados
	var linha1 = listOf("TED   ","000001","desc1",Date().toString())
	var linha2 = listOf("TED   ","000002","desc2",Date().toString())
	var linha3 = listOf("PIX   ","000003","desc3",Date().toString())
	var linha4 = listOf("PICPAY","000004","desc4",Date().toString())
	var dados = listOf(linha1,linha2,linha3,linha4)

	// Monta Header
	val Linha = xlWs.createRow(0)
	Linha.createCell(0).setCellValue("PRODUTO")
	Linha.createCell(1).setCellValue("CODIGO")
	Linha.createCell(2).setCellValue("DESCRICAO")
	Linha.createCell(3).setCellValue("DATA")
	Linha.createCell(4).setCellValue("SUCESSO(SIM/NAO)")

	// Estilo da C??lula
	var style: CellStyle = xlWb.createCellStyle()
	// Setting Background color
	style.fillBackgroundColor = IndexedColors.GREEN.getIndex()
	style.fillPattern = FillPatternType.BIG_SPOTS

	// Setting Foreground Color
	style = xlWb.createCellStyle()
	style.fillForegroundColor = IndexedColors.BLUE.getIndex()
	style.fillPattern = FillPatternType.SOLID_FOREGROUND


	// Monta dados da Lista Suspensa em outra planilha e a deixa oculta
	xlWs2.createRow(0).createCell(0).setCellValue("Opcoes")
	xlWs2.createRow(1).createCell(0).setCellValue("SIM")
	xlWs2.createRow(2).createCell(0).setCellValue("NAO")
	xlWb.setSheetHidden(1,true)

	//Congela primeira linha
	xlWs.createFreezePane(0, 1);
	val isLocked: CellStyle = xlWb.createCellStyle()
	isLocked.locked = false

	//Permite que a coluna D e E sejam edit??veis
	//xlWs.setDefaultColumnStyle(3 , isLocked)
	xlWs.setDefaultColumnStyle(4 , isLocked)
	xlWs.protectSheet("Teste")
	xlWs2.protectSheet("Teste")

	//Protege Workbook
	xlWb.lockStructure()
//	xlWb.setWorkbookPassword("teste123",HashAlgorithm.sha512)


	var rowIdx = 1
	for (linha in dados) {
		val Linha = xlWs.createRow(rowIdx++)
		Linha.createCell(0).setCellValue(linha.get(0))
		Linha.createCell(1).setCellValue(linha.get(1))
		Linha.createCell(2).setCellValue(linha.get(2))
		Linha.createCell(3).setCellValue(linha.get(3))
	}

	var dataValidation: DataValidation? = null
	var constraint: DataValidationConstraint? = null
	var validationHelper: DataValidationHelper? = null

	// Coloca lista suspensa
	validationHelper = XSSFDataValidationHelper(xlWs)
	val addressList = CellRangeAddressList(1, rowIdx-1, 4, 4)
	//Opcao via Planilha auxiliar
	constraint = validationHelper.createFormulaListConstraint("Opcoes!A$2:A$3")
	//Opcao via string
	//constraint = validationHelper.createExplicitListConstraint(arrayOf("SIM", "NAO")
	dataValidation = validationHelper.createValidation(constraint, addressList)
	dataValidation.suppressDropDownArrow = true // Mostrar ListBox
	dataValidation.showErrorBox = true // Mostrar mensagem de erro
	dataValidation.createErrorBox("Valida????o dos dados","Este campo deve ser preenchido com 'SIM' ou 'NAO'!")
	dataValidation.errorStyle = 0 // ERROR=0,WARNING=1,INFO=2
	xlWs.addValidationData(dataValidation)

	// Auto ajuste das colunas
	xlWs.autoSizeColumn(0)
	xlWs.autoSizeColumn(1)
	xlWs.autoSizeColumn(2)
	xlWs.autoSizeColumn(3)
	xlWs.autoSizeColumn(4)

	// Auto filtro
	xlWs.setAutoFilter(CellRangeAddress(0,rowIdx-1,0,4))


	//Grava Arquivo
	//val outputStream = FileOutputStream(filepath)
	//xlWb.write(outputStream)

	val xlsxFile = File(filepath)

	//Biff8EncryptionKey.setCurrentUserPassword("teste123");
	val fos = FileOutputStream(xlsxFile)
	xlWb.write(fos)
	xlWb.close()

//////////////////////////////
	val fs = POIFSFileSystem()

	val info =  EncryptionInfo(EncryptionMode.agile)
	val encryptor = info.encryptor
	encryptor.confirmPassword("teste123")

	val opc = OPCPackage.open(xlsxFile)
	val os = encryptor.getDataStream(fs)
	opc.save(os)

	// Write out the encrypted version
	val fos2 = FileOutputStream(xlsxFile)
	fs.writeFilesystem(fos2)

////////////////////////////////////

}

/**
 * LEITURA DE PLANILHA
 */
fun readFromExcelFile(filepath: String) {
	val inputStream = FileInputStream(filepath)
	//Instantiate Excel workbook using existing file:
	var xlWb = WorkbookFactory.create(inputStream,"teste123")
	//var xlWb = WorkbookFactory.create(inputStream,"Password")

	//Get reference to first sheet:
	val xlWs = xlWb.getSheetAt(0)

	//Dimensoes da Planilha
	var ultimaLinha = xlWs.lastRowNum
	var linhaPlan=0
	var colunaPlan: Int
	var colunas = (xlWs.getRow(0).lastCellNum) - 1

	println("Possui " + ultimaLinha + " linhas e " + colunas + " colunas!")
	println("\n")

	//Lista Dados
	while ( linhaPlan <= ultimaLinha) {
		colunaPlan=0
		while (colunaPlan <=colunas) {
			print(xlWs.getRow(linhaPlan).getCell(colunaPlan))
			print("\t")
			colunaPlan++
		}
		print("\n")
		linhaPlan++
	}
}
