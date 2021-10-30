package com.clmgni.integradorKafka

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.xssf.usermodel.XSSFDataValidation
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import java.io.FileInputStream
import java.io.FileOutputStream


@SpringBootApplication
class IntegradorKafkaApplication

fun main(args: Array<String>) {
	runApplication<IntegradorKafkaApplication>(*args)

	//Gera Dados
	var cliente1 = listOf("Aaa","1","desc1")
	var cliente2 = listOf("Baa","2","desc2")
	var cliente3 = listOf("Caa","3","desc3")
	var clientes = listOf(cliente1,cliente2,cliente3)

	//Cria Planilha
	println("Gera planilha...")
	writeToExcelFile("e:/planilhateste.xlsx",clientes)

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
	//Instancia aba da planilha Excel
	val xlWs = xlWb.createSheet("Transferencias")
	val xlWs2 = xlWb.createSheet("Opcoes")

	//Dados
	var linha1 = listOf("Aaaaa","000001","desc1")
	var linha2 = listOf("Bbbbb","000002","desc2")
	var linha3 = listOf("Ccccc","000003","desc3")
	var linha4 = listOf("Ddddd","000004","desc4")
	var dados = listOf(linha1,linha2,linha3,linha4)

	// Monta Header
	val Linha = xlWs.createRow(0)
	Linha.createCell(0).setCellValue("Nome")
	Linha.createCell(1).setCellValue("Codigo")
	Linha.createCell(2).setCellValue("Descricao")
	Linha.createCell(3).setCellValue("Data")
	Linha.createCell(4).setCellValue("Status")

	// Monta dados da Lista Suspensa em outra planilha e a deixa oculta
	xlWs2.createRow(0).createCell(0).setCellValue("Opcoes")
	xlWs2.createRow(1).createCell(0).setCellValue("SIM")
	xlWs2.createRow(2).createCell(0).setCellValue("NAO")
	xlWb.setSheetHidden(1,true)

	//Congela primeira linha
	xlWs.createFreezePane(0, 1);
	val isLocked: CellStyle = xlWb.createCellStyle()
	isLocked.locked = false

	//Permite que a coluna D e E sejam editáveis
	xlWs.setDefaultColumnStyle(3 , isLocked)
	xlWs.setDefaultColumnStyle(4 , isLocked)
	xlWs.protectSheet("Teste")
	xlWs2.protectSheet("Teste")

	var rowIdx = 1
	for (linha in dados) {
		val Linha = xlWs.createRow(rowIdx++)
		Linha.createCell(0).setCellValue(linha.get(0))
		Linha.createCell(1).setCellValue(linha.get(1))
		Linha.createCell(2).setCellValue(linha.get(2))
	}

	var dataValidation: DataValidation? = null
	var constraint: DataValidationConstraint? = null
	var validationHelper: DataValidationHelper? = null

	// Coloca lista suspensa
	validationHelper = XSSFDataValidationHelper(xlWs)
	val addressList = CellRangeAddressList(1, rowIdx-1, 3, 3)
	//Opcao via Planilha auxiliar
	constraint = validationHelper.createFormulaListConstraint("Opcoes!A$2:A$3")
	//Opcao via string
	//constraint = validationHelper.createExplicitListConstraint(arrayOf("SIM", "NAO")
	dataValidation = validationHelper.createValidation(constraint, addressList)
	dataValidation.suppressDropDownArrow = true // Mostrar ListBox
	dataValidation.showErrorBox = true // Mostrar mensagem de erro
	dataValidation.createErrorBox("Validação dos dados","Este campo deve ser preenchido com 'SIM' ou 'NAO'!")
	dataValidation.errorStyle = 0 // ERROR=0,WARNING=1,INFO=2
	xlWs.addValidationData(dataValidation)


	//Grava Arquivo
	val outputStream = FileOutputStream(filepath)
	xlWb.write(outputStream)
	xlWb.close()
}

/**
 * LEITURA DE PLANILHA
 */
fun readFromExcelFile(filepath: String) {
	val inputStream = FileInputStream(filepath)
	//Instantiate Excel workbook using existing file:
	var xlWb = WorkbookFactory.create(inputStream)

	//Row index specifies the row in the worksheet (starting at 0):
	val rowNumber = 0
	//Cell index specifies the column within the chosen row (starting at 0):
	val columnNumber = 0

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
