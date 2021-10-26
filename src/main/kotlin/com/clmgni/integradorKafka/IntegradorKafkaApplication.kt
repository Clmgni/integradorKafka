package com.clmgni.integradorKafka

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import java.io.FileInputStream
import java.io.FileOutputStream


@SpringBootApplication
class IntegradorKafkaApplication

fun main(args: Array<String>) {
	runApplication<IntegradorKafkaApplication>(*args)
	var teste1 = teste("xxx",1)
	var teste2 = teste2()
	println(teste2)
	var teste = writeToExcelFile("e:/planilhateste.xlsx")
	println(teste)
}

/**
 * Writes the value "TEST" to the cell at the first row and first column of worksheet.
 */
fun writeToExcelFile(filepath: String) {
	//Instantiate Excel workbook:

	val xlWb = XSSFWorkbook()
	//Instantiate Excel worksheet:
	val xlWs = xlWb.createSheet()

	//Row index specifies the row in the worksheet (starting at 0):
	val rowNumber = 0
	//Cell index specifies the column within the chosen row (starting at 0):
	val columnNumber = 0

	//Write text value to cell located at ROW_NUMBER / COLUMN_NUMBER:

	var cliente1 = listOf("Aaa","1","desc1")
	var cliente2 = listOf("Baa","2","desc2")
	var cliente3 = listOf("Caa","3","desc3")

	var clientes = listOf(cliente1,cliente2,cliente3)

	//Header
	val Linha = xlWs.createRow(0)
	Linha.createCell(0).setCellValue("Nome")
	Linha.createCell(1).setCellValue("Codigo")
	Linha.createCell(2).setCellValue("Descricao")
	Linha.createCell(3).setCellValue("Status")

	xlWs.createFreezePane(0, 1);
	val isLocked: CellStyle = xlWb.createCellStyle()
	isLocked.locked = false

	//Permite que a coluna D seja editável
	xlWs.setDefaultColumnStyle(3 , isLocked)
	xlWs.protectSheet("Teste")
	var rowIdx = 1
	for (cliente in clientes) {
		val Linha = xlWs.createRow(rowIdx++)

		Linha.createCell(0).setCellValue(cliente.get(0))
		Linha.createCell(1).setCellValue(cliente.get(1))
		Linha.createCell(2).setCellValue(cliente.get(2))
	}

	//Write file:
	val outputStream = FileOutputStream(filepath)
	xlWb.write(outputStream)
	xlWb.close()
}

/**
 * Reads the value from the cell at the first row and first column of worksheet.
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
	println(xlWs.getRow(rowNumber).getCell(columnNumber))
}


fun teste2(): Any {

	return true
}

fun teste(campo1: String,campo2: Int): Boolean {
	var Campo1: String
	var Campo2: Int

	println("Teste: $campo1 ")
	return true
}