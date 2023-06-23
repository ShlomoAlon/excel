import androidx.compose.material.MaterialTheme
import androidx.compose.desktop.ui.tooling.preview.Preview
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.material.Button
import androidx.compose.material.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.window.Window
import androidx.compose.ui.window.application
import androidx.compose.ui.Modifier
import androidx.compose.ui.Alignment
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import com.darkrockstudios.libraries.mpfilepicker.FilePicker
import org.apache.poi.ss.usermodel.*
import org.apache.poi.*
import org.apache.poi.openxml4j.util.ZipSecureFile
import java.io.File
import java.io.FileInputStream
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

//SKU	Total Units	Case Count	Weight (lb)	Width (in)	Length (in)	Height (in)
data class Sku(
    val total_units: Int?,
    val case_count: Int,
    val weight: Double,
    val width: Double,
    val length: Double,
    val height: Double
)

fun create_sku_dic(path: String): Map<String, Sku> {
    var file = FileInputStream(path)
    var workbook = XSSFWorkbook(file)
    var sheet = workbook.getSheetAt(0)
    var sku_dic = mutableMapOf<String, Sku>()
    for (row in sheet) {
        if (row.getRowNum() == 0) {
            continue
        } else if (row.getCell(0) == null || row.getCell(0).stringCellValue == "") {
            break
        }
        var sku = row.getCell(0).stringCellValue
        println(sku)
        var case_count = row.getCell(2).numericCellValue.toInt()
        var weight = row.getCell(3).numericCellValue
        var width = row.getCell(4).numericCellValue
        var length = row.getCell(5).numericCellValue
        var height = row.getCell(6).numericCellValue
        sku_dic[sku] = Sku(null, case_count, weight, width, length, height)
    }
    return sku_dic
}

class ExcelFile(val path_to_file: String, path_to_information: String) {
    var skus = create_sku_dic(path_to_information)
    private var file = FileInputStream(path_to_file)
    var workbook = XSSFWorkbook(file)
    var sheet = workbook.getSheetAt(0)
    private val first_row = 5
    val first_column = 12
    var first_empty_row = 0
    init {
        for (row in sheet) {
            if (row.getRowNum() < first_row) {
                continue
            } else if (row.getCell(0) == null || row.getCell(0).stringCellValue == "") {
                first_empty_row = row.getRowNum()
                break
            }
        }
        for (i in first_row until first_empty_row ) {
            var row = sheet.getRow(i)
            for (cell in row){
                if (cell.columnIndex >= first_column){
                    cell.setBlank()
                }
            }
        }
    }
    val box_weight_row = first_empty_row + 2
    fun fill_info(){
        var current_case = first_column
        var current_row = first_row
        while (current_row < first_empty_row){
            val row = sheet.getRow(current_row)
            val sku = row.getCell(0).stringCellValue
            val sku_item = skus[sku]!!
            val expected_quantity = row.getCell(9).numericCellValue.toInt()
            for (i in 0 until expected_quantity / sku_item.case_count){
                val current_cell = row.getCell(current_case)
                current_cell.setCellValue(sku_item.case_count.toDouble())
                sheet.getRow(box_weight_row).getCell(current_case).setCellValue(sku_item.weight)
                sheet.getRow(box_weight_row + 1).getCell(current_case).setCellValue(sku_item.width)
                sheet.getRow(box_weight_row + 2).getCell(current_case).setCellValue(sku_item.length)
                sheet.getRow(box_weight_row + 3).getCell(current_case).setCellValue(sku_item.height)
                current_case += 1
            }
            if (expected_quantity % sku_item.case_count != 0){
                val current_cell = row.getCell(current_case)
                current_cell.setCellValue(expected_quantity % sku_item.case_count.toDouble())
                sheet.getRow(box_weight_row).getCell(current_case).setCellValue(sku_item.weight * (expected_quantity % sku_item.case_count) / sku_item.case_count.toDouble())
                sheet.getRow(box_weight_row + 1).getCell(current_case).setCellValue(sku_item.width)
                sheet.getRow(box_weight_row + 2).getCell(current_case).setCellValue(sku_item.length)
                sheet.getRow(box_weight_row + 3).getCell(current_case).setCellValue(sku_item.height)
                current_case += 1
            }
            current_row += 1
        }
        var fileOut = FileOutputStream(this.path_to_file)
        workbook.write(fileOut)
        print("done")
    }
}

@Composable
@Preview
fun App() {
    var text by remember { mutableStateOf("Hello, World!") }

    MaterialTheme {
        Column(
            modifier = Modifier.fillMaxSize(), horizontalAlignment = Alignment.CenterHorizontally,
            verticalArrangement = Arrangement.SpaceEvenly
        ) {
            var showFilePicker1 by remember { mutableStateOf(false) }
            var showFilePicker2 by remember { mutableStateOf(false) }
            var file_path1 by remember { mutableStateOf("the excel file to be modified has not been chosen") }
            var file_path2 by remember { mutableStateOf("the information file has not been chosen") }

            val fileType = "xlsx, xlx, xls"
            FilePicker(showFilePicker1, fileExtension = fileType) { path ->
                if (path != null) {
                    file_path1 = path
                }
                showFilePicker1 = false
            }
            FilePicker(showFilePicker2, fileExtension = fileType) { path ->
                if (path != null) {
                    file_path2 = path
                }
                showFilePicker2 = false
            }
            Button(
                onClick = {
                    showFilePicker1 = true
                },
                modifier = Modifier.fillMaxWidth().weight(1f),
                colors = androidx.compose.material.ButtonDefaults.buttonColors(backgroundColor = Color.Red)
            ) {
                Text(file_path1)
            }
            Button(onClick = {
                showFilePicker2 = true
            }, modifier = Modifier.fillMaxWidth().weight(1f)) {
                Text(file_path2)
            }
            var why = ""
            var button_text by remember { mutableStateOf("run the template filler") }
            Button(
                onClick = {
                    var template = ExcelFile(file_path1, file_path2)
                    template.fill_info()
                    button_text = "done"
                },
                modifier = Modifier.fillMaxWidth().weight(1f),
                colors = androidx.compose.material.ButtonDefaults.buttonColors(backgroundColor = Color.Green)
            ) {
                Text(button_text)
            }
        }
    }
}

fun main() = application {
    ZipSecureFile.setMinInflateRatio(0.00001)
    Window(onCloseRequest = ::exitApplication) {
        App()
    }
}
