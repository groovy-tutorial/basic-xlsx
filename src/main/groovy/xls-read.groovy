@Grapes([
    @Grab('org.apache.poi:poi:3.10.1'),
    @Grab('org.apache.poi:poi-ooxml:3.10.1')])
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import static org.apache.poi.ss.usermodel.Cell.*

import java.nio.file.Paths
import groovy.json.JsonOutput

def header = []
def values = []

Paths.get('countryInfo.xlsx').withInputStream { input ->

    def workbook = new XSSFWorkbook(input)    
    def sheet = workbook.getSheetAt(0)
    
    for (cell in sheet.getRow(0).cellIterator()) { 
        header << cell.stringCellValue
    }
    
    def headerFlag = true
    for (row in sheet.rowIterator()) {
        if (headerFlag) {
            headerFlag = false
            continue
        }
        def rowData = [:]
        for (cell in row.cellIterator()) {
            def value = ''
            
            switch(cell.cellType) {
                case CELL_TYPE_STRING:
                    value = cell.stringCellValue
                    break
                case CELL_TYPE_STRING:
                    value = cell.numericCellValue
                    break
                default:
                    value = ''                
            }
            rowData << ["${header[cell.columnIndex]}" : value]
        }
        values << rowData
    }
}

Paths.get('countryInfo.xlsx.json').withWriter { jsonWriter ->
    jsonWriter.write JsonOutput.prettyPrint(JsonOutput.toJson(values))
}