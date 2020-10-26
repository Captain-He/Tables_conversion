import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import readTables.readTables;

public class tool {
    public static String findSinglCell(String fileDir,int sheetNum, int rowNum, int colNum ){
        readTables readdevtables = new readTables();
        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        wb = readdevtables.readExcel(fileDir);
        String aimStr = "";
        if (wb != null) {
            sheet = wb.getSheetAt(sheetNum);
            row = sheet.getRow(rowNum);
            if (row != null) {

                String cellData = (String) readdevtables.getCellFormatValue(row
                        .getCell(colNum));
                if (cellData != null && cellData.length() > 0) {
                    aimStr = cellData;
                }

            }
            aimStr += "\r\n";
        }
        return aimStr;
    }
}
