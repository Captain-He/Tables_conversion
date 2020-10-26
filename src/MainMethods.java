import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import readTables.readTables;

import javax.swing.*;
import java.io.*;

public class MainMethods {

    public static String readTables(String dir) {
        readTables readdevtables = new readTables();
        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        wb = readdevtables.readExcel(dir);
        String aimStr = "";

        int num = 0;
        if (wb != null) {
            // ��ȡ��һ��sheet
            sheet = wb.getSheetAt(9);
            int rownum = sheet.getPhysicalNumberOfRows();
            //�ڶ��еڶ��� �� �ڶ��� �� ʮ����
            row = sheet.getRow(1);
            if (row != null) {

                for (int k = 1; k < 12; k++) {
                    String cellData = (String) readdevtables.getCellFormatValue(row
                            .getCell(k));
                    if (cellData != null && cellData.length() > 0) {
                        aimStr += cellData + " ";
                        num++;
                    }
                }
            }
            aimStr= aimStr.substring(0, aimStr.length() - 1);
            aimStr +="\r\n";

            int clNum = 0;
            for(int i=2;i<sheet.getPhysicalNumberOfRows();i++){
                Row row1 = sheet.getRow(i);
                row1.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
                String cellData = (String) readdevtables.getCellFormatValue(row1.getCell(0));
                if (cellData != null && cellData.length() > 0&&isEquals(cellData,"��ҵҪ���ɶ�")) {
                    clNum = i;
                    break;
                }
            }
            row = sheet.getRow(clNum);
            if (row != null) {

                for (int k = 1; k <= num; k++) {
                    String cellData = (String) readdevtables.getCellFormatValue(row
                            .getCell(k));
                    if (cellData != null && cellData.length() > 0) {
                        aimStr += formatDouble(cellData) + " ";
                    }
                }
            }
        }
        return aimStr.substring(0, aimStr.length() - 1);
    }

    public static void WriteToFile(String str, String filePath)
            throws IOException {
        BufferedWriter bw = null;
        try {
            FileOutputStream out = new FileOutputStream(filePath, false);// true,����??:????����???????????????��??,??????false
            bw = new BufferedWriter(new OutputStreamWriter(out, "GBK"));
            bw.write(str.substring(0, str.length() - 1));// ????
            bw.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //System.out.println("ת�����");
            bw.close();
        }
    }
    public static String formatDouble(String d){
        double b = Double.valueOf(d);
        return String.format("%.3f",b);
    }
    public static boolean isEquals(String a,String b){
        return a.replaceAll("\n", "").replaceAll(" ", "").equals(b);
    }
}
