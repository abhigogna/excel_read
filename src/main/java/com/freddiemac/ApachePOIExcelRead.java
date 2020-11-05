package com.freddiemac;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.StringJoiner;

public class ApachePOIExcelRead {

    private static final String FILE_NAME = "/Users/abhinavgogna/Downloads/Validation_Summary.xlsx";

    public static List<String> getExcelData(String fileName, int skipRows){
        List<String> fileData = new ArrayList<String>();
        FileInputStream excelFile = null;
        try {
            excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            StringJoiner sb ;
            int rowCount = 0;
            while (iterator.hasNext()) {
                rowCount++;
                Row currentRow = iterator.next();
                if (rowCount <= skipRows) {
                    continue;
                }
                sb = new StringJoiner(";");
                for (Cell currentCell : currentRow) {
                    if (currentCell.getCellType() == CellType.STRING) {
                        sb.add(currentCell.getStringCellValue());
                    } else if (currentCell.getCellType() == CellType.NUMERIC) {
                        sb.add(String.valueOf(currentCell.getNumericCellValue()));
                    }
                }
                fileData.add(sb.toString());
            }
            excelFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return fileData;
    }

    public static String convertListToHTMLTable(List<String> fileData){
        List<String> tableList = new ArrayList<>();
        StringBuffer buf = new StringBuffer();
        buf.append("<!DOCTYPE html>" +
                "<body>" +
                "<table>" +
                "<tr>" +
                "<th>"+fileData.get(0)+"</th>" +
                "</tr>");
        for (int i = 1; i < fileData.size(); i++) {
            String splitStr [] = fileData.get(i).split(";");
            buf.append("<tr><td>")
                    .append(splitStr[0])
                    .append("</td><td>")
                    .append(splitStr[1])
                    .append("</td></tr>");
        }
        buf.append("</table>" +
                "</body>" +
                "</html>");
        return buf.toString();
    }

    public static void main(String[] args) {
        System.out.println(getExcelData(FILE_NAME, 7));
        System.out.println(convertListToHTMLTable(getExcelData(FILE_NAME, 7)));
    }
}
