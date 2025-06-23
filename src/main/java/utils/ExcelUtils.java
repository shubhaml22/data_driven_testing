package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;


import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

    public static FileInputStream fi;
    public static FileOutputStream fo;
    public static XSSFWorkbook wb;
    public static XSSFSheet ws;
    public static XSSFRow row;
    public static XSSFCell cell;
    public static XSSFCellStyle style;

    public static int getRowCount(String xlFile,String xlSheet) throws IOException {

        fi=new FileInputStream(xlFile);
        wb=new XSSFWorkbook(fi);
        ws= wb.getSheet(xlSheet);
        int rowcount=ws.getLastRowNum();

        wb.close();
        fi.close();
        return rowcount;
    }

//    public static int getCellCount(String xlFile,String xlSheet,int rowNum) throws IOException {
//        fi=new FileInputStream(xlFile);
//        wb=new XSSFWorkbook(fi);
//        ws= wb.getSheet(xlSheet);
//        row=ws.getRow(rowNum);
//        int cellcount=row.getLastCellNum();
//        wb.close();
//        fi.close();
//        return cellcount;
//
//    }
    public static String getCellData(String xlFile,String xlSheet,int rowNum,int cellNum) throws IOException {
        fi=new FileInputStream(xlFile);
        wb=new XSSFWorkbook(fi);
        ws=wb.getSheet(xlSheet);
        row=ws.getRow(rowNum);
        cell=row.getCell(cellNum);

        String data;
        try {
//            data= cell.toString();
            DataFormatter formatter=new DataFormatter();
            data=formatter.formatCellValue(cell);//return the formating value of a cell  as a String regardless of the cell type.
        }
        catch (Exception e){
            data=" ";
        }
        wb.close();
        fi.close();
        return data;

    }

//    public static void setCellData(String xlFile,String xlSheet,int rowNum,int cellNum,String data) throws IOException {
//
//        fi=new FileInputStream(xlFile);
//        wb=new XSSFWorkbook(fi);
//        ws=wb.getSheet(xlSheet);
//        row=ws.getRow(rowNum);
//        cell=row.createCell(cellNum);
//        cell.setCellValue(data);
//
//        fo=new FileOutputStream(xlFile);
//        wb.write(fo);
//
//        wb.close();
//        fi.close();
//        fo.close();
//
//    }

    public static void fillGreenColor(String xlFile,String xlSheet,int rowNum,int cellNum) throws IOException {
        fi=new FileInputStream(xlFile);
        wb=new XSSFWorkbook(fi);
        ws=wb.getSheet(xlSheet);
        row=ws.getRow(rowNum);
        cell=row.getCell(cellNum);

        style=wb.createCellStyle();

        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell.setCellStyle(style);
        fo=new FileOutputStream(xlFile);
        wb.write(fo);
        wb.close();
        fi.close();
        fo.close();

    }
    public static void fillRedColor(String xlFile,String xlSheet,int rowNum,int cellNum) throws IOException {
        fi=new FileInputStream(xlFile);
        wb=new XSSFWorkbook(fi);
        ws=wb.getSheet(xlSheet);
        row=ws.getRow(rowNum);
        cell=row.getCell(cellNum);

        style=wb.createCellStyle();

        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell.setCellStyle(style);
        fo=new FileOutputStream(xlFile);
        wb.write(fo);
        wb.close();
        fi.close();
        fo.close();

    }

}


