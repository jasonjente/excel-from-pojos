package project.excelize;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Locale;

public class ExcelConverter {

    public static void convertPojoToSpreadsheet(List<Object> objectList){

        Workbook workbook = new XSSFWorkbook();
        String outFileName = "outputFile";
        Sheet sheet = workbook.createSheet(outFileName);
        //Create a simple header row with the field names of the parsed objects
        Row header = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(false);
        headerStyle.setFont(font);
        //gets all the declared fields of the reflected class and then creates a cell with the name of each attribute.
        Class<? extends Object> c = objectList.get(0).getClass();

        int i =0;
        //Get all available methods
        Method[] methods = c.getDeclaredMethods();
        for(Method method:methods){
            //TODO: optional use of arguments containing fields that we want to be excluded.
            //Use only the methods that start with the prefix get
            if(!method.getName().startsWith("get"))
                continue;
            String name = method.getName().substring(3);
            Cell headerCell = header.createCell(i);
            headerCell.setCellValue(name);
            headerCell.setCellStyle(headerStyle);
            method.setAccessible(true);
            //TODO: optional use of argument for width otherwise defaultValue
            sheet.setColumnWidth(i, 6000);
            i++;
        }
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        int rowIndex = 0;
        for(Object object:objectList){
            Row row = sheet.createRow(rowIndex+1);
            int columnIndex = 0;
            for(Method method:object.getClass().getDeclaredMethods()){
                Cell cell = row.createCell(columnIndex);
                String data = "";
                Object returnObject = null;
                if (method.getName().startsWith("get") && method.getGenericParameterTypes().length == 0) {

                    try {
                        returnObject = method.invoke(object);
                    } catch (IllegalAccessException | InvocationTargetException e) {

                    }
                    if(returnObject == null){
                        data = "";
                    } else {
                        data = returnObject.toString();
                    }
                    cell.setCellStyle(style);
                    cell.setCellValue(data);
                    columnIndex ++;
                }

            }
            rowIndex++;
        }

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        //TODO: optional use of argument for output name ortherwise default value ,something like objectName-date.xslx
        String fileLocation = path.substring(0, path.length() - 1) + "out.xlsx";

        FileOutputStream outputStream;
        try {
            outputStream = new FileOutputStream(fileLocation);
            workbook.write(outputStream);
        } catch (IOException e) {
        }finally {
            try {
                workbook.close();
            } catch (IOException e) {

            }
        }
    }

}
