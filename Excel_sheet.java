/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package MotorPH.Lab_Work_2;
import java.io.File;
import java.util.LinkedList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
/**
 *
 * @author Giann Gernale
 */
public class Excel_sheet {
    public static void main(String[] args) {
        String xlsxFile = "C:\\Users\\Giann Gernale\\Downloads\\MotorPH Employee Data.xlsx";
        LinkedList<String[]> dataList = new LinkedList<String[]>();

        try (Workbook workbook = WorkbookFactory.create(new File(xlsxFile))) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String[] rowData = new String[row.getLastCellNum()];
                for (Cell cell : row) {
                    int columnIndex = cell.getColumnIndex();
                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                        rowData[columnIndex] = Double.toString(cell.getNumericCellValue());
                    } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                        evaluator.evaluateFormulaCell(cell);
                        rowData[columnIndex] = Double.toString(cell.getNumericCellValue());
                    } else {
                        rowData[columnIndex] = cell.getStringCellValue();
                    }
                }
                dataList.add(rowData);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Print the LinkedList
        int rowIndex;   
        int columnIndex;
        System.out.println ("Which row?: ");
        Scanner scanner = new Scanner(System.in);
        System.out.println("Which colum?: ");
        Scanner scanner1 = new Scanner (System.in);
        int Identifier = scanner.nextInt();
        int Identifier1 = scanner1.nextInt();
        rowIndex = Identifier;
        columnIndex = Identifier1;
        String [] row = dataList.get(rowIndex);
        String data = row [columnIndex];
        
        System.out.println(data);
    }
}
