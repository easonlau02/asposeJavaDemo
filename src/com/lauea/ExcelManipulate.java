package com.lauea;

import com.aspose.cells.CheckBox;
import com.aspose.cells.CheckBoxCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.Iterator;

public class ExcelManipulate {
    public String inputPath;
    public ExcelManipulate(String inputPath){
        this.inputPath = inputPath;
    }
    public static void main(String[] args) {
        ExcelManipulate excelManipulate = new ExcelManipulate("./samples/activex.xls");
        try{
            // initial woorkbook
            Workbook asposeWb = new Workbook(excelManipulate.inputPath);
            // get worksheet 0
            Worksheet ws = asposeWb.getWorksheets().get(0);
            // get all checkbox from worksheet 0
            CheckBoxCollection cbc = ws.getCheckBoxes();
            Iterator<CheckBox> iterator = cbc.iterator();
            while (iterator.hasNext()){
                CheckBox cb = iterator.next();
                // getText == checkbox name, getValue == checked/unchecked
                System.out.println(cb.getText() + " = "+cb.getValue());
            }
        }catch (Exception ex) {
            System.out.println(ex.getMessage());
        }

    }
}
