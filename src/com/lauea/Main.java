package com.lauea;

import com.aspose.cells.CheckBox;
import com.aspose.cells.CheckBoxCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.Iterator;

public class Main {
    public static String INPUTPATH = "./activex.xls";
    public static void main(String[] args) {
        try{
            Workbook asposeWb = new Workbook(INPUTPATH);
            Worksheet ws = asposeWb.getWorksheets().get(0);
            CheckBoxCollection cbc = ws.getCheckBoxes();
            Iterator<CheckBox> iterator = cbc.iterator();
            while (iterator.hasNext()){
                CheckBox cb = iterator.next();
                System.out.println(cb.getText() + " = "+cb.getValue());
            }
        }catch (Exception ex) {
            System.out.println(ex.getMessage());
        }

    }
}
