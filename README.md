# aspoeJavaDemo
[Apache License 2.0](https://github.com/easonlau02/asposeJavaDemo/blob/master/LICENSE)
### How to read checkbox object from excel via [Aspose.Cells](https://products.aspose.com/cells/java)(Need license)
> If you come here, maybe you encounter issue about read Active X object,like checkbox from Excel via Apache POI.
Yes, now Apache POI can not support type of Active X.
Per research, I found one paid product can resovle your problem. Here give your a sample.

* Sample
![Sample](https://raw.githubusercontent.com/easonlau02/asposeJavaDemo/master/images/sample.png "Sample")
* Sample Code
```
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
        ExcelManipulate excelManipulate = new ExcelManipulate("./activex.xls");
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
```
* Result
![Result](https://raw.githubusercontent.com/easonlau02/asposeJavaDemo/master/images/result.png "Result")
### Words in the End
[Aspose.Cells](https://products.aspose.com/cells/java) is a good product of Aspose. If you are available, [purchase it for license](https://purchase.aspose.com/pricing/cells/java)

