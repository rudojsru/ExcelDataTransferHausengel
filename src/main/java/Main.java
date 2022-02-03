import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Main {
    public static void main(String[] args) throws IOException {


        String fileLocation1 = "C:/Users/OneDrive/Pulpit/Temp/BGW_Kündigungen.xlsx";
        String fileLocation2 = "C:/Users/OneDrive/Pulpit/Temp/Selb.xlsx";


        Map<String, List<String>> data1 = getStringListMap(fileLocation1);
        Map<String, List<String>> data2 = getStringListMap(fileLocation2);

        for (Map.Entry<String, List<String>> entry1 : data1.entrySet()) {
            for (Map.Entry<String, List<String>> entry2 : data2.entrySet()) {
                if (entry1.getKey().equals(entry2.getKey())) {
                    List list = data1.get(entry1.getKey());
                    String newVariable = entry2.getValue().get(1);
                    list.add(newVariable.substring(0, newVariable.length() - 2));
                    data1.put(entry1.getKey(), list);
                }
            }
        }

        wreiter(data1);
    }

    private static void wreiter(Map<String, List<String>> data1) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("All_Data ");
        XSSFRow row;
        Set<String> keyid = data1.keySet();

        int rowid = 0;

        for (String key : keyid) {

            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = data1.get(key).toArray();
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
            }
        }
        FileOutputStream out = new FileOutputStream(
                new File("C:/Users/OneDrive/Pulpit/Temp/BGW_ALL.xlsx"));

        workbook.write(out);
        out.close();
    }

    private static Map<String, List<String>> getStringListMap(String fileLocation) throws IOException {
        Sheet sheet = InputFile.reader(fileLocation);
        Map<String, List<String>> data1 = new LinkedHashMap <>();

        for (Row row : sheet) {
            List list = new ArrayList();
            for (Cell cell : row) {

                switch (cell.getCellType()) {
                    case STRING:
                        list.add(cell.getStringCellValue() + "");
                        break;
                    case NUMERIC:
                        list.add(cell.getNumericCellValue() + "");
                        break;
                }
            }
            String get0 = list.get(0).toString();
            get0 = get0.substring(0, get0.length() - 2);
            list.set(0, get0);

            data1.put((String) list.get(0), list);
        }
        return data1;
    }
}
