package pl.docmanagementsystem.documentmanagementsystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelReplace {

    public static void main(String[] args) {
        String directoryPath = "C:\\Users\\kotowskim\\Desktop\\stamper\\src\\test\\resources\\document-templates"; // główny katalog
        Map<String, String> replacements = new HashMap<>();
        replacements.put("getDateAndPlaceOfEmbarkation().port", "getPortByPuid(signingOnPortPuid).name");
        replacements.put("getDateAndPlaceOfDisembarkation().port", "getPortByPuid(signingOffPortPuid).name");
        replacements.put("getDateAndPlaceOfDisembarkation().noadPort", "getPortByPuid(signingOffPortPuid).noadName");
        replacements.put("getDateAndPlaceOfEmbarkation().noadPort", "getPortByPuid(signingOnPortPuid).noadName");
        replacements.put("getDateAndPlaceOfEmbarkation().date", "getPortByPuid(signingOnPortPuid).date");
        replacements.put("getDateAndPlaceOfDisembarkation().date", "getPortByPuid(signingOffPortPuid).date");
        replacements.put("getDateAndPlaceOfEmbarkation().country", "getPortByPuid(signingOnPortPuid).country");
        replacements.put("getDateAndPlaceOfDisembarkation().country", "getPortByPuid(signingOffPortPuid).country");
        replacements.put("getDateAndPlaceOfDisembarkation().state", "getPortByPuid(signingOffPortPuid).state");
        replacements.put("getDateAndPlaceOfEmbarkation().state", "getPortByPuid(signingOnPortPuid).state");

        try (Stream<Path> paths = Files.walk(Paths.get(directoryPath))) {
            paths.filter(Files::isRegularFile)
                .filter(path -> path.toString().endsWith(".xlsx") || path.toString().endsWith(".xlsm"))
                .forEach(path -> {
                    try {
                        if (replacePhrasesInXlsx(path.toString(), replacements)) {
                            System.out.println("Phrases replaced in: " + path.toString());
                        } else {
                            System.out.println("No phrases found to replace in: " + path.toString());
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static boolean replacePhrasesInXlsx(String filePath, Map<String, String> replacements) throws IOException {
        boolean changesMade = false;

        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fis);

        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        for (Map.Entry<String, String> entry : replacements.entrySet()) {
                            if (cellValue.contains(entry.getKey())) {
                                cellValue = cellValue.replace(entry.getKey(), entry.getValue());
                                changesMade = true;
                            }
                        }
                        cell.setCellValue(cellValue);
                    }
                }
            }
        }

        fis.close();

        if (changesMade) {
            FileOutputStream fos = new FileOutputStream(new File(filePath));
            workbook.write(fos);
            fos.close();
        }

        workbook.close();
        return changesMade;
    }

}
