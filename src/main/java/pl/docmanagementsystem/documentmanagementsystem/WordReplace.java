package pl.docmanagementsystem.documentmanagementsystem;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;

public class WordReplace {

    private static final String OLD_TEXT = "isArrival";
    private static final String NEW_TEXT = "XD";

    public static void main(String[] args) {
        String directoryPath = "C:\\Users\\kurzels\\IdeaProjects\\stamper\\src\\test\\resources\\word\\documents"; // ścieżka do katalogu głównego

        try {
            List<Path> files = Files.walk(Paths.get(directoryPath))
                .filter(path -> path.toString().endsWith(".docx"))
                .collect(Collectors.toList());

            for (Path path : files) {
                File file = path.toFile();
                boolean isModified = replaceInWordFile(file);
                if (isModified) {
                    System.out.println("Zmieniono plik: " + file.getAbsolutePath());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static boolean replaceInWordFile(File file) throws Exception {
        boolean isModified = false;

        try (FileInputStream fis = new FileInputStream(file)) {
            XWPFDocument document = new XWPFDocument(fis);

            // Przeszukiwanie zwykłych paragrafów
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                isModified |= replaceTextInRuns(paragraph.getRuns());
            }

            // Przeszukiwanie tabel
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            isModified |= replaceTextInRuns(paragraph.getRuns());
                        }
                    }
                }
            }

            // Przeszukiwanie komentarzy
            // nie działa jeszcze
            for (XWPFFootnote footnote : document.getFootnotes()) {
                for (XWPFParagraph paragraph : footnote.getParagraphs()) {
                    isModified |= replaceTextInRuns(paragraph.getRuns());
                }
            }

            if (isModified) {
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    document.write(fos);
                }
            }
        }

        return isModified;
    }

    private static boolean replaceTextInRuns(List<XWPFRun> runs) {
        boolean isModified = false;
        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text != null && text.contains(OLD_TEXT)) {
                isModified = true;
                text = text.replace(OLD_TEXT, NEW_TEXT);
                run.setText(text, 0);
            }
        }
        return isModified;
    }
}