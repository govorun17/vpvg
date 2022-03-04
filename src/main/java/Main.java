import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        try {
//            String path2Exel = "files/exel.xlsx";
            String path2Exel = "files/exel20.xlsx";
            Workbook exel = new XSSFWorkbook(new FileInputStream(path2Exel));
            List<ParcedExel> dataset = new ArrayList<>();

            Sheet sheet = exel.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() >= 2) {
                    if (row.getCell(1).getStringCellValue() == null || row.getCell(1).getStringCellValue().isEmpty()) break;
                    dataset.add(new ParcedExel(
                            row.getCell(1).getStringCellValue(),
                            row.getCell(2).getStringCellValue(),
                            row.getCell(3).getStringCellValue(),
                            row.getCell(4).getStringCellValue(),
                            row.getCell(5).getStringCellValue(),
                            row.getCell(6).getStringCellValue(),
                            row.getCell(7).getStringCellValue()
                    ));
                }
            }
            exel.close();

            String path2Word = "files/word2.docx";

            for (ParcedExel parcedExel : dataset) {
                String newFile = "created/word" + parcedExel.getNum().replace('/', '-') + ".docx";

                XWPFDocument word = new XWPFDocument(new FileInputStream(path2Word));
                for (XWPFParagraph paragraph : word.getParagraphs()) {
                    List<XWPFRun> runs = paragraph.getRuns();
                    if (runs != null) {
                        for (XWPFRun run : runs) {
                            String text = run.getText(0);
                            if (text != null) {
                                if (text.contains("sdp")) {
                                    text = text.replace("sdp", parcedExel.getParsedDate());
                                    run.setText(text, 0);
                                }
                                if (text.contains("num")) {
                                    text = text.replace("num", parcedExel.getNum());
                                    run.setText(text, 0);
                                }
                                if (text.contains("vc")) {
                                    text = text.replace("vc", parcedExel.getVc());
                                    run.setText(text, 0);
                                }
                                if (text.contains("osnovanie")) {
                                    text = text.replace("osnovanie", parcedExel.getOsnovanie());
                                    run.setText(text, 0);
                                }
                                if (text.contains("sfera")) {
                                    text = text.replace("sfera", parcedExel.getSfera());
                                    run.setText(text, 0);
                                }
                                if (text.contains("sd")) {
                                    text = text.replace("sd", parcedExel.getStartDate());
                                    run.setText(text, 0);
                                }
                                if (text.contains("ed")) {
                                    text = text.replace("ed", parcedExel.getEndDate());
                                    run.setText(text, 0);
                                }
                                if (text.contains("of")) {
                                    text = text.replace("of", parcedExel.getIspolnitel());
                                    run.setText(text, 0);
                                }
                            }
                        }
                    }
                }
                word.write(new FileOutputStream(newFile));
                word.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
