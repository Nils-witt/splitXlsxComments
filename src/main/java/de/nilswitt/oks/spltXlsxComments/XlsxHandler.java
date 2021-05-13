package de.nilswitt.oks.spltXlsxComments;

import de.nilswitt.oks.spltXlsxComments.model.Comment;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class XlsxHandler {

    private final String fileName;
    private final Workbook workbook;
    private final Sheet sheet;
    private final ArrayList<Comment> comments = new ArrayList<>();
    private HashMap<String, ArrayList<Comment>> map = new HashMap<>();

    public XlsxHandler(String fileName) throws IOException {
        this.fileName = fileName;
        workbook = new XSSFWorkbook(fileName);
        sheet = workbook.getSheetAt(0);
    }

    public void read() {
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row row;

        rowIterator.next();
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            try {
                String value = row.getCell(3).getStringCellValue();

                value = value.replaceAll("\n", " ").replaceAll("\r", " ");
                value = value.replaceAll("\"", "'");

                Comment comment = new Comment((int) row.getCell(0).getNumericCellValue(), row.getCell(1).getStringCellValue(), row.getCell(2).getStringCellValue(), value);
                comments.add(comment);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public void splitToRecipients() {
        map = new HashMap<>();

        for (Comment comment : comments) {

            ArrayList<Comment> userList = map.get(comment.getForUser());

            if (userList == null) {
                userList = new ArrayList<>();
            }

            userList.add(comment);

            map.put(comment.getForUser(), userList);
        }


    }

    public void createRecipientFiles() {
        String[] header = new String[]{"id", "author", "recipient", "value"};

        map.forEach((key, value) -> {
            String fileName = key.replaceAll(",", "").replaceAll(" ", "-").replaceAll(",", "-");
            String xlsxPath = String.format("./data/out/xlsx/%s.xlsx", fileName);
            String csvPath = String.format("./data/out/csv/%s.csv", fileName);

            XSSFWorkbook destbook = new XSSFWorkbook();
            Sheet destSheet = destbook.createSheet();

            Row row = destSheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                row.createCell(i).setCellValue(header[i]);
            }

            FileWriter out = null;
            CSVPrinter csvPrinter = null;
            try {
                out = new FileWriter(csvPath);

                csvPrinter = CSVFormat.DEFAULT.withDelimiter(';').withHeader("id", "author","recipient","value").print(out);
            } catch (IOException ioException) {
                ioException.printStackTrace();
            }

            for (Comment comment : value) {
                row = destSheet.createRow(row.getRowNum() + 1);
                row.createCell(0).setCellValue(comment.getId());
                row.createCell(1).setCellValue(comment.getAuthor());
                row.createCell(2).setCellValue(comment.getForUser());
                row.createCell(3).setCellValue(comment.getValue());
                if(csvPrinter != null){
                    try {
                        csvPrinter.printRecord(comment.getId(),comment.getAuthor(),comment.getForUser(),comment.getValue());
                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }
                }
            }

            try {
                FileOutputStream os = new FileOutputStream(xlsxPath);
                destbook.write(os);
                os.close();

                destbook.close();
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

        });
    }
}
