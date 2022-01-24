import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

public class DoWork {

    private static final SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");

    private List<String> numbers = new ArrayList<>();
    private List<String> result = new ArrayList<>();

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private XSSFRow row;
    private XSSFCell cell;

    //Берем числа (текстовые переменные) из определенной колонки в таблице
    public void dataToFind(String path, int index) throws IOException {

        workbook = (XSSFWorkbook) WorkbookFactory.create(new File(path));
        sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();

        for (int r = 0; r <= rows; r++) {

            row = sheet.getRow(r);

            String cellData = getCellText(row.getCell(index)).trim();

            if (!cellData.equals("")) {
                numbers.add(cellData);
            }
        }
        workbook.close();
    }


    public void search(String fromPath, String toPath) throws IOException {

        Stream<Path> filePathStream = Files.walk(Paths.get(fromPath));

        filePathStream
                .forEach(filePath -> {

                    String fileNameStr = filePath.getFileName().toString().trim();

                    try {

                        for (String number : numbers) {
                            if (fileNameStr.contains(number)) {

                                result.add(number);

                                String crtFolder = toPath + "\\" + number;
                                String crtInFolder = crtFolder + "\\dir";

                                if (!Files.exists(Paths.get(crtFolder))) {
                                    Files.createDirectory(Paths.get(crtFolder));
                                }

                                if (Files.isDirectory(filePath)) {
                                    FileUtils.copyDirectory(filePath.toFile(),
                                            new File(crtInFolder));
                                }

                                if (Files.isRegularFile(filePath)) {
                                    if (!Files.exists(Paths.get(crtInFolder + "\\" + filePath.getFileName()))) {
                                        Files.copy(filePath, Path.of(crtFolder + "\\" + fileNameStr));
                                    }
                                }
                            }
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });

        numbers.clear();
    }


    void setResult(String toPath) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Отчет");


        for (int i = 0; i < result.size(); i++) {
            row = sheet.createRow(i);
            cell = row.createCell(0);
            cell.setCellValue(result.get(i));
        }

        FileOutputStream outstream = new FileOutputStream(toPath + "\\Отчет.xlsx");
        workbook.write(outstream);
        outstream.close();

        result.clear();


    }


    private String getCellText(Cell cell) {

        String result = "";

        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    result = cell.getRichStringCellValue().getString();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        result = sdf.format(cell.getDateCellValue());
                    } else {
                        result = Double.toString(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    result = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    result = cell.getCellFormula();
                    break;
                case BLANK:
                    result = "";
                    break;
                default:
                    System.out.println("Что-то пошло не так");
            }
        }
        return result;
    }
}
