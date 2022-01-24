import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {


        DoWork doWork = new DoWork();

        String excelPath = "D:\\TestDir\\Поиск.xlsx";
        String fromPath = "D:\\TestDir\\1. From";
        String toPath = "D:\\TestDir\\2 To";

        doWork.dataToFind(excelPath, 0);
        doWork.search(fromPath, toPath);
        doWork.setResult(toPath);
    }
}
