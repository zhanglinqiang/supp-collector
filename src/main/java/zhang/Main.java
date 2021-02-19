package zhang;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import zhang.model.ResultModel;
import zhang.service.WorkbookService;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.util.List;

@Slf4j
public class Main {
    private static WorkbookService workbookService = new WorkbookService();

    public static void main(String[] args) {
        try {
            if(args.length < 1){
                System.err.println("参数错误, 未指定数据源目录");
                return;
            }
            String dirPath = args[0].trim();
            File file = new File(dirPath);
            if(!(file.exists())){
                System.err.println("参数错误, 数据源目录不存在");
                return;
            }
            if(!file.isDirectory()){
                System.err.println("参数错误, 数据源目录错误");
                return;
            }

            long start = System.currentTimeMillis();
//            final String dirPath = "C:\\Users\\leen\\Desktop\\data";
            System.out.println("数据目录: " + dirPath);
            final ResultModel parse = workbookService.parse(dirPath);

            if (parse.getSpecInfos().size() > 0) {
                final String outputFileName = String.format("%s.xlsx", parse.getSpecInfos().get(0).getProject());
                final File outputFile = Paths.get(new File(dirPath).getParent(), outputFileName).toFile();
                FileUtils.deleteQuietly(outputFile);
                try (final XSSFWorkbook workbook = workbookService.generateExcelOutputDatafile(parse.getSpecInfos());
                     final FileOutputStream fileOutputStream = new FileOutputStream(outputFile)) {
                    workbook.write(fileOutputStream);
                }
                System.out.println(String.format("耗时: %.2fs", (System.currentTimeMillis() - start) * 1.0 / 1000));
                System.out.println(String.format("报告路径： %s", outputFile.getAbsolutePath()));
            } else {
                log.error("could not get any specInfos, length = 0");
            }

            final List<String> errorMessage = parse.getErrorMessage();
            if (errorMessage.size() > 0) {
                System.out.println("******************错误文件列表******************");
                errorMessage.forEach(System.out::println);
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
    }
}
