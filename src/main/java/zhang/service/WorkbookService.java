package zhang.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import zhang.model.ResultModel;
import zhang.model.SpecInfo;
import zhang.model.VariableInfo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@Slf4j
public class WorkbookService {
    private static final String QNAM = "QNAM";
    private static final String QLABEL = "QLABEL";

    public XSSFWorkbook generateExcelOutputDatafile(List<SpecInfo> specInfos) throws IOException, InvalidFormatException {
        final XSSFWorkbook workbook = new XSSFWorkbook();
        final XSSFSheet suppXSheet = workbook.createSheet(specInfos.get(0).getProject());
        suppXSheet.setColumnWidth(0, 256*8);
        suppXSheet.setColumnWidth(1, 256*15);
        suppXSheet.setColumnWidth(2, 256*50);
        int index = 0;
        for (SpecInfo specInfo : specInfos) {
            XSSFRow row = suppXSheet.createRow(index);
            addStringCell(row, 0, "Domain");
            addStringCell(row, 1, QNAM);
            addStringCell(row, 2, QLABEL);
            for (VariableInfo variableInfo : specInfo.getVariableInfos()) {
                row = suppXSheet.createRow(++index);
                addStringCell(row, 0, specInfo.getDomainName());
                addStringCell(row, 1, variableInfo.getName());
                addStringCell(row, 2, variableInfo.getLabel());
            }
            index+=2;
        }
        return workbook;
    }

    private void addStringCell(XSSFRow row, int index, String value) {
        final XSSFCell filenameCell = row.createCell(index);
        filenameCell.setCellValue(value);
    }

    public ResultModel parse(String dirPath) throws IOException {
        List<String> errorMessage = new ArrayList<>();
        List<SpecInfo> specInfos = new ArrayList<>();
        Path dir = Paths.get(dirPath);
        if (Files.exists(dir)) {
            try (DirectoryStream<Path> dataPaths = Files.newDirectoryStream(dir, path -> !path.getFileName().toString().startsWith("~$")
                    && path.getFileName().toString().endsWith(".xlsx"))) {
                dataPaths.forEach(dataPath -> {
                    final String fileName = dataPath.getFileName().toString();
                    System.out.println(String.format("fileName: %s", fileName));
                    if (fileName.contains("_")) {
                        try (Workbook workbook = WorkbookFactory.create(dataPath.toFile())) {
                            final String[] split = fileName.split("_");
                            String project = split[0];
                            System.out.println(String.format("project: %s", project));
                            String domainName = split[1].split("\\.")[0];
//                            System.out.println(String.format("domainName: %s", domainName));
                            String sheetName = String.format("Supp%s", domainName);
                            final Sheet sheet = workbook.getSheet(sheetName);
                            if (sheet != null) {
                                final ArrayList<VariableInfo> variableInfos = new ArrayList<>();
                                final SpecInfo specInfo = new SpecInfo(fileName, project, domainName, variableInfos);
                                boolean found = false;
                                for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                                    final Row row = sheet.getRow(i);
                                    if (row != null) {
                                        final Cell firstCell = row.getCell(0);
                                        final Cell secondCell = row.getCell(1);
                                        if (firstCell != null && secondCell != null) {
                                            if (QNAM.equals(getCellStringValue(firstCell)) &&
                                                    QLABEL.equals(getCellStringValue(secondCell))) {
                                                System.out.println(String.format("find QNAME and QLABEL row, row index = '%d'", i));
                                                found = true;
//                                                System.out.println("QNAM\t\tQLABEL");
                                                continue;
                                            }
                                            if (found) {
//                                                System.out.println(String.format("%s\t\t%s", getCellStringValue(firstCell), getCellStringValue(secondCell)));
                                                variableInfos.add(new VariableInfo(getCellStringValue(firstCell), getCellStringValue(secondCell)));
                                            }
                                        }
                                    }
                                }
                                if(variableInfos.size() > 0){
                                    variableInfos.sort(Comparator.comparing(VariableInfo::getName));
                                    System.out.println(String.format("variable size: %s", variableInfos.size()));
                                    specInfos.add(specInfo);
                                }else{
                                    errorMessage.add(String.format("could not find any variables, sheet = '%s', filename = '%s'", sheetName, fileName));
                                }
                            } else {
                                errorMessage.add(String.format("could not find target sheet, sheet = '%s', filename = '%s'", sheetName, fileName));
                            }
                        } catch (Exception e) {
                            errorMessage.add(String.format("read xlsx failed, filename = '%s'", fileName));
                            log.error(e.getMessage(), e);
                        }
                    } else {
                        errorMessage.add(String.format("invalid filename, filename = '%s'", fileName));
                    }
                    System.out.println();
                });
            }
        }
        return new ResultModel(errorMessage, specInfos);
    }

    private String getCellStringValue(Cell cell) {
        if (cell.getCellType().equals(CellType.STRING)) {
            return cell.getStringCellValue();
        }
        return "";
    }
}
