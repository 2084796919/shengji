package org.example.fileMove;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 通过文件名称 36、2025.05#791.pdf 匹配excel凭证号 “记791” 并且日期为 2025.05月数据，
 * 将匹配到的结果数据输出到审计结果.xlsx中
 */
public class DocumentScanner {
  private static final SimpleDateFormat FILE_DATE_FORMAT = new SimpleDateFormat("yyyy.MM");
  private static final SimpleDateFormat RECORD_DATE_FORMAT = new SimpleDateFormat("yyyy/MM/dd");

  public static void main(String[] args) {
    String sourceFolderPath = "C:\\Users\\20847\\Desktop\\c\\应付账款抽凭";
    String dataSearchFilePath = sourceFolderPath + "\\数据搜索.xlsx";
    String outputFilePath = sourceFolderPath + "\\审计结果.xlsx";

    try {
      // 1. 扫描文件夹并筛选有效文件（按1、2、3排序）
      List<File> validFiles = scanAndSortFiles(sourceFolderPath);
      if (validFiles.isEmpty()) return;

      // 2. 读取数据搜索文件
      List<Map<String, String>> dataSearchRecords = readDataSearchFile(dataSearchFilePath);

      // 3. 严格匹配：每个文件必须唯一对应一条数据
      List<MatchResult> matchResults = new ArrayList<>();
      List<String> unmatchedFiles = new ArrayList<>();
      List<String> multiMatchFiles = new ArrayList<>();

      for (File file : validFiles) {
        String fileName = file.getName();
        String voucherNumber = "记" + fileName.substring(fileName.indexOf("#") + 1, fileName.indexOf(".pdf"));
        String fileYearMonth = fileName.substring(fileName.indexOf("、") + 1, fileName.indexOf("#"));

        // 匹配数据
        List<Map<String, String>> matchedRecords = findMatchedRecords(
                dataSearchRecords, voucherNumber, fileYearMonth
        );

        // 检查匹配结果
        if (matchedRecords.isEmpty()) {
          unmatchedFiles.add(fileName);
        } else if (matchedRecords.size() > 1) {
          multiMatchFiles.add(fileName);
        } else {
          String indexNumber = fileName.substring(0, fileName.indexOf("、"));
          matchResults.add(new MatchResult(matchedRecords.get(0), indexNumber, fileName));
        }
      }

      // 4. 输出异常信息
      printAbnormalCases(unmatchedFiles, multiMatchFiles);

      // 5. 生成结果Excel（带辅助排查列）
      generateResultExcelWithDebugColumns(matchResults, unmatchedFiles, multiMatchFiles, outputFilePath);

      System.out.println("处理完成，结果已保存至: " + outputFilePath);

    } catch (Exception e) {
      System.out.println("处理失败: " + e.getMessage());
      e.printStackTrace();
    }
  }

  // ================ 以下是工具方法 ================ //

  // 扫描文件并按1、2、3排序
  private static List<File> scanAndSortFiles(String folderPath) {
    File dir = new File(folderPath);
    File[] files = dir.listFiles((d, name) -> name.matches("^\\d+、\\d{4}\\.\\d{2}#\\d+\\.pdf$"));
    if (files == null || files.length == 0) {
      System.out.println("错误：未找到符合命名规则的文件（格式：数字、YYYY.MM#凭证号.pdf）");
      return Collections.emptyList();
    }
    Arrays.sort(files, Comparator.comparingInt(f ->
            Integer.parseInt(f.getName().substring(0, f.getName().indexOf("、")))
    ));
    return Arrays.asList(files);
  }

  // 读取数据搜索文件
  private static List<Map<String, String>> readDataSearchFile(String filePath) throws IOException {
    List<Map<String, String>> records = new ArrayList<>();
    try (Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath))) {
      Sheet sheet = workbook.getSheetAt(0);
      for (Row row : sheet) {
        if (row.getRowNum() == 0) continue; // 跳过标题行
        Map<String, String> record = new HashMap<>();
        record.put("日期", getCellValue(row.getCell(0)));      // A列：日期
        record.put("凭证号", getCellValue(row.getCell(1)));    // B列：凭证号
        record.put("摘要", getCellValue(row.getCell(2)));      // C列：摘要
        record.put("科目全名", getCellValue(row.getCell(3)));  // D列：科目全名
        record.put("币别", getCellValue(row.getCell(4)));      // E列：币别
        record.put("借方金额", getCellValue(row.getCell(5)));   // F列：借方金额
        records.add(record);
      }
    }
    return records;
  }

  // 严格匹配：凭证号+年月必须唯一
  private static List<Map<String, String>> findMatchedRecords(
          List<Map<String, String>> records, String voucherNumber, String fileYearMonth
  ) throws ParseException {
    List<Map<String, String>> matched = new ArrayList<>();
    Date fileDate = FILE_DATE_FORMAT.parse(fileYearMonth);
    Calendar fileCalendar = Calendar.getInstance();
    fileCalendar.setTime(fileDate);

    for (Map<String, String> record : records) {
      if (!voucherNumber.equals(record.get("凭证号"))) continue;

      Date recordDate = RECORD_DATE_FORMAT.parse(record.get("日期"));
      Calendar recordCalendar = Calendar.getInstance();
      recordCalendar.setTime(recordDate);

      // 检查年月是否一致
      if (fileCalendar.get(Calendar.YEAR) == recordCalendar.get(Calendar.YEAR) &&
              fileCalendar.get(Calendar.MONTH) == recordCalendar.get(Calendar.MONTH)) {
        matched.add(record);
      }
    }
    return matched;
  }

  // 输出异常信息
  private static void printAbnormalCases(List<String> unmatchedFiles, List<String> multiMatchFiles) {
    if (!unmatchedFiles.isEmpty()) {
      System.out.println("\n=== 以下文件未匹配到数据 ===");
      unmatchedFiles.forEach(System.out::println);
    }
    if (!multiMatchFiles.isEmpty()) {
      System.out.println("\n=== 以下文件匹配到多条数据（需检查） ===");
      multiMatchFiles.forEach(System.out::println);
    }
  }

  // 生成结果Excel（带辅助排查列）
  private static void generateResultExcelWithDebugColumns(
          List<MatchResult> matchResults,
          List<String> unmatchedFiles,
          List<String> multiMatchFiles,
          String outputPath
  ) throws IOException {
    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet("审计结果");

    // 标题行（新增两列）
    String[] headers = {
            "日期", "凭证编号", "业务内容", "科目名称", "二级科目",
            "借方金额", "贷方金额", "附件", "索引号", "审计结论"
    };
    Row headerRow = sheet.createRow(0);
    for (int i = 0; i < headers.length; i++) {
      headerRow.createCell(i).setCellValue(headers[i]);
    }

    // 数据行（正常匹配的记录）
    int rowNum = 1;
    for (MatchResult result : matchResults) {
      Row row = sheet.createRow(rowNum++);
      Map<String, String> record = result.getRecord();

      // 原始数据列
      row.createCell(0).setCellValue(record.get("日期"));
      row.createCell(1).setCellValue(record.get("凭证号"));
      row.createCell(2).setCellValue(record.get("摘要"));
      row.createCell(3).setCellValue(record.get("科目全名"));
      row.createCell(4).setCellValue(record.get("币别"));

      // 借方金额（千分位格式）
      try {
        double amount = Double.parseDouble(record.get("借方金额"));
        Cell amountCell = row.createCell(5);
        amountCell.setCellValue(amount);
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
        amountCell.setCellStyle(style);
      } catch (NumberFormatException e) {
        row.createCell(5).setCellValue(record.get("借方金额"));
      }

      row.createCell(6).setCellValue(""); // 贷方金额
      row.createCell(7).setCellValue(""); // 附件
      row.createCell(8).setCellValue("F2202-50-" + result.getIndexNumber());
      row.createCell(9).setCellValue("无异常");

      // 新增的辅助排查列
      //row.createCell(10).setCellValue(record.get("凭证号")); // 匹配凭证号
      //row.createCell(11).setCellValue(record.get("日期"));   // 匹配日期
    }

    // 追加未匹配的文件（标记为异常）
    for (String fileName : unmatchedFiles) {
      Row row = sheet.createRow(rowNum++);
      row.createCell(8).setCellValue("F2202-50-" + fileName.substring(0, fileName.indexOf("、")));
      row.createCell(9).setCellValue("异常：未匹配到数据");
      // 辅助列留空
      row.createCell(10).setCellValue("");
      row.createCell(11).setCellValue("");
    }

    // 追加匹配到多条的文件（标记为异常）
    for (String fileName : multiMatchFiles) {
      Row row = sheet.createRow(rowNum++);
      row.createCell(8).setCellValue("F2202-50-" + fileName.substring(0, fileName.indexOf("、")));
      row.createCell(9).setCellValue("异常：匹配到多条数据");
      // 辅助列留空
      row.createCell(10).setCellValue("");
      row.createCell(11).setCellValue("");
    }

    // 调整列宽
    for (int i = 0; i < headers.length; i++) {
      sheet.autoSizeColumn(i);
    }

    // 保存文件
    try (FileOutputStream fos = new FileOutputStream(outputPath)) {
      workbook.write(fos);
    }
    workbook.close();
  }

  // 辅助方法：获取单元格值
  private static String getCellValue(Cell cell) {
    if (cell == null) return "";
    switch (cell.getCellType()) {
      case STRING: return cell.getStringCellValue();
      case NUMERIC:
        return DateUtil.isCellDateFormatted(cell)
                ? RECORD_DATE_FORMAT.format(cell.getDateCellValue())
                : String.valueOf(cell.getNumericCellValue());
      default: return "";
    }
  }

  // 匹配结果封装类（新增fileName字段）
  static class MatchResult {
    private final Map<String, String> record;
    private final String indexNumber;
    private final String fileName;

    public MatchResult(Map<String, String> record, String indexNumber, String fileName) {
      this.record = record;
      this.indexNumber = indexNumber;
      this.fileName = fileName;
    }

    public Map<String, String> getRecord() {
      return record;
    }

    public String getIndexNumber() {
      return indexNumber;
    }

    public String getFileName() {
      return fileName;
    }
  }
}