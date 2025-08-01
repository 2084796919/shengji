package org.example.fileMove;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.stream.Collectors;

public class AccountsPayableDocumentScanner {

  public static void main(String[] args) {
    // 配置参数
    String sourceFolderPath = "C:\\Users\\20847\\Desktop\\c\\应付账款抽凭";
    String searchExcelPath = "C:\\Users\\20847\\Desktop\\c\\数据搜索.xlsx";
    String outputExcelName = "应付账款抽凭统计.xlsx";

    try {
      // 1. 扫描文件夹获取文件列表
      List<FileInfo> fileInfos = scanSourceFolder(sourceFolderPath);

      // 2. 读取数据搜索Excel
      Map<String, List<ExcelData>> searchData = readSearchExcel(searchExcelPath);

      // 3. 生成统计报表
      generateReport(fileInfos, searchData, Paths.get(sourceFolderPath, outputExcelName).toString());

      System.out.println("处理完成，报表已生成");
    } catch (Exception e) {
      System.err.println("处理过程中发生错误: " + e.getMessage());
      e.printStackTrace();
    }
  }

  // Excel数据类
  private static class ExcelData {
    String date; // 日期
    String voucherNo; // 凭证号
    String summary; // 摘要
    String subject; // 科目全名
    String subSubject; // 币别
    String debitAmount; // 借方金额

    public ExcelData(String date, String voucherNo, String summary,
                     String subject, String subSubject, String debitAmount) {
      this.date = date;
      this.voucherNo = voucherNo;
      this.summary = summary;
      this.subject = subject;
      this.subSubject = subSubject;
      this.debitAmount = debitAmount;
    }
  }

  private static List<FileInfo> scanSourceFolder(String folderPath) throws IOException {
    File folder = new File(folderPath);
    File[] files = folder.listFiles(File::isFile); // 获取所有文件

    if (files == null || files.length == 0) {
      throw new IOException("文件夹中没有文件");
    }

    List<FileInfo> fileInfos = new ArrayList<>();
    for (File file : files) {
      String name = file.getName();
      try {
        // 尝试解析符合格式的文件名
        if (name.matches("^\\d+、\\d+\\.\\d+#\\d+\\.pdf$")) {
          String[] parts = name.split("[、#.]"); // 拆分文件名
          fileInfos.add(new FileInfo(parts[0], parts[3], name, false));
        } else {
          // 记录不符合格式的文件
          fileInfos.add(new FileInfo("", "", name, true));
        }
      } catch (Exception e) {
        // 记录解析异常的文件
        fileInfos.add(new FileInfo("", "", name, true));
      }
    }

    // 按前缀数字排序（仅对正常文件有效）
    fileInfos.sort((a, b) -> {
      if (a.isAbnormal || b.isAbnormal) return 0;
      return Integer.compare(Integer.parseInt(a.prefix), Integer.parseInt(b.prefix));
    });

    return fileInfos;
  }

  // 修改后的FileInfo类
  private static class FileInfo {
    String prefix; // 文件前缀(1,2,3...)
    String number; // 凭证号(如590)
    String fullName; // 完整文件名
    boolean isAbnormal; // 是否异常文件

    public FileInfo(String prefix, String number, String fullName, boolean isAbnormal) {
      this.prefix = prefix;
      this.number = number;
      this.fullName = fullName;
      this.isAbnormal = isAbnormal;
    }
  }

  // 读取数据搜索Excel
  private static Map<String, List<ExcelData>> readSearchExcel(String excelPath) throws IOException {
    Map<String, List<ExcelData>> result = new HashMap<>();

    try (Workbook workbook = new XSSFWorkbook(Files.newInputStream(Paths.get(excelPath)))) {
      Sheet sheet = workbook.getSheetAt(0); // 第一个sheet

      for (int i = 1; i <= sheet.getLastRowNum(); i++) { // 从第2行开始
        Row row = sheet.getRow(i);
        if (row == null) continue;

        // B列是凭证号(格式应为"记590")
        Cell voucherCell = row.getCell(1);
        if (voucherCell == null) continue;

        String voucherText = getCellValueAsString(voucherCell);
        if (!voucherText.startsWith("记")) continue;

        String voucherNo = voucherText.substring(1); // 提取数字部分

        // 读取其他列数据
        String date = getCellValueAsString(row.getCell(0)); // A列 日期
        String summary = getCellValueAsString(row.getCell(2)); // C列 摘要
        String subject = getCellValueAsString(row.getCell(3)); // D列 科目全名
        String subSubject = getCellValueAsString(row.getCell(4)); // E列 币别
        String debitAmount = getCellValueAsString(row.getCell(5)); // F列 借方金额

        // 添加到结果集
        result.computeIfAbsent(voucherNo, k -> new ArrayList<>())
                .add(new ExcelData(date, voucherNo, summary, subject, subSubject, debitAmount));
      }
    }

    return result;
  }

  // 生成报表
  private static void generateReport(List<FileInfo> fileInfos,
                                     Map<String, List<ExcelData>> searchData,
                                     String outputPath) throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("应付账款抽凭统计");

      // 创建表头
      String[] headers = {"日期", "凭证编号", "业务内容", "科目名称", "二级科目",
              "借方金额", "贷方金额", "附件", "索引号", "审计结论"};
      Row headerRow = sheet.createRow(0);
      for (int i = 0; i < headers.length; i++) {
        headerRow.createCell(i).setCellValue(headers[i]);
      }

      // 填充数据
      int rowNum = 1;
      for (FileInfo fileInfo : fileInfos) {
        if (fileInfo.isAbnormal) {
          // 异常文件记录
          Row row = sheet.createRow(rowNum++);
          row.createCell(0).setCellValue("异常文件");
          row.createCell(1).setCellValue(fileInfo.fullName);
          row.createCell(9).setCellValue("文件命名格式不符合要求");
          continue;
        }

        List<ExcelData> matchedData = searchData.get(fileInfo.number);
        if (matchedData == null || matchedData.isEmpty()) {
          // 没有匹配数据的异常情况
          Row row = sheet.createRow(rowNum++);
          row.createCell(0).setCellValue("未匹配");
          row.createCell(1).setCellValue(fileInfo.fullName);
          row.createCell(8).setCellValue("F2202-50-" + fileInfo.prefix);
          row.createCell(9).setCellValue("异常: 未找到匹配凭证");
          continue;
        }

        // 合并索引号单元格
        if (matchedData.size() > 1) {
          sheet.addMergedRegion(new CellRangeAddress(
                  rowNum, rowNum + matchedData.size() - 1, 8, 8));
          sheet.addMergedRegion(new CellRangeAddress(
                  rowNum, rowNum + matchedData.size() - 1, 9, 9));
        }

        // 写入每行数据
        for (ExcelData data : matchedData) {
          Row row = sheet.createRow(rowNum++);
          row.createCell(0).setCellValue(data.date);
          row.createCell(1).setCellValue(data.voucherNo);
          row.createCell(2).setCellValue(data.summary);
          row.createCell(3).setCellValue(data.subject);
          row.createCell(4).setCellValue(data.subSubject);
          row.createCell(5).setCellValue(data.debitAmount);
          row.createCell(6).setCellValue(""); // 贷方金额为空
          row.createCell(7).setCellValue(""); // 附件为空
          row.createCell(8).setCellValue("F2202-50-" + fileInfo.prefix);
          row.createCell(9).setCellValue("无异常");
        }
      }

      // 自动调整列宽
      for (int i = 0; i < headers.length; i++) {
        sheet.autoSizeColumn(i);
      }

      // 保存文件
      try (FileOutputStream out = new FileOutputStream(outputPath)) {
        workbook.write(out);
      }
    }
  }

  // 获取单元格值
  private static String getCellValueAsString(Cell cell) {
    if (cell == null) return "";

    switch (cell.getCellType()) {
      case STRING:
        return cell.getStringCellValue().trim();
      case NUMERIC:
        return String.valueOf((int) cell.getNumericCellValue());
      case BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      case FORMULA:
        try {
          return cell.getStringCellValue();
        } catch (IllegalStateException e) {
          return String.valueOf(cell.getNumericCellValue());
        }
      default:
        return "";
    }
  }
}