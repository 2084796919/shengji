package org.example.fileMove;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * 递归将文件夹下面的图片名字与 Excel 的文件中  Z列凭证号号匹配，并将文件移动到索引文件夹下
 * 处理结果会记录在 Excel 中，并生成一个问题文件 Sheet
 */
public class PurchaseImageArchiver1 {

  public static final int RECEIPT_NUMBER_IDX = 25; // Z列(0-based)
  public static final int FOLDER_NAME_IDX = 28;    // AC列(0-based)
  public static final String PROBLEM_SHEET_NAME = "问题文件";

  public static void main(String[] args) {
    // 配置参数
    String sourceFolderPath = "C:\\Users\\gu\\Desktop\\f\\凭证汇总";
    String excelFilePath = "C:\\Users\\gu\\Desktop\\f\\凭证汇总\\广东高义-采购细节测试样本检查记录630-刘丹.xlsx";
    String sheetName = "1-6月样本检查记录";

    // 记录处理结果
    Map<String, List<String>> processedFiles = new LinkedHashMap<>();
    List<String> unmatchedFiles = new ArrayList<>();
    List<String> errorFiles = new ArrayList<>();

    try (Workbook workbook = new XSSFWorkbook(Files.newInputStream(Paths.get(excelFilePath)))) {
      // 1. 读取Excel映射关系
      Map<String, String> receiptToFolderMap = readExcelMapping(workbook, sheetName);
      if (receiptToFolderMap == null) {
        System.err.println("Excel映射关系读取失败，程序终止");
        return;
      }

      // 2. 预创建绿色样式
      CellStyle greenStyle = createGreenStyle(workbook);

      // 3. 处理每个子目录
      File sourceFolder = new File(sourceFolderPath);
      File[] subDirs = sourceFolder.listFiles(File::isDirectory);

      if (subDirs == null || subDirs.length == 0) {
        System.out.println("源文件夹中没有子目录");
        return;
      }

      for (File subDir : subDirs) {
        System.out.println("\n正在处理目录: " + subDir.getName());
        processSubDirectory(subDir, receiptToFolderMap,
                processedFiles, unmatchedFiles, errorFiles,
                workbook, sheetName, greenStyle);
      }
      // 4. 创建问题文件Sheet并写入数据
      createProblemFilesSheet(workbook, unmatchedFiles, errorFiles);

      // 4. 保存修改后的Excel
      saveModifiedExcel(workbook, excelFilePath);

      // 5. 输出结果
      printResults(processedFiles, unmatchedFiles, errorFiles);

    } catch (Exception e) {
      System.err.println("处理失败: " + e.getMessage());
      e.printStackTrace();
    }
  }

  private static void createProblemFilesSheet(Workbook workbook,
                                              List<String> unmatchedFiles,
                                              List<String> errorFiles) {
    // 删除已存在的problem sheet（如果存在）
    int problemSheetIndex = workbook.getSheetIndex(PROBLEM_SHEET_NAME);
    if (problemSheetIndex != -1) {
      workbook.removeSheetAt(problemSheetIndex);
    }

    // 创建新的problem sheet
    Sheet problemSheet = workbook.createSheet(PROBLEM_SHEET_NAME);

    // 创建表头
    Row headerRow = problemSheet.createRow(0);
    headerRow.createCell(0).setCellValue("文件路径");
    headerRow.createCell(1).setCellValue("问题类型");
    headerRow.createCell(2).setCellValue("错误信息");

    // 写入未匹配文件
    int rowNum = 1;
    for (String filePath : unmatchedFiles) {
      Row row = problemSheet.createRow(rowNum++);
      row.createCell(0).setCellValue(filePath);
      row.createCell(1).setCellValue("未匹配");
      row.createCell(2).setCellValue("未找到对应的入库单号");
    }

    // 写入错误文件
    for (String error : errorFiles) {
      Row row = problemSheet.createRow(rowNum++);

      // 从错误信息中提取文件路径和错误详情
      String[] parts = error.split(" → ", 2);
      String filePath = parts[0];
      String errorMsg = parts.length > 1 ? parts[1] : "";

      row.createCell(0).setCellValue(filePath);
      row.createCell(1).setCellValue("处理失败");
      row.createCell(2).setCellValue(errorMsg);
    }

    // 自动调整列宽
    for (int i = 0; i < 3; i++) {
      problemSheet.autoSizeColumn(i);
    }
  }

  private static Map<String, String> readExcelMapping(Workbook workbook, String sheetName) {
    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      System.err.println("错误: 未找到工作表 '" + sheetName + "'");
      return null;
    }

    Map<String, String> mapping = new HashMap<>();
    for (Row row : sheet) {
      if (row == null) continue;

      Cell receiptCell = row.getCell(RECEIPT_NUMBER_IDX); // Z列
      Cell folderCell = row.getCell(FOLDER_NAME_IDX);     // AC列

      if (receiptCell != null && folderCell != null) {
        String receiptNumber = getCellValueAsString(receiptCell).trim();
        String folderName = getCellValueAsString(folderCell).trim();

        if (!receiptNumber.isEmpty() && !folderName.isEmpty()) {
          // 1. 首先添加原始完整的键值对
          mapping.put(receiptNumber, folderName);

          // 2. 检查是否包含"、"分隔符
          if (receiptNumber.contains("、")) {
            // 拆分复合入库单号
            String[] subNumbers = receiptNumber.split("、");
            for (String subNum : subNumbers) {
              String trimmedSubNum = subNum.trim();
              if (!trimmedSubNum.isEmpty()) {
                // 为每个子编号添加映射（指向同一个文件夹）
                mapping.put(trimmedSubNum, folderName);
              }
            }
          }
        }
      }
    }
    return mapping;
  }

  private static void processSubDirectory(File subDir,
                                          Map<String, String> mapping,
                                          Map<String, List<String>> processedFiles,
                                          List<String> unmatchedFiles,
                                          List<String> errorFiles,
                                          Workbook workbook, String sheetName,
                                          CellStyle greenStyle) {
    File[] files = subDir.listFiles(File::isFile);
    if (files == null || files.length == 0) {
      System.out.println("  目录中没有可处理的文件");
      return;
    }

    List<String> processedInThisDir = new ArrayList<>();
    Sheet sheet = workbook.getSheet(sheetName);

    for (File file : files) {
      String fileName = file.getName();
      String baseName = getBaseName(fileName);

      String receiptNumber = findReceiptNumber(mapping, baseName);

      if (receiptNumber != null) {
        String targetFolderName = mapping.get(receiptNumber);
        Path targetPath = Paths.get(subDir.getAbsolutePath(), targetFolderName);

        try {
          Files.createDirectories(targetPath);
          Path destination = targetPath.resolve(fileName);
          Files.move(file.toPath(), destination, StandardCopyOption.REPLACE_EXISTING);

          processedInThisDir.add(fileName + " → " + targetFolderName);
          markMatchedRow(sheet, receiptNumber, greenStyle);
        } catch (IOException e) {
          String errorMsg = String.format("%s/%s → 移动失败: %s",
                  subDir.getName(), fileName, e.getMessage());
          errorFiles.add(errorMsg);
          System.err.println("  " + errorMsg);
        }
      } else {
        String unmatchedMsg = subDir.getName() + "/" + fileName;
        unmatchedFiles.add(unmatchedMsg);
        System.out.println("  未匹配: " + unmatchedMsg);
      }
    }

    if (!processedInThisDir.isEmpty()) {
      processedFiles.put(subDir.getName(), processedInThisDir);
    }
  }

  private static String findReceiptNumber(Map<String, String> mapping, String baseName) {
    // 1. 首先尝试完整匹配
    if (mapping.containsKey(baseName)) {
      return baseName;
    }

    // 2. 尝试各种后缀去除方式
    String[] testNames = {
            baseName,
            baseName.replaceFirst("-\\d+$", ""),      // 处理"XXX-1"格式
            baseName.replaceFirst("\\(\\d+\\)$", ""),  // 处理"XXX(1)"格式
            baseName.replaceFirst("\\s*\\(\\d+\\)$", ""), // 处理"XXX (1)"格式
            baseName.replaceFirst("\\(\\d+\\)$", "").replaceFirst("-\\d+$", ""), // 组合处理
            baseName.replaceFirst("\\s*\\(\\d+\\)$", "").replaceFirst("-\\d+$", "")
    };

    // 去重并保留原始顺序
    Set<String> uniqueNames = new LinkedHashSet<>(Arrays.asList(testNames));

    for (String testName : uniqueNames) {
      if (!testName.equals(baseName) && mapping.containsKey(testName)) {
        return testName;
      }
    }

    return null;
  }

  private static String getBaseName(String fileName) {
    int lastDot = fileName.lastIndexOf('.');
    return lastDot > 0 ? fileName.substring(0, lastDot) : fileName;
  }

  private static CellStyle createGreenStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }

  private static void markMatchedRow(Sheet sheet, String receiptNumber, CellStyle style) {
    for (Row row : sheet) {
      if (row == null) continue;

      Cell cell = row.getCell(RECEIPT_NUMBER_IDX);
      if (cell != null && receiptNumber.equals(getCellValueAsString(cell).trim())) {
        // 应用样式到整行
        for (int i = 0; i < row.getLastCellNum(); i++) {
          Cell currentCell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
          currentCell.setCellStyle(style);
        }
      }
    }
  }

  private static void saveModifiedExcel(Workbook workbook, String originalFilePath) throws IOException {
    String outputExcelPath = originalFilePath.replace(".xlsx", "_processed.xlsx");
    File tempFile = new File(outputExcelPath + ".temp");

    try (FileOutputStream out = new FileOutputStream(tempFile)) {
      workbook.write(out);
    }

    Files.move(tempFile.toPath(), Paths.get(outputExcelPath), StandardCopyOption.REPLACE_EXISTING);
    System.out.println("\n修改后的Excel已保存为: " + outputExcelPath);
  }

  private static String getCellValueAsString(Cell cell) {
    if (cell == null) return "";

    switch (cell.getCellType()) {
      case STRING:
        return cell.getStringCellValue();
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

  private static void printResults(Map<String, List<String>> processedFiles,
                                   List<String> unmatchedFiles,
                                   List<String> errorFiles) {
    System.out.println("\n========== 处理结果汇总 ==========");

    // 成功处理的文件
    int totalProcessed = processedFiles.values().stream().mapToInt(List::size).sum();
    System.out.println("\n成功处理的文件 (" + totalProcessed + "):");
    processedFiles.forEach((folder, files) -> {
      System.out.println("\n[" + folder + "] (" + files.size() + "):");
      files.forEach(file -> System.out.println("  " + file));
    });

    // 未匹配的文件
    if (!unmatchedFiles.isEmpty()) {
      System.out.println("\n未匹配的文件 (" + unmatchedFiles.size() + "):");
      unmatchedFiles.forEach(file -> System.out.println("  " + file));
    }

    // 处理失败的文件
    if (!errorFiles.isEmpty()) {
      System.out.println("\n处理失败的文件 (" + errorFiles.size() + "):");
      errorFiles.forEach(error -> System.out.println("  " + error));
    }

    // 统计信息
    System.out.println("\n========== 统计信息 ==========");
    System.out.println("总成功处理文件: " + totalProcessed);
    System.out.println("总未匹配文件: " + unmatchedFiles.size());
    System.out.println("总处理失败文件: " + errorFiles.size());
  }
}