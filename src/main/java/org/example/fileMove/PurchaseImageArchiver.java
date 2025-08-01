package org.example.fileMove;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * 将文件夹下面的图片名字与 excel的文件中 入库单号匹配，并将文件移动到索引文件夹下
 */
public class PurchaseImageArchiver {

  public static final int TARGET_IDX = 16;
  public static final int FOLDER_NAME_IDX = 28;

  public static void main(String[] args) {
    // 配置参数
    String sourceFolderPath = "C:\\Users\\gu\\Desktop\\e\\采购细节测试-广东高义\\test\\采购入库单截图";
    String excelFilePath = "C:\\Users\\gu\\Desktop\\e\\采购细节测试-广东高义\\test\\广东高义-采购细节测试样本检查记录630-刘丹.xlsx";
    String sheetName = "1-6月样本检查记录";
    String outputBasePath = "C:\\Users\\gu\\Desktop\\e\\采购细节测试-广东高义\\test\\采购入库单截图";

    // 记录处理结果
    List<String> processedFiles = new ArrayList<>();
    List<String> unmatchedFiles = new ArrayList<>();
    List<String> errorFiles = new ArrayList<>();

    // 1. 读取Excel文件并构建映射关系
    Workbook workbook = null;
    try {
      FileInputStream fis = new FileInputStream(excelFilePath);
      workbook = new XSSFWorkbook(fis);

      Map<String, String> receiptToFolderMap = readExcelMapping(workbook, sheetName);
      if (receiptToFolderMap == null) {
        System.err.println("Excel文件读取失败，程序终止");
        return;
      }

      // 预创建绿色样式
      CellStyle greenStyle = createGreenStyle(workbook);

      // 2. 处理源文件夹中的所有文件
      processSourceFiles(sourceFolderPath, outputBasePath, receiptToFolderMap,
              processedFiles, unmatchedFiles, errorFiles, workbook, sheetName, greenStyle);

      // 3. 保存修改后的Excel文件
      saveModifiedExcel(workbook, excelFilePath);

      // 4. 输出处理结果
      printResults(processedFiles, unmatchedFiles, errorFiles);

    } catch (Exception e) {
      System.err.println("处理过程中发生错误: " + e.getMessage());
      e.printStackTrace();
    } finally {
      if (workbook != null) {
        try {
          workbook.close();
        } catch (IOException e) {
          System.err.println("关闭工作簿时出错: " + e.getMessage());
        }
      }
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

      Cell receiptCell = row.getCell(TARGET_IDX); // Q列
      Cell folderCell = row.getCell(FOLDER_NAME_IDX); // AC列

      if (receiptCell != null && folderCell != null) {
        String receiptNumber = getCellValueAsString(receiptCell).trim();
        String folderName = getCellValueAsString(folderCell).trim();
        if (!receiptNumber.isEmpty() && !folderName.isEmpty()) {
          mapping.put(receiptNumber, folderName);
        }
      }
    }
    return mapping;
  }

  private static CellStyle createGreenStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }

  private static void processSourceFiles(String sourcePath, String outputBasePath,
                                         Map<String, String> mapping,
                                         List<String> processedFiles,
                                         List<String> unmatchedFiles,
                                         List<String> errorFiles,
                                         Workbook workbook, String sheetName,
                                         CellStyle greenStyle) {
    File sourceFolder = new File(sourcePath);
    File[] files = sourceFolder.listFiles();

    if (files == null || files.length == 0) {
      System.out.println("源文件夹中没有可处理的文件");
      return;
    }

    // 确保输出基础目录存在
    try {
      Files.createDirectories(Paths.get(outputBasePath));
    } catch (IOException e) {
      System.err.println("创建输出目录失败: " + e.getMessage());
      return;
    }

    Sheet sheet = workbook.getSheet(sheetName);
    for (File file : files) {
      if (file.isDirectory()) {
        continue; // 跳过子目录
      }

      String fileName = file.getName();
      String baseName = fileName.contains(".")
              ? fileName.substring(0, fileName.lastIndexOf('.'))
              : fileName;

      // 尝试匹配的优先级：
      // 1. 完整文件名（如CGRK-250415012167-1）
      // 2. 去除数字后缀的基础名（如CGRK-250415012167）
      String receiptNumber = findReceiptNumber(mapping, baseName);

      if (receiptNumber != null) {
        String targetFolderName = mapping.get(receiptNumber);
        Path targetPath = Paths.get(outputBasePath, targetFolderName);

        try {
          // 创建目标文件夹
          Files.createDirectories(targetPath);

          // 移动文件到目标位置
          Path destination = targetPath.resolve(fileName);
          Files.move(file.toPath(), destination, StandardCopyOption.REPLACE_EXISTING);

          processedFiles.add(fileName + " -> " + targetFolderName);

          // 标记Excel中的匹配行
          markMatchedRow(sheet, receiptNumber, greenStyle);
        } catch (IOException e) {
          errorFiles.add(fileName + " (移动失败: " + e.getMessage() + ")");
        }
      } else {
        unmatchedFiles.add(fileName);
      }
    }
  }


  /**
   * 从映射表中查找匹配的入库单号
   * 支持多种后缀格式处理：
   * 1. 完整名称直接匹配
   * 2. 短横线数字后缀：XXX-1 → XXX
   * 3. 括号数字后缀：XXX(1) → XXX
   * 4. 空格+括号数字后缀：XXX (1) → XXX
   * 5. 组合情况：XXX-1(2) → XXX-1 → XXX
   */
  private static String findReceiptNumber(Map<String, String> mapping, String baseName) {
    // 1. 首先尝试完整匹配
    if (mapping.containsKey(baseName)) {
      return baseName;
    }

    // 2. 尝试各种后缀去除方式
    String[] testNames = {
            baseName,
            // 处理"XXX-1"格式
            baseName.replaceFirst("-\\d+$", ""),
            // 处理"XXX(1)"格式
            baseName.replaceFirst("\\(\\d+\\)$", ""),
            // 处理"XXX (1)"格式（带空格）
            baseName.replaceFirst("\\s*\\(\\d+\\)$", ""),
            // 组合处理：先去掉括号后缀，再去掉短横线后缀
            baseName.replaceFirst("\\(\\d+\\)$", "").replaceFirst("-\\d+$", ""),
            baseName.replaceFirst("\\s*\\(\\d+\\)$", "").replaceFirst("-\\d+$", "")
    };

    // 去重并保留原始顺序
    Set<String> uniqueNames = new LinkedHashSet<>(Arrays.asList(testNames));

    // 按优先级检查各个可能的名称
    for (String testName : uniqueNames) {
      if (!testName.equals(baseName) && mapping.containsKey(testName)) {
        return testName;
      }
    }

    return null; // 未找到匹配
  }

  private static void markMatchedRow(Sheet sheet, String receiptNumber, CellStyle style) {
    for (Row row : sheet) {
      if (row == null) continue;

      Cell cell = row.getCell(16); // Q列
      if (cell != null && receiptNumber.equals(getCellValueAsString(cell).trim())) {
        // 应用绿色样式到整行
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
    } catch (Exception e) {
      System.out.println(e.getMessage());
    }

    // 重命名临时文件为最终文件
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

  private static void printResults(List<String> processed, List<String> unmatched, List<String> errors) {
    System.out.println("\n========== 处理结果 ==========");
    System.out.println("成功处理的文件 (" + processed.size() + "):");
    processed.forEach(System.out::println);

    System.out.println("\n未匹配的文件 (" + unmatched.size() + "):");
    unmatched.forEach(System.out::println);

    if (!errors.isEmpty()) {
      System.out.println("\n处理失败的文件 (" + errors.size() + "):");
      errors.forEach(System.out::println);
    }
  }
}