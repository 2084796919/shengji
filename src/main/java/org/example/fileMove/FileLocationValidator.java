package org.example.fileMove;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 校验文件夹下的文件与excel中的索引是否匹配，通过入库单号关联
 */
public class FileLocationValidator {

  public static void main(String[] args) {
    // 配置参数
    String archiveBasePath = "C:\\Users\\gu\\Desktop\\e\\采购细节测试-广东高义\\test\\采购入库单截图";
    String excelFilePath = "C:\\Users\\gu\\Desktop\\e\\采购细节测试-广东高义\\test\\广东高义-采购细节测试样本检查记录630-刘丹.xlsx";
    String sheetName = "1-6月样本检查记录";

    // 记录校验结果
    List<String> correctFiles = new ArrayList<>();
    List<String> incorrectFiles = new ArrayList<>();
    List<String> unmatchedFiles = new ArrayList<>();

    // 1. 读取Excel文件并构建映射关系
    Map<String, String> receiptToFolderMap = readExcelMapping(excelFilePath, sheetName);
    if (receiptToFolderMap == null) {
      System.err.println("Excel文件读取失败，程序终止");
      return;
    }

    // 2. 校验归档文件夹结构
    validateArchiveStructure(archiveBasePath, receiptToFolderMap, correctFiles, incorrectFiles, unmatchedFiles);

    // 3. 输出校验结果
    printValidationResults(correctFiles, incorrectFiles, unmatchedFiles);
  }

  private static Map<String, String> readExcelMapping(String excelFilePath, String sheetName) {
    try (FileInputStream fis = new FileInputStream(excelFilePath);
         Workbook workbook = new XSSFWorkbook(fis)) {

      Sheet sheet = workbook.getSheet(sheetName);
      if (sheet == null) {
        System.err.println("错误: 未找到工作表 '" + sheetName + "'");
        return null;
      }

      Map<String, String> mapping = new HashMap<>();
      for (Row row : sheet) {
        if (row == null) continue;

        Cell receiptCell = row.getCell(16); // Q列
        Cell folderCell = row.getCell(28); // AC列

        if (receiptCell != null && folderCell != null) {
          String receiptNumber = getCellValueAsString(receiptCell).trim();
          String folderName = getCellValueAsString(folderCell).trim();
          if (!receiptNumber.isEmpty() && !folderName.isEmpty()) {
            mapping.put(receiptNumber, folderName);
          }
        }
      }
      return mapping;
    } catch (Exception e) {
      System.err.println("读取Excel文件时出错: " + e.getMessage());
      return null;
    }
  }

  private static void validateArchiveStructure(String basePath,
                                               Map<String, String> mapping,
                                               List<String> correctFiles,
                                               List<String> incorrectFiles,
                                               List<String> unmatchedFiles) {
    File baseDir = new File(basePath);
    if (!baseDir.exists() || !baseDir.isDirectory()) {
      System.err.println("归档基础目录不存在或不是目录");
      return;
    }

    // 遍历所有子文件夹
    File[] folders = baseDir.listFiles(File::isDirectory);
    if (folders == null) return;

    for (File folder : folders) {
      String folderName = folder.getName();

      // 检查文件夹中的文件
      File[] files = folder.listFiles(File::isFile);
      if (files == null) continue;

      for (File file : files) {
        String fileName = file.getName();
        String baseName = fileName.contains(".")
                ? fileName.substring(0, fileName.lastIndexOf('.'))
                : fileName;

        // 查找匹配的入库单号
        String receiptNumber = findReceiptNumber(mapping, baseName);

        if (receiptNumber != null) {
          // 检查文件夹名称是否正确
          String expectedFolder = mapping.get(receiptNumber);
          if (folderName.equals(expectedFolder)) {
            correctFiles.add(fileName + " -> " + folderName);
          } else {
            incorrectFiles.add(fileName + " (当前位置: " + folderName +
                    ", 应放位置: " + expectedFolder + ")");
          }
        } else {
          unmatchedFiles.add(fileName + " (在文件夹: " + folderName + ")");
        }
      }
    }

    // 检查是否有应该存在但未创建的文件夹
    Set<String> existingFolders = Arrays.stream(folders)
            .map(File::getName)
            .collect(Collectors.toSet());

    Set<String> expectedFolders = new HashSet<>(mapping.values());
    expectedFolders.removeAll(existingFolders);

    if (!expectedFolders.isEmpty()) {
      System.out.println("\n以下文件夹应该存在但未创建:");
      expectedFolders.forEach(System.out::println);
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

  private static void printValidationResults(List<String> correctFiles,
                                             List<String> incorrectFiles,
                                             List<String> unmatchedFiles) {
    System.out.println("\n========== 校验结果 ==========");
    System.out.println("正确归档的文件 (" + correctFiles.size() + "):");
    correctFiles.forEach(System.out::println);

    System.out.println("\n位置不正确的文件 (" + incorrectFiles.size() + "):");
    incorrectFiles.forEach(System.out::println);

    System.out.println("\n未匹配到入库单号的文件 (" + unmatchedFiles.size() + "):");
    unmatchedFiles.forEach(System.out::println);
  }
}