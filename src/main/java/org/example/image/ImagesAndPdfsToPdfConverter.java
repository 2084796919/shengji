package org.example.image;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.util.Matrix;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 图片插入到pdf汇总
 * 图片名称 1、 2、顺序
 */
public class ImagesAndPdfsToPdfConverter {

  public static void main(String[] args) {
    String inputFolderPath = "C:\\Users\\20847\\Desktop\\a\\GDGY-CGKZCS-4-付款凭证"; // 替换为你的文件夹路径
    String outputPdfName = "merged_output.pdf"; // 输出的PDF文件名

    try {
      mergeImagesAndPdfsToPdf(inputFolderPath, outputPdfName);
      System.out.println("PDF 合并成功！");
    } catch (IOException e) {
      System.err.println("处理过程中发生错误: " + e.getMessage());
      e.printStackTrace();
    }
  }

  public static void mergeImagesAndPdfsToPdf(String inputFolderPath, String outputPdfName) throws IOException {
    try (PDDocument finalPdf = new PDDocument()) {
      File folder = new File(inputFolderPath);

      if (!folder.exists() || !folder.isDirectory()) {
        throw new IOException("指定的路径不是有效文件夹: " + inputFolderPath);
      }

      // 获取支持的图片和PDF文件
      List<File> files = Arrays.asList(folder.listFiles((dir, name) ->
              name.toLowerCase().matches(".*\\.(jpg|jpeg|png|gif|bmp|pdf)$")
      ));

      // 按文件名开头的数字排序，数字相同时保持原顺序
      files.sort(Comparator.comparingInt(f -> extractLeadingNumber(f.getName())));

      if (files.isEmpty()) {
        throw new IOException("文件夹中没有找到支持的图片或PDF文件");
      }

      // 处理每个文件
      for (File file : files) {
        String fileName = file.getName().toLowerCase();
        if (fileName.endsWith(".pdf")) {
          // 处理PDF文件（合并所有页）
          try (PDDocument pdfDoc = PDDocument.load(file)) {
            PDFMergerUtility merger = new PDFMergerUtility();
            merger.appendDocument(finalPdf, pdfDoc);
          }
        } else {
          // 处理图片文件（插入到新的一页）
          try {
            PDPage page = new PDPage(PDRectangle.A4);
            finalPdf.addPage(page);

            PDImageXObject pdImage = PDImageXObject.createFromFile(file.getAbsolutePath(), finalPdf);

            float pageWidth = page.getMediaBox().getWidth();
            float pageHeight = page.getMediaBox().getHeight();
            float imageWidth = pdImage.getWidth();
            float imageHeight = pdImage.getHeight();

            float scaleX = pageWidth / imageWidth;
            float scaleY = pageHeight / imageHeight;
            float scale = Math.min(scaleX, scaleY);

            float scaledWidth = imageWidth * scale;
            float scaledHeight = imageHeight * scale;

            float x = (pageWidth - scaledWidth) / 2;
            float y = (pageHeight - scaledHeight) / 2;

            try (PDPageContentStream contentStream = new PDPageContentStream(finalPdf, page)) {
              contentStream.transform(Matrix.getTranslateInstance(x, y));
              contentStream.drawImage(pdImage, 0, 0, scaledWidth, scaledHeight);
            }
          } catch (IOException e) {
            System.err.println("跳过无法处理的文件: " + file.getName() + " - " + e.getMessage());
          }
        }
      }

      // 保存PDF到源文件夹
      String outputPath = new File(folder, outputPdfName).getAbsolutePath();
      finalPdf.save(outputPath);
    }
  }

  // 提取文件名开头的数字（如 "1、WXGYBZ..." → 返回 1）
  private static int extractLeadingNumber(String fileName) {
    Pattern pattern = Pattern.compile("^(\\d+)、"); // 匹配 "数字、" 开头
    Matcher matcher = pattern.matcher(fileName);
    if (matcher.find()) {
      return Integer.parseInt(matcher.group(1));
    }
    throw new RuntimeException("文件异常"+fileName);
  }
}