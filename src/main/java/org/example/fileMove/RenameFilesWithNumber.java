package org.example.fileMove;

import java.io.File;
import java.util.Arrays;

/**
 * 重命名目录下的所有文件，添加前缀为数字和顿号
 * 例如：1、文件名.txt  2、文件名.pdf
 * 注意：此代码不会处理子目录中的文件
 */
public class RenameFilesWithNumber {
  public static void main(String[] args) {
    // 配置参数


    // 指定要处理的目录路径
    String directoryPath = "C:\\Users\\20847\\Desktop\\c\\应付账款抽凭"; // <-- 修改为你的目录路径

    File dir = new File(directoryPath);

    // 检查目录是否存在且是目录
    if (!dir.exists() || !dir.isDirectory()) {
      System.out.println("错误：目录不存在或路径不是目录！");
      return;
    }

    // 获取目录下所有文件（不包括子目录）
    File[] files = dir.listFiles(File::isFile);

    if (files == null || files.length == 0) {
      System.out.println("该目录为空或无法读取文件。");
      return;
    }

    // 按文件名排序（可选：按修改时间排序可改为 Arrays.sort(files, (a,b) -> Long.compare(a.lastModified(), b.lastModified()));）
    Arrays.sort(files);

    int counter = 1;
    for (File file : files) {
      String originalName = file.getName();
      String newName = counter + "、" + originalName;
      File newFile = new File(dir, newName);

      // 检查新文件名是否已存在，避免覆盖
      if (newFile.exists()) {
        System.out.println("跳过: " + newName + " 已存在。");
      } else if (file.renameTo(newFile)) {
        System.out.println("重命名: " + originalName + " -> " + newName);
      } else {
        System.out.println("无法重命名: " + originalName);
      }

      counter++;
    }

    System.out.println("重命名完成，共处理 " + files.length + " 个文件。");
  }
}