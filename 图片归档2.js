const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

/**
 * 复制文件
 * @param {string} source - 源文件路径
 * @param {string} destination - 目标文件路径
 */
function copyFile(source, destination) {
  try {
    // 确保目标目录存在
    const destDir = path.dirname(destination);
    if (!fs.existsSync(destDir)) {
      fs.mkdirSync(destDir, { recursive: true });
    }
    // 复制文件
    fs.copyFileSync(source, destination);
    return true;
  } catch (error) {
    console.error(`复制文件失败: ${error.message}`);
    return false;
  }
}

/**
 * 递归搜索文件
 * @param {string} dir - 搜索目录
 * @param {string} pattern - 文件名模式
 * @returns {string[]} 找到的文件路径数组
 */
function searchFiles(dir, pattern) {
  const results = [];
  
  // 检查目录是否存在
  if (!fs.existsSync(dir)) {
    return results;
  }
  
  try {
    const files = fs.readdirSync(dir, { withFileTypes: true });
    
    for (const file of files) {
      const fullPath = path.join(dir, file.name);
      if (file.isDirectory()) {
        results.push(...searchFiles(fullPath, pattern));
      } else if (file.isFile()) {
        // 简单的模式匹配
        const regex = new RegExp(pattern.replace(/\*/g, '.*'));
        if (regex.test(file.name)) {
          results.push(fullPath);
        }
      }
    }
  } catch (error) {
    console.error(`搜索文件失败: ${error.message}`);
  }
  
  return results;
}

/**
 * 复制身份证图片
 * @param {string} tablePath - Excel表格路径
 * @param {string} searchDir - 搜索目录
 * @param {string} outputDir - 输出目录
 * @param {string} nameSuffix - 文件名后缀
 */
async function copyIdImages(tablePath, searchDir, outputDir, nameSuffix) {
  try {
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`创建目标文件夹: ${outputDir}`);
    }

    // 检查Excel文件是否存在
    if (!fs.existsSync(tablePath)) {
      console.error(`Excel文件不存在: ${tablePath}`);
      return;
    }

    // 读取Excel文件
    const workbook = xlsx.readFile(tablePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);

    // 检查是否有从事专业列
    const firstRow = data[0];
    const hasProfession = firstRow && '从事专业' in firstRow;
    console.log(`检测到${hasProfession ? '从事专业' : '无从事专业'}列`);

    // 遍历表格中的每一行
    let foundCount = 0;
    const notFoundList = [];
    console.log("开始搜索并复制文件...");

    for (const row of data) {
      // 获取身份证号码（第一列）
      const idNumber = String(Object.values(row)[0]).trim();
      
      // 获取从事专业
      let profession = '';
      if (hasProfession) {
        profession = String(row['从事专业'] || '').trim();
      }

      // 构建目标文件夹路径
      let targetFolder = outputDir;
      if (profession) {
        targetFolder = path.join(outputDir, profession);
        if (!fs.existsSync(targetFolder)) {
          fs.mkdirSync(targetFolder, { recursive: true });
          console.log(`创建专业文件夹: ${targetFolder}`);
        }
      }

      // 搜索并复制文件
      const targetPattern = `${idNumber}${nameSuffix}.*`;
      const foundFiles = searchFiles(searchDir, targetPattern);

      if (foundFiles.length > 0) {
        for (const file of foundFiles) {
          const fileName = path.basename(file);
          const destPath = path.join(targetFolder, fileName);
          if (copyFile(file, destPath)) {
            console.log(`已找到并复制: ${fileName} 到 ${profession ? '专业文件夹' : '根文件夹'}`);
            foundCount++;
          }
        }
      } else {
        notFoundList.push(idNumber);
      }
    }

    console.log("\n" + "=".repeat(30));
    console.log("搜索完成！");
    console.log(`成功复制图片: ${foundCount} 张`);
    if (notFoundList.length > 0) {
      console.log(`未找到相关图片的身份证号 (${notFoundList.length} 个):`);
      for (const item of notFoundList) {
        console.log(` - ${item}`);
      }
    }
    console.log("=".repeat(30));
  } catch (error) {
    console.error(`处理过程中出错: ${error.message}`);
  }
}

// 主函数
function main() {
  const MY_TABLE = "身份证号.xlsx";
  
  // 读取配置文件
  let config = {
    SEARCH_ROOT: "F:\\学习\\python\\暂住证批量处理\\output\\45",
    DEST_FOLDER: "未勾选",
    NAME_SUFFIX: "_居住证"
  };
  
  try {
    if (fs.existsSync('config.json')) {
      const configData = fs.readFileSync('config.json', 'utf8');
      const parsedConfig = JSON.parse(configData);
      config = {
        SEARCH_ROOT: parsedConfig.SEARCH_ROOT || config.SEARCH_ROOT,
        DEST_FOLDER: parsedConfig.DEST_FOLDER || config.DEST_FOLDER,
        NAME_SUFFIX: parsedConfig.NAME_SUFFIX || config.NAME_SUFFIX
      };
    }
  } catch (error) {
    console.error(`读取配置文件失败: ${error.message}`);
  }
  
  const { SEARCH_ROOT, DEST_FOLDER, NAME_SUFFIX } = config;
  
  // 检查搜索目录是否存在
  if (!fs.existsSync(SEARCH_ROOT)) {
    console.error(`搜索目录不存在: ${SEARCH_ROOT}`);
    console.log("请检查config.json中的SEARCH_ROOT配置是否正确");
    return;
  }
  
  copyIdImages(MY_TABLE, SEARCH_ROOT, DEST_FOLDER, NAME_SUFFIX);
}

// 运行主函数
if (require.main === module) {
  main();
}

// 导出函数供其他模块使用
module.exports = { copyIdImages };
