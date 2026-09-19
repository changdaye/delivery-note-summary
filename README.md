# Excel 送货单多文件智能合并服务 (Delivery Note Summary)

[![Java 8+](https://img.shields.io/badge/Java-8%2B-orange.svg)](https://www.oracle.com/java/)
[![Spring Boot 2.7](https://img.shields.io/badge/Spring%20Boot-2.7.17-brightgreen.svg)](https://spring.io/projects/spring-boot)
[![Apache POI 5.2](https://img.shields.io/badge/Apache%20POI-5.2.3-blue.svg)](https://poi.apache.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

基于 **Spring Boot** 与 **Apache POI** 的 Excel 送货单与多源表格批量智能合并 Web 服务。

专门用于解决日常业务中大量离散 Excel 送货单、对账单格式繁杂、手工复制粘贴效率低下且易错的问题。支持同时上传多个 Excel 表格，服务会自动复用首表表头、无缝合并所有数据行，并严格保留原始单元格格式、日期与公式。

---

## ✨ 核心特性

- **多文件批量合并**：支持一次性上传数十份甚至上百份 Excel 送货单文件（`.xlsx` / `.xls` 双格式兼容）。
- **智能表头继承**：自动提取第一个文件的首行作为全局标准表头，后续文件仅提取有效数据行，避免表头重复堆叠。
- **单元格数据类型保真**：基于 Apache POI 深度解析，完美保留 **日期（Date）**、**浮点/整数（Numeric）**、**公式（Formula）**、**文本（String）** 及 **布尔值（Boolean）**。
- **样式与排版克隆**：保留原单元格对齐方式、边框与基础样式（通过 `cloneStyleFrom`）。
- **纯流式极速响应**：基于内存字节流（ByteArrayOutputStream）处理并即时以附件形式触发浏览器下载，不产生服务器磁盘垃圾。

---

## 🛠️ 技术栈

- **后端核心框架**：Spring Boot 2.7.17 (Web)
- **文档处理引擎**：Apache POI 5.2.3 / POI-OOXML
- **开发辅助工具**：Lombok / Maven 3.6+
- **运行环境要求**：JDK 1.8 或更高版本

---

## 🚀 接口规范 (API Reference)

### 批量合并 Excel 送货单

```http
POST /api/excel/merge
Content-Type: multipart/form-data
```

#### 请求参数

| 参数名 | 类型 | 必填 | 说明 |
| :--- | :--- | :--- | :--- |
| `files` | `List<MultipartFile>` | 是 | 待合并的 Excel 文件列表（支持多选上传） |

#### 响应结果

- **HTTP 状态码**：`200 OK`
- **Content-Type**：`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- **Content-Disposition**：`attachment; filename="merged_excel.xlsx"`
- **返回体**：已合并完成的标准 `.xlsx` 二进制文件流。

#### cURL 调用示例

```bash
curl -X POST http://localhost:8080/api/excel/merge \
  -F "files=@/path/to/delivery_note_01.xlsx" \
  -F "files=@/path/to/delivery_note_02.xlsx" \
  -F "files=@/path/to/delivery_note_03.xlsx" \
  --output merged_result.xlsx
```

---

## 📦 构建与部署

```bash
# 1. 克隆仓库
git clone https://github.com/changdaye/delivery-note-summary.git
cd delivery-note-summary

# 2. Maven 编译与打包
mvn clean package -DskipTests

# 3. 运行应用
java -jar target/delivery-note-summary-0.0.1-SNAPSHOT.jar
```

服务默认运行在 `http://localhost:8080`。

