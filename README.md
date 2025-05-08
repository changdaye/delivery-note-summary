# 送货单合并工具

这是一个基于Spring Boot的Web应用程序，用于合并多个Excel格式的送货单文件。

## 功能特点

- 支持多个Excel文件的合并
- 自动保留第一个文件的表头
- 保持单元格格式和样式
- 支持日期、数字、文本等多种数据类型

## 技术栈

- Spring Boot
- Apache POI
- Lombok

## 使用方法

1. 启动应用程序
2. 访问接口：POST `/api/excel/merge`
3. 上传多个Excel文件（使用multipart/form-data格式，参数名为files）
4. 获取合并后的Excel文件

## API文档

### 合并Excel文件

```
POST /api/excel/merge
Content-Type: multipart/form-data

参数：
- files: Excel文件列表（支持多文件）

返回：
- 合并后的Excel文件（application/vnd.openxmlformats-officedocument.spreadsheetml.sheet）
```

## 开发环境要求

- JDK 8或更高版本
- Maven 3.6或更高版本

## 构建和运行

```bash
# 构建项目
mvn clean package

# 运行应用
java -jar target/deliverynotesummary-0.0.1-SNAPSHOT.jar
```