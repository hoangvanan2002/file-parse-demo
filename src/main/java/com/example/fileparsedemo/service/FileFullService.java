package com.example.fileparsedemo.service;

import com.example.fileparsedemo.model.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParserFactory;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.function.BiConsumer;

@Slf4j
@Service
public class FileFullService {
    public ResultResponse readFileXlsx(MultipartFile file) {
        ResultResponse result = new ResultResponse();
        try {
            OPCPackage pkg = OPCPackage.open(file.getInputStream());
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = reader.getSharedStringsTable();
            XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
            SheetHandler handler = new SheetHandler(sst);
            parser.setContentHandler(handler);
            InputStream sheet = reader.getSheetsData().next();
            parser.parse(new InputSource(sheet));
            sheet.close();

            result.setBomDetails(handler.getBomDetails());
            result.setOperations(handler.getOperations());
            result.setTechnologyProcesses(handler.getTechnologyProcesses());
            result.setTechnologyProcessOperations(handler.getTechnologyProcessOperations());
            result.setCompatibilityOperationMachines(handler.getCompatibilityOperationMachines());
            result.setProducts(handler.getProducts());
        } catch (Exception e) {
            log.error("Error reading Excel file", e);
        }
        return result;
    }

    public byte[] writeFileXlsx(ResultResponse result) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.setCompressTempFiles(true);
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.BLACK.getIndex());
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            writeSheet(workbook, "bom_detail", headerStyle,
                    new String[]{"STT", "Bom Level", "Item Code", "Item Type", "Quantity", "Tỷ lệ hao hụt (%)", "Quy trình"},
                    result.getBomDetails(),
                    (row, detail) -> {
                        row.createCell(0).setCellValue("STT");
                        row.createCell(1).setCellValue(detail.getBomLevel().getValue());
                        row.createCell(2).setCellValue(detail.getItemCode().getValue());
                        row.createCell(3).setCellValue(detail.getType().getValue());
                        row.createCell(4).setCellValue(detail.getQuantity().getValue());
                        row.createCell(5).setCellValue(detail.getComponentYield().getValue());
                        row.createCell(6).setCellValue(detail.getTechnologyProcessCode().getValue());
                    });

            writeSheet(workbook, "operation", headerStyle,
                    new String[]{"Mã CĐ", "Tên CĐ", "Nhóm CĐ", "Nhân lực", "CT(s)", "Bộ phận", "Chức năng", "Mã tổ", "Tần suất chuyển đổi LOT", "Tỉ lệ hoàn thành CĐ", "Tỉ lệ vào/ra", "Leadtime"},
                    result.getOperations(),
                    (row, op) -> {
                        row.createCell(0).setCellValue(op.getOperationCode().getValue());
                        row.createCell(1).setCellValue(op.getOperationName().getValue());
                        row.createCell(2).setCellValue(op.getOperationGroup().getValue());
                        row.createCell(3).setCellValue(op.getEmployeeQuantity().getValue());
                        row.createCell(4).setCellValue(op.getCycleTime().getValue());
                        row.createCell(5).setCellValue(op.getDivisionId().getValue());
                        row.createCell(6).setCellValue(op.getEmployeeGroupCode().getValue());
                        row.createCell(7).setCellValue(op.getTransferFrequencyLot().getValue());
                        row.createCell(8).setCellValue(op.getCompletionRate().getValue());
                        row.createCell(9).setCellValue(op.getCompletionRate().getValue());
                        row.createCell(10).setCellValue(op.getInOutRatio().getValue());
                        row.createCell(11).setCellValue(op.getLeadTime().getValue());
                    });

            writeSheet(workbook, "technology_process", headerStyle,
                    new String[]{"Tên quy trình", "Mã quy trình"},
                    result.getTechnologyProcesses(),
                    (row, tp) -> {
                        row.createCell(0).setCellValue(tp.getTechnologyProcessName().getValue());
                        row.createCell(1).setCellValue(tp.getTechnologyProcessCode().getValue());
                    });

            writeSheet(workbook, "technology_process_operation", headerStyle,
                    new String[]{"Mã quy trình công nghệ", "Mã công đoạn", "Thứ tự công đoạn", "Mô tả", "Line"},
                    result.getTechnologyProcessOperations(),
                    (row, tpo) -> {
                        row.createCell(0).setCellValue(tpo.getTechnologyProcessCode().getValue());
                        row.createCell(1).setCellValue(tpo.getOperationCode().getValue());
                        row.createCell(2).setCellValue(tpo.getOperationOrder().getValue());
                        row.createCell(3).setCellValue(tpo.getDescription().getValue());
                        row.createCell(4).setCellValue(tpo.getOperationLine().getValue());
                    });

            writeSheet(workbook, "compatibility_operation_machine", headerStyle,
                    new String[]{"Mã máy", "Độ ưu tiên", "Thời gian di chuyển (phút)"},
                    result.getCompatibilityOperationMachines(),
                    (row, com) -> {
                        row.createCell(0).setCellValue(com.getMachineCode().getValue());
                        row.createCell(1).setCellValue(com.getPriority().getValue());
                        row.createCell(2).setCellValue(com.getAltTransferMinute().getValue());
                    });

            writeSheet(workbook, "products", headerStyle,
                    new String[]{"Mã hàng hóa", "Mã KH", "Tên tiếng Việt", "Tên tiếng Anh", "Dòng SP", "Loại", "Đặc tính GH", "Model", "Đơn vị"},
                    result.getProducts(),
                    (row, product) -> {
                        row.createCell(0).setCellValue(product.getProductCode().getValue());
                        row.createCell(1).setCellValue(product.getCustomerCode().getValue());
                        row.createCell(2).setCellValue(product.getProductName().getValue());
                        row.createCell(3).setCellValue(product.getProductEnName().getValue());
                        row.createCell(4).setCellValue(product.getProductLine().getValue());
                        row.createCell(5).setCellValue(product.getProductType().getValue());
                        row.createCell(6).setCellValue(product.getDeliveryCharacteristicCode().getValue());
                        row.createCell(7).setCellValue(product.getProductModel().getValue());
                        row.createCell(8).setCellValue(product.getUnit().getValue());
                    });

            workbook.write(out);
            workbook.dispose();
            workbook.close();
            return out.toByteArray();

        } catch (Exception e) {
            log.error("Error writing Excel file", e);
            return null;
        }
    }

    private <T> void writeSheet(SXSSFWorkbook workbook, String sheetName, CellStyle headerStyle,
                                String[] headers, List<T> data, BiConsumer<Row, T> writer) {
        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
        // Data rows
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            writer.accept(row, data.get(i));
            for (int j = 0; j < headers.length; j++) {
                Cell cell = row.getCell(j);
                if (cell == null) cell = row.createCell(j);
                cell.setCellStyle(dataStyle);
            }
        }
    }
}
