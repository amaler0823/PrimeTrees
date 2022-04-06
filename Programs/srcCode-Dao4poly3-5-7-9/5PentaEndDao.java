package com;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class RotaEnd {
    int row = 600;
//    int row = 100;
    int max=0;
    int tmprelt = 0;

    public RotaEnd() {
    }

    public static void main(String[] args) {
        RotaEnd rotaEnd = new RotaEnd();
        rotaEnd.writeExcel();
    }

    void writeExcel() {
        String fileName = "Y:\\2021job\\@tree\\测试RotaEnd.xlsx";
        ((ExcelWriterSheetBuilder)EasyExcel.write(fileName).sheet("RotaEnd模板").registerWriteHandler(new RotaEndRecordBorrowExcelHandler())).doWrite(this.dataList());
    }

    //    https://segmentfault.com/a/1190000038566393
    private List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList();
//        int ma = 1;
        int ma = 0;
        int a = 0;

        for(int i = 1; i < this.row; ++i) {
            StringBuilder stringBuilder = new StringBuilder();

            max = (3*i*i-i)/2;
            for(int j = ma+1; j <= this.max; ++j) {
                a = Integer.valueOf(j);
                String stmp=a+"";
                if (this.numIsP(Integer.valueOf(j))) {
                    stmp = "P" + stmp;
                }

                //for RZhong
                if (this.numIsInZ(a)) {
                    stmp = "Z" + stmp;
                }

                stringBuilder.append(stmp + ",");
            }
            ma = a;//            https://blog.csdn.net/qq_38974638/article/details/117395831
            list.add(new ArrayList(Arrays.asList(stringBuilder.toString().split(","))));
            stringBuilder.delete(0, stringBuilder.toString().length());
        }
        return list;
    }

    Boolean numIsInZ(int num) {//6,52,93,248,377,783,1927,2502,4031,5018
        //0,4,6,11,14,21,25,34,39,50
        int z = 0;
        int a = 4;
        int b = 2;
        int flag = 0;

        for (int i = 0; i < this.row; i++) {
            if (reltIsInt(z, num) && 0!=tmprelt) {
                if((z+1)<=tmprelt){
                    return true;
                }
            }
            tmprelt=0;

            if(0==flag){
                z = z + a;
                if(5>a){
                    a++;
                }else {
                    a=a+2;
                }

                flag=1;
            }else{
                z = z + b;
                b++;
                flag=0;
            }

        }
        return false;
    }

    Boolean reltIsInt(int z, int num) {
        //3nn-5n-z*2-num*2=0
        //3nn-5n-0*2-6*2=0
        //3nn-5n-4*2-52*2=0
        //-c
        BigInteger c = BigInteger.valueOf(z).multiply(BigInteger.valueOf(2)).add((BigInteger.valueOf(num).multiply(BigInteger.valueOf(2))));
        //-4ac
        BigInteger fac = BigInteger.valueOf(4).multiply(c).multiply(BigInteger.valueOf(3));
        //bb-4ac
        BigInteger bbfac = BigInteger.valueOf(((long) Math.pow(5, 2))).add(fac);

        if (BigInteger.ZERO.compareTo(bbfac) <= 0) {
            //have solution is
            //sqrt bb-4ac
            BigDecimal dividendONE = BigDecimal.valueOf(Math.sqrt(Double.valueOf(bbfac.toString())));
            //-b +  sqrt bb-4ac
            BigDecimal dividend = dividendONE.add(BigDecimal.valueOf(5));
            BigDecimal dividend2 = BigDecimal.valueOf(5).subtract(dividendONE);
            //2a
            BigDecimal divisor = BigDecimal.valueOf(6);
            /// (-b +  sqrt bb-4ac)/2a
//            BigDecimal relt = dividend.divide(divisor);
//                    String ssolu = solu.toString();170.0

            int tmp=0;
            try {
//                System.out.println(relt.intValueExact());
//                relt.intValueExact();
                tmp=dividend.divide(divisor).intValueExact();
            } catch (Exception e) {
                return false;
            }

//            tmprelt = relt.intValueExact();
            tmprelt = tmp;
        }

        return true;
    }

    Boolean numIsP(int num) {
        Boolean b = true;
        for (int i = 2; i < num; i++) {
            if (BigInteger.ZERO.equals(BigInteger.valueOf(num).mod(BigInteger.valueOf(i)))) {
                return false;
            }
        }
        return b;
    }

    Boolean numIsEven(int num) {
        Boolean b = true;
//        for (int i = 2; i < num; i++) {
            if (!BigInteger.ZERO.equals(BigInteger.valueOf(num).mod(BigInteger.valueOf(2)))) {
                return false;
            }
//        }
        return b;
    }

}


//https://blog.csdn.net/qq_38637558/article/details/116571437
class RotaEndRecordBorrowExcelHandler implements CellWriteHandler {

    RotaEndRecordBorrowExcelHandler() {
    }

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer, Integer integer1, Boolean aBoolean) {
    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean isHead) {
//        if (cell.getColumnIndex() == 2) {
//            if ("20".equals(cell.getNumericCellValue())) {


        if (cell.getStringCellValue().contains("Z")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.RED.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }

        if (cell.getStringCellValue().contains("P")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(3);
            sheet.setDefaultRowHeight((short) 170);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }

        if (cell.getStringCellValue().contains("E")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }



        if (cell.getStringCellValue().contains("OP")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.GOLD.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }



    }
//    }
}
