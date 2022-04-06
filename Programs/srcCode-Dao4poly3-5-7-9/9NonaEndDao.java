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

public class NonaEnd {
//    int row = 900;
    int row = 400;
    int max=0;
    int tmprelt = 0;

    public NonaEnd() {
    }

    public static void main(String[] args) {
        NonaEnd nonaEnd = new NonaEnd();
        nonaEnd.writeExcel();
    }

    void writeExcel() {
        String fileName = "Y:\\2021job\\@tree\\测试NonaEnd.xlsx";
        ((ExcelWriterSheetBuilder)EasyExcel.write(fileName).sheet("NonaEnd模板").registerWriteHandler(new NonaEndRecordBorrowExcelHandler())).doWrite(this.dataList());
    }

    //    https://segmentfault.com/a/1190000038566393
    private List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList();
        int ma = 0;
        int a = 0;

        for(int i = 1; i < this.row; ++i) {
            StringBuilder stringBuilder = new StringBuilder();

            max = (7*i*i-5*i)/2;;
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

    Boolean numIsInZ(int num) {//-3- 25

        //-3，0，6，12，22，31，45，57
        int z = -3;
        int a = 3;
        int b = 6;
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
                a=a+3;
                flag=1;
            }else{
                z = z + b;
                b=b+4;
                flag=0;
            }
        }
        return false;
    }

    Boolean reltIsInt(int z, int num) {
        //7nn-17n-z*2-num*2=0
        //7nn-17n-(-3)*2-num*2=0
        //7nn-17n-(0)*2-num*2=0
        //-c
        BigInteger c = BigInteger.valueOf(z).multiply(BigInteger.valueOf(2)).add((BigInteger.valueOf(num).multiply(BigInteger.valueOf(2))));
        //-4ac
        BigInteger fac = BigInteger.valueOf(4).multiply(c).multiply(BigInteger.valueOf(7));
        //bb-4ac
        BigInteger bbfac = BigInteger.valueOf(((long) Math.pow(17, 2))).add(fac);

        if (BigInteger.ZERO.compareTo(bbfac) <= 0) {
            //have solution is
            //sqrt bb-4ac
            BigDecimal dividendONE = BigDecimal.valueOf(Math.sqrt(Double.valueOf(bbfac.toString())));
            //-b +  sqrt bb-4ac
            BigDecimal dividend = dividendONE.add(BigDecimal.valueOf(17));
            //2a
            BigDecimal divisor = BigDecimal.valueOf(14);
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
class NonaEndRecordBorrowExcelHandler implements CellWriteHandler {

    NonaEndRecordBorrowExcelHandler() {
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
