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
import java.util.*;

public class FirstTreeNet {
    int row = 900;
    int tmprelt = 0;
    StringBuffer jcoll = new StringBuffer();
    StringBuffer jzccoll = new StringBuffer();
    StringBuffer zfirstcoll = new StringBuffer();
    Map<String, String> zfirstmap = new HashMap<>();

    public FirstTreeNet() {
    }

    public static void main(String[] args) {
        FirstTreeNet firstTreeNet = new FirstTreeNet();
        firstTreeNet.writeExcel();
//        firstTreeNet.dataList();

//        firstTreeNet.getJCollection(9,39);
//        firstTreeNet.getJCollection(18,156);
//        String[] sj = firstTreeNet.jcoll.toString().split(",");
//        String[] var9 = sj;
//        int var10 = sj.length;
//        for(int var11 = 0; var11 < var10; ++var11) {
//            String tmp = var9[var11];
//            System.out.println(tmp);
//        }
    }

    void writeExcel() {
        String fileName = "Y:\\2021job\\@tree\\测试first.xlsx";
        ((ExcelWriterSheetBuilder)EasyExcel.write(fileName).sheet("first模板").registerWriteHandler(new FirstRecordBorrowExcelHandler())).doWrite(this.dataList());
    }

    //    https://segmentfault.com/a/1190000038566393
    private List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList();
        int ma = 1;
        int a = 0;

        for (int i = 0; i < this.row; ++i) {
            for (int j = 0; j < i; ++j) {
                a = Integer.valueOf(j + ma) + 1;
                if (this.numIsInZ(a)) {
                    //i, a
                    this.getJCollection(i, a);
                }
            }
            ma = a;//
        }

        ma = 1;
        a = 0;

        for (int i = 0; i < this.row; ++i) {
            StringBuilder stringBuilder = new StringBuilder();
            for (int j = 0; j < i; ++j) {
                a = Integer.valueOf(j + ma) + 1;
                String stmp = a + "";
                if (this.numIsP(a)) {
                    stmp = "P" + stmp;
                }
//                if (this.numIsInZ(a)) {
//                    stmp = "Z" + stmp;
//                }
                if (this.numIsLessFirst(i, a)) {
                    stmp = "Z" + stmp;
                }
                if (this.numIsInJ(a)) {
                    stmp = "J" + stmp;
                }
                stringBuilder.append(stmp + ",");
            }
            ma = a;//            https://blog.csdn.net/qq_38974638/article/details/117395831
            list.add(new ArrayList(Arrays.asList(stringBuilder.toString().split(","))));
            stringBuilder.delete(0, stringBuilder.toString().length());
        }
        return list;
    }

    Boolean numIsLessFirst(int i, int num) {
        ++i;
        //n(n+1)/2-4=num
        int c = i * (i - 1) / 2 - num;
        if (null != zfirstmap.get(c + "")) {
            if (num <= Integer.valueOf(zfirstmap.get(c + ""))) {
                return true;
            }
        }
        return false;
    }

    Boolean numIsInJ(int num) {
        String[] sj = this.jcoll.toString().split(",");
        String[] var9 = sj;
        int var10 = sj.length;
        for (int var11 = 0; var11 < var10; ++var11) {
            String tmp = var9[var11];
            if ((num + "").equals(tmp)) {
                return true;
            }
        }
        return false;
    }

    Boolean numIsInJZC(int num) {
        String[] sj = this.jzccoll.toString().split(",");
        String[] var9 = sj;
        int var10 = sj.length;
        for (int var11 = 0; var11 < var10; ++var11) {
            String tmp = var9[var11];
            if ((num + "").equals(tmp)) {
                return true;
            }
        }
        return false;
    }

    void getJCollection(int i, int a) {//9-39/////18/156
        //10-5
        //-6  i-9   a-39     +11 5-8    i=9-1
        //-10 i-16  a-126    +21 10-15  i=16-1
        //-15 i-18  a-156    +20 10-17  i=18-1
        //-21 i-23  a-255    +24 12-22  i=23-1
        //+25 12
        int ad = 0;
        int high = i - 1;
        //j-c   //n*(n-1)+c=a     //n(n-1)/2+11=39
        //n(n-1)/2+c=39
        int c = a - (high * (high - 1) / 2);
        //c even odd
        try {
            BigDecimal.valueOf(c).divide(BigDecimal.valueOf(2)).intValueExact();
            ad = c;
        } catch (Exception e) {
            ad = c - 1;
        }
//        System.out.println(ad);//12

        int jzc = i * (i + 1) / 2 - a;
        Boolean hasP = false;

        if (!numIsInJZC(jzc)) {
            //(jc(+1)-11)/2+5
            int low = (ad - 10) / 2 + 5;
            StringBuffer tmpJrelt = new StringBuffer();
            for (int j = low; j <= high; j++) {
                int num = j * (j - 1) / 2 + c;
                tmpJrelt.append(num + ",");
                if (this.numIsP(num)) {
                    tmpJrelt.delete(0, tmpJrelt.toString().length());
                    hasP = true;
                    break;
                }
            }
            if (!hasP) {
                //6，10，15
                this.jzccoll.append(jzc + ",");

                //hashmap
                //jzc tmpJrelt[*]
                //6-39,
                // 10-126,
                // 15-156
                zfirstmap.put(jzc + "", tmpJrelt.toString().split(",")[tmpJrelt.toString().split(",").length - 1]);

                //21,26,32,39
                //66,76,87,99,112,126
                this.jcoll.append(tmpJrelt);
            }
        }
    }

    Boolean numIsInZ(int num) {//6- 22,30,39,49,60,72
        //6,10,15,21,28,36,45,55,66,78,91
        int z = 3;
        int a = 3;
        for (int i = 0; i < this.row; i++) {
            z = z + a;
            if (reltIsInt(z, num) && 0 != tmprelt) {
                if ((z + 1) <= tmprelt) {
                    return true;
                }
            }
            tmprelt = 0;
            a++;
        }
        return false;
    }

    Boolean reltIsInt(int z, int num) {
        //nn+n-6*2-num*2=0
        //-c
        BigInteger c = BigInteger.valueOf(z).multiply(BigInteger.valueOf(2)).add((BigInteger.valueOf(num).multiply(BigInteger.valueOf(2))));
        //-4ac
        BigInteger fac = BigInteger.valueOf(4).multiply(c);
        //bb-4ac
        BigInteger bbfac = BigInteger.ONE.add(fac);

        if (BigInteger.ZERO.compareTo(bbfac) <= 0) {
            //have solution is
            //sqrt bb-4ac
            BigDecimal dividendONE = BigDecimal.valueOf(Math.sqrt(Double.valueOf(bbfac.toString())));
            //-b +  sqrt bb-4ac
            BigDecimal dividend = dividendONE.subtract(BigDecimal.ONE);
            //2a
            BigDecimal divisor = BigDecimal.valueOf(2);
            /// (-b +  sqrt bb-4ac)/2a
            BigDecimal relt = dividend.divide(divisor);
//                    String ssolu = solu.toString();170.0

            try {
//                System.out.println(relt.intValueExact());
                relt.intValueExact();
            } catch (Exception e) {
                return false;
            }
            tmprelt = relt.intValueExact();
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

}


//https://blog.csdn.net/qq_38637558/article/details/116571437
class FirstRecordBorrowExcelHandler implements CellWriteHandler {
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
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
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
            cellStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.index);
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

        if (cell.getStringCellValue().contains("J")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
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