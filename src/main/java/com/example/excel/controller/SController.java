package com.example.excel.controller;

import com.example.excel.domain.S;
import com.example.excel.service.SService;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.List;


@Controller
public class SController {
    @Autowired
    private SService sService;

    @RequestMapping("/getAllFormExcel")
    public String getAllFormExcel(Model model){
        //需要解析的Excel文件
        File file=new File("D:/s.xls");
        List<S> list=new ArrayList<S>();//返回的集合
        try {
            //获取工作薄
            FileInputStream fileInputStream= FileUtils.openInputStream(file);
            HSSFWorkbook workbook=new HSSFWorkbook(fileInputStream);
            //获取第一个工作表
            HSSFSheet hs=workbook.getSheetAt(0);
            //获取该工作表的第一个行号和最后一个行号
            int last=hs.getLastRowNum();
            int first=hs.getFirstRowNum();
            //遍历获取单元格信息
            for(int i=first+1;i<=last;i++){//第一行是String类型
                HSSFRow row=hs.getRow(i);
                int firstCellNum=row.getFirstCellNum();
                int lastCellNum=row.getLastCellNum();
                S s=new S();
                for(int j=firstCellNum;j<lastCellNum;j++){

                    HSSFCell cell=row.getCell(j);
                    //读取数据前设置单元格类型
//                    cell.setCellType(CellType.STRING);
//                    String value=cell.getStringCellValue();
//                    System.out.print(value+" ");
                    if(j==firstCellNum){
                        s.setId((int) cell.getNumericCellValue());
                    }else if(j==firstCellNum+1){
                        s.setName(cell.getStringCellValue());
                    }else if(j==firstCellNum+2){
                        s.setCourse(cell.getStringCellValue());
                    }else if (j==firstCellNum+3){
                        s.setScore(cell.getNumericCellValue());
                    }
                }
                list.add(s);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        model.addAttribute("list",list);
        List<S> list1=sService.getAllStudent();
        model.addAttribute("list1",list1);
        return "index";
    }
    @RequestMapping("/export")
    @ResponseBody
    public String createExcel(HttpServletResponse response) throws IOException {
        //从数据库先获取集合信息
        List<S> sList=sService.getAllStudent();
        String title[]={"编号","姓名","课程","分数"};
        //1.创建Excel工作簿,不传入参数
        HSSFWorkbook workbook=new HSSFWorkbook();
        //2.创建一个工作表 名字为sheet2
        HSSFSheet sheet=workbook.createSheet("sheet2");
        //3.创建第一行
        HSSFRow row=sheet.createRow(0);
        HSSFCell cell=null;
        //4.插入第一行数据
        for (int i=0;i<title.length;i++){
            cell=row.createCell(i);
            cell.setCellValue(title[i]);
//            sheet.autoSizeColumn(i);//单元格宽度自适应
        }
        int i;
        //5.插入其他数据
        for (i=1;i<=sList.size();i++){
            row=sheet.createRow(i);//创建行
            S s=sList.get(i-1);
            for(int j=0;j<title.length;j++){
                cell=row.createCell(j);
                if (j==0){
                    cell.setCellValue(s.getId());
                }else if (j==1){
                    cell.setCellValue(s.getName());
                }else if (j==2){
                    cell.setCellValue(s.getCourse());
                }else if(j==3){
                    cell.setCellValue(s.getScore());
                }else{
                    cell.setCellValue("字段列超过了设定的范围");
                }
            }
        }
        //最后一行做函数运算
        row=sheet.createRow(i);//创建行
        cell=row.createCell(0);
        cell.setCellValue("总人数:");
        cell=row.createCell(1);
//        cell.setCellFormula("COUNT(C1,C"+i+")");
        cell.setCellValue(i-1);
        cell=row.createCell(2);
        cell.setCellValue("最大成绩:");
        cell=row.createCell(3);
        cell.setCellFormula("MAX(D2:D"+i+")");
        cell=row.createCell(4);
        cell.setCellValue("成绩总和");
        cell=row.createCell(5);
        cell.setCellFormula("SUM(D2:D"+i+")");
        //保存到固定的位置
//        File file=new File("D://生成的Excel表.xls");
//        file.createNewFile();
//        FileOutputStream outputStream=FileUtils.openOutputStream(file);
//        workbook.write(outputStream);
//        outputStream.close();
        //输出Excel文件
        OutputStream outputStream=response.getOutputStream();
        response.reset();
        //文件名这里可以改
        response.setHeader("Content-disposition", "attachment; filename=report.xls");
        response.setContentType("application/msexcel");
        workbook.write(outputStream);
        outputStream.close();
        return "success";//实战中应该通过异步处理
    }
    @RequestMapping("/export1")
    @ResponseBody
    public String createCellStyleExcel(HttpServletResponse response){
        //1.创建工作薄
        HSSFWorkbook workbook=new HSSFWorkbook();
        //2.创建工作表
        HSSFSheet sheet=workbook.createSheet("cellStyle");
        //设置第几列的单元格宽度sheet.setColumnWidth(0, 256*width+184);
        sheet.setColumnWidth(3,256*20+184);
        //3.创建行
        HSSFRow row=sheet.createRow(1);
        //设置行高
        row.setHeight((short) 800);
        HSSFCell cell=row.createCell(1);
        cell.setCellValue("test of merging");
        //合并单元格，功能浅显易懂
        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row (0-based)
                1, //first column (0-based)
                4 //last column (0-based)
        ));
        //设置背景色
        row=sheet.createRow(2);
        cell=row.createCell(2);
        HSSFCellStyle cellStyle=workbook.createCellStyle();//创建样式
        cellStyle.setFillBackgroundColor((short) 13);//背景颜色
        //设置单元格无边框
//        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置文字居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("麻了");
        //---------------------------
        //创建下一个单元格
        cell=row.createCell(3);
        cell.setCellValue("改变字体改变隔膜改变小气我可以改变自己");
        HSSFCellStyle cellStyle1=workbook.createCellStyle();//创建样式
        HSSFFont font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 16);//设置字体大小

        HSSFFont font2 = workbook.createFont();
        font2.setFontName("仿宋_GB2312");
        font2.setBold(true);//粗体
        font2.setFontHeightInPoints((short) 12);//设置字体大小

        cellStyle1.setFont(font);//选择需要用到的字体格式
        //设置自动换行
//        cellStyle1.setWrapText(true);
        cell.setCellStyle(cellStyle1);
        //--------------------------------------------

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(
                    new File("cellstyle.xlsx"));
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "cellstyle.xlsx written successfully";
    }
}
