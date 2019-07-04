package com.example.excel.controller;

import com.example.excel.domain.S;
import com.example.excel.service.SService;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
        }

        //5.插入其他数据
        for (int i=1;i<=sList.size();i++){
            sheet.autoSizeColumn(i);//单元格宽度自适应
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
}
