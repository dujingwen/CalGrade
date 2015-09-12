package com.CalGrade.PoiExcel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class XlsDto2Excel {
	public static void xlsDto2Excel(List<Result> xls) throws Exception{
		
		//��ȡ������
		int CountColumnNum = 5;
		
		//����Excel�ĵ�
		HSSFWorkbook hwb = new HSSFWorkbook();
		Result result = null;
		HSSFSheet sheet = hwb.createSheet("db");
		HSSFRow firstrow = sheet.createRow(0);
		HSSFCell[] firstcell = new HSSFCell[CountColumnNum];
		String[] names = new String[CountColumnNum];
		names[0] = "ѧ��";
		names[1] = "����";
		names[2] = "A1";
		names[3] = "A2";
		names[4] = "F1";
		for(int j=0; j < CountColumnNum; j++){
			firstcell[j] = firstrow.createCell(j);
			firstcell[j].setCellValue(new HSSFRichTextString(names[j]));
		}
		for(int i=0; i < xls.size();i++){
			//����һ��
			HSSFRow row = sheet.createRow(i+1);
			result = xls.get(i);
			for(int colu = 0; colu <= 4;colu++){
				HSSFCell xh = row.createCell(0);
				xh.setCellValue(result.getSno());
				HSSFCell xm = row.createCell(1);
				xm.setCellValue(result.getSname());
				HSSFCell a1 = row.createCell(2);
				a1.setCellValue(result.getA1());
				HSSFCell a2 = row.createCell(3);
				a2.setCellValue(result.getA2());
				HSSFCell f1 = row.createCell(4);
				f1.setCellValue(result.getF1());
				
			}
		}
		
		//�����ļ�����������Excel
		OutputStream out = new FileOutputStream("F:/JavaProject/dbBook.xls");
		hwb.write(out);
		out.close();
		System.out.println("���ݿ⵼���ɹ�");
	}
}
