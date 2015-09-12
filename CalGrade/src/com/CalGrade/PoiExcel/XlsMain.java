package com.CalGrade.PoiExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class XlsMain {
	public static void main(String[] args) throws IOException {  
        XlsMain xlxsMain = new XlsMain();  
        List<Result> list = new ArrayList<Result>();
        Result xls = null;
        list = xlxsMain.readXls();
        
        try {
			XlsDto2Excel.xlsDto2Excel(list);
		} catch (Exception e) {
			e.printStackTrace();
		}
        
     /*   for(int i = 0;i<list.size();i++){
        	xls = (Result)list.get(i);
        	System.out.println(xls.getSno()+"     "+xls.getSname()+"    "+xls.getA1()+"    "+xls.getA2()+"    "+xls.getF1());
        }   */
	}
	
	Vector<Double> bxlist = new Vector<Double>();
	Vector<Double> xxlist = new Vector<Double>();
	Vector<Double> bxxf = new Vector<Double>();
	Vector<Double> xxxf = new Vector<Double>();
	
	private List<Result> readXls() throws IOException{
		InputStream is = new FileInputStream("F:\\JavaProject\\�籾.xls");
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		Result result = null;
		List<Result> list = new ArrayList<Result>();
		for(int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if(hssfSheet == null){
				continue;
			}
			
			
			//��ȡ�γ�����
			int num = 3;
			HSSFRow hssfRow = hssfSheet.getRow(num);
			if(hssfRow==null){
				continue;
			}
			int a[] = new int[100];
			int c = 0;
			int StuNum = 0;
			for(int col = 3;col <= hssfRow.getLastCellNum(); col++){
				HSSFCell cell = hssfRow.getCell(col);
				if(cell == null){
					continue;
				}
				String str = getValue(cell);
				if(str.equals("")){
					continue;
				}
				if(str.contains("����")){
					a[c++] = 1;
				}
				else if(str.contains("ѡ��")){
					a[c++] = 2;
				}
			}
			
			
			//��ȡ�γ�ѧ��
			Vector<Double> credit = new Vector<Double>();
			int num1 = 4;
			HSSFRow cre = hssfSheet.getRow(num1);
			if(cre==null){
				continue;
			}
			for(int col = 3;col<3+c; col++){
				HSSFCell cell = cre.getCell(col);
				if(cell == null){
					continue;
				}
				Double xf = Double.parseDouble(getValue(cell).replace("��", ""));
				credit.addElement(xf);
			}
			
			
			//ѭ��������
			for(int rowNum = 6; rowNum<=hssfSheet.getLastRowNum();rowNum++){
				bxlist.clear();
				xxlist.clear();
				bxxf.clear();
				xxxf.clear();
				hssfRow = hssfSheet.getRow(rowNum);
				if(hssfRow == null){
					continue;
				}
				result = new Result();
				HSSFCell sno = hssfRow.getCell(0);
				if(sno == null){
					continue;
				}
				if(getValue(sno).equals(""))
					break;
				else{
					result.setSno(getValue(sno));
					StuNum++;
				}
				HSSFCell sname = hssfRow.getCell(1);
				if(sname == null){
					continue;
				}
				result.setSname(getValue(sname));
				
		
				for(int colNum = 3;colNum<3+c;colNum++){
				//	System.out.println(colNum+"        "+a[colNum-3]);
					HSSFCell score = hssfRow.getCell(colNum);
					if(score==null){
						continue;
					}
					String tmp = getValue(score).trim();
					if(tmp.equals("")){
						continue;
					}
					if(tmp.contains(",")){
						double sco = solveMultiScore(tmp);
						if(a[colNum-3]==1){
							bxlist.addElement(sco);
							bxxf.addElement(credit.elementAt(colNum-3));
						}
						else{
							xxlist.addElement(sco);
							xxxf.addElement(credit.elementAt(colNum-3));
						}
						continue;
					}
					int flag = 0;
					for(int i=0;i<=9;i++){
						if(tmp.indexOf('0'+i)!=-1){
							flag = 1;
							break;
						}
					}
					for(int i=0;i<=3;i++){
						if(tmp.indexOf('A'+i)!=-1){
							flag = 2;
							break;
						}
					}
					
					double x = 0;
					if(flag==1){
						if(a[colNum-3]==1){
							bxlist.addElement(Double.parseDouble(tmp));
							bxxf.addElement(credit.elementAt(colNum-3));
						}
						else{
							xxlist.addElement(Double.parseDouble(tmp));
							xxxf.addElement(credit.elementAt(colNum-3));
						}
					}
					else if(flag==0){
						x = solveHz(tmp);
						if(a[colNum-3]==1){
							bxlist.addElement(x);
							bxxf.addElement(credit.elementAt(colNum-3));
						}
						else{
							xxlist.addElement(x);
							xxxf.addElement(credit.elementAt(colNum-3));
						}
					}
					else{
						x = solveZm(tmp);
						if(a[colNum-3]==1){
							bxlist.addElement(x);
							bxxf.addElement(credit.elementAt(colNum-3));
						}
						else{
							xxlist.addElement(x);
							xxxf.addElement(credit.elementAt(colNum-3));
						}
					}
				}
								
				//����A1
				
				float score = 0,sum = 0;
				for(int k=0;k<bxlist.size();k++){
					sum += bxxf.elementAt(k);
					score += bxlist.elementAt(k)*bxxf.elementAt(k);
					result.setA1(score/sum);
				}
				
				if(rowNum==6){
					System.out.println("sum="+sum+"   score="+score);
				}
				
				//����A2
				
				float score1 = 0;
				for(int k=0;k<xxlist.size();k++){
					score1 += xxlist.elementAt(k)*xxxf.elementAt(k)*0.002;
					result.setA2(score1);
				}
				
				
				//����F1
				result.setF1(score/sum+score1);
				list.add(result);
				//System.out.println("size = "+list.size());
			} 
			
		}
		return list;
	}
	
	
	
	private String getValue(HSSFCell hssfCell) {  
        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {  
            // ���ز������͵�ֵ  
            return String.valueOf(hssfCell.getBooleanCellValue());  
        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {  
            // ������ֵ���͵�ֵ  
            return String.valueOf(hssfCell.getNumericCellValue());  
        } else {  
            // �����ַ������͵�ֵ  
            return String.valueOf(hssfCell.getStringCellValue());  
        }  
    }  
	
	private double solveMultiScore(String str){
		double ret = 0.0;
		String scores[] = new String[8];
		scores = str.split(",");
		
		//�ж������֡���ĸ���Ǻ���
		int flag = 0;
		for(int i=0;i<=9;i++){
			if(scores[0].indexOf('0'+i)!=-1){
				flag = 1;
				break;
			}
		}
		for(int i=0;i<=3;i++){
			if(scores[0].indexOf('A'+i)!=-1){
				flag = 2;
				break;
			}
		}
		
		if(flag==0){  //����
			double x = solveHz(scores[0]);
			if(x>=60.0) ret = x;
		}
		else if(flag==1){   //����
			double tmp = Double.parseDouble(scores[0]);
			if(tmp==0.0){
				double x = Double.parseDouble(scores[1]);
				if(x>=60.0) ret = x;
			}
			else{
				if(tmp>=60.0) ret = tmp;
			}
		}
		else{  //��ĸ
			double x = solveZm(scores[0]);
			if(x>=60.0) ret = x;
		}
			
		return ret;
	}
	
	private double solveZm(String str){
		double x = 0;
		if(str.contains("A-")){
			x = 87;
		}
		else if(str.contains("A")){
			x = 90;
		}
		else if(str.contains("B+")){
			x = 83;
		}
		else if(str.contains("B-")){
			x = 77;
		}
		else if(str.contains("B")){
			x = 79;
		}
		else if(str.contains("C+")){
			x = 73;
		}
		else if(str.contains("C-")){
			x = 65;
		}
		else if(str.contains("C")){
			x = 70;
		}
		else if(str.contains("D")){
			x = 61;
		}
		return x;
	}
	
	private double solveHz(String str){
		double x = 0;
		if(str.contains("����")){
			x = 90;
		}
		else if(str.contains("����")){
			x = 80;
		}
		else if(str.contains("�е�")){
			x = 70;
		}
		else if(str.contains("����")){
			x = 60;
		}
		return x;
	}
}
