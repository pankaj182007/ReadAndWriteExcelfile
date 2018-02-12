package com.ReadAndWriteXLfile;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_And_WriteXLfile {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		FileInputStream fileIn=new FileInputStream("F:\\selenium\\ReadAndWriteXLfile\\read.xlsx");
		XSSFWorkbook wbook=new XSSFWorkbook(fileIn);
		XSSFSheet sheet=wbook.getSheetAt(0);
		
		Cell c1;
		Row r1;
		
		FileOutputStream fileOut=new FileOutputStream("F:\\selenium\\ReadAndWriteXLfile\\write.xlsx");
		XSSFWorkbook wbookout=new XSSFWorkbook();
		XSSFSheet sheet2=wbookout.createSheet("write data");
		
		
		int i=0;
		Iterator<Row> r2=sheet.rowIterator();
		while (r2.hasNext()) 
		{
			r1= r2.next();
			Row rw=sheet2.createRow(i++);
			
			Iterator<Cell> c2=r1.cellIterator();
			int j=0;
			while (c2.hasNext()) 
			{
				c1=c2.next();
	//reading cell			
				if (c1.getCellType()==XSSFCell.CELL_TYPE_STRING)
				{
					
					System.out.print(c1.getStringCellValue()+"||");

				}
				else if (c1.getCellType()==XSSFCell.CELL_TYPE_NUMERIC)
				{
					System.out.print(c1.getNumericCellValue()+"||");
				}
				
	//writing cell	
				
				Cell cw=rw.createCell(j++);
				if (c1.getCellType()==XSSFCell.CELL_TYPE_STRING)
				{
					cw.setCellValue(c1.getStringCellValue());
				}
				else if (c1.getCellType()==Cell.CELL_TYPE_NUMERIC)
				{
					cw.setCellValue(c1.getNumericCellValue());
				}
			}
			System.out.println();
			System.out.println("Row "+i+" is written successfully:");
			System.out.println();
		}	
		
		wbookout.write(fileOut);
		System.out.println("exel file written successfully");
		wbookout.close();
		
		}
		

	}


