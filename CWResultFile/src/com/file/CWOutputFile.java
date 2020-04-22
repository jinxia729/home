package com.file;	
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
	 
import jxl.Cell;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
	 
 
public class CWOutputFile {
	
	/*
	 * wOutputFile����д����ļ�
	 * wOutputFile(�ļ�·��,�������,������֤��,��������,Ԥ�ڽ��,ʵ�ʽ��)
	 */
	public void wOutputFile(String filepath, String caseNo, String testPoint, String url,String methods,String testData, String preResult, String fresult) throws WriteException, BiffException, IOException {
	    File output = new File(filepath);
		String result = "";
		InputStream instream = new FileInputStream(filepath);
		Workbook readwb = Workbook.getWorkbook(instream);
		// �����ļ�����һ����������
		WritableWorkbook wbook = Workbook.createWorkbook(output, readwb); 
		WritableSheet readsheet = wbook.getSheet(0);
		
		// int rsColumns = readsheet.getColumns(); //��ȡSheet������������������
		// ��ȡSheet������������������
		int rsRows = readsheet.getRows(); 
		/******************************** ������ʽ���� ****************************/
		
		// ������ʽ
		WritableFont font = new WritableFont(WritableFont.createFont("����"), 10,
		WritableFont.NO_BOLD);
		WritableCellFormat wcf = new WritableCellFormat(font);
 
		/***********************************************************************/
 
		Cell cell1 = readsheet.getCell(0, rsRows);
 
		if (cell1.getContents().equals("")) {
			Label labetest1 = new Label(0, rsRows, caseNo); // ��1��--������ţ�
			Label labetest2 = new Label(1, rsRows, testPoint); // ��2��--�������⣻
			Label labetest3 = new Label(2, rsRows, url); // ��3��--url��
			Label labetest4 = new Label(3, rsRows, methods); // ��4��--����ʽ��
			Label labetest5 = new Label(4, rsRows, testData); // ��5��--�������ݣ�
			Label labetest6 = new Label(5, rsRows, preResult); // ��6��--Ԥ�ڽ��
			Label labetest7 = new Label(6, rsRows, fresult); // ��7��--ʵ�ʽ����
			
			// ����ֵͬʱ��ȲŻ���ʾͨ��
			if (fresult.contains(preResult)) {
				result = "ͨ��";
				wcf.setBackground(Colour.BRIGHT_GREEN); // ͨ��������ע��ɫ
			} else {
				result = "��ͨ��";
				wcf.setBackground(Colour.RED); // ��ͨ��������ע��ɫ
			}
			Label labetest8 = new Label(7, rsRows, result,wcf); // ��8��--ִ�н����
			readsheet.addCell(labetest1);
			readsheet.addCell(labetest2);
			readsheet.addCell(labetest3);
			readsheet.addCell(labetest4);
			readsheet.addCell(labetest5);
			readsheet.addCell(labetest6);
			readsheet.addCell(labetest7);
			readsheet.addCell(labetest8);
		}
 
		wbook.write();
		wbook.close();
	}
 
	/*
	 * cOutputFile������������ļ� 
	 * cOutputFile���������ļ�·������ΪwOutputFile�����
	 */
	public String cOutputFile(String tradeType) throws IOException,
			WriteException {
		String temp_str = "";
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
		// ��ȡʱ���
		temp_str = sdf.format(date);
		// ��ʱ�����������ļ���ȷ��Ψһ
		String filepath = "E:\\result\\" + tradeType + "_output_" + "_" + temp_str
				+ ".xls";
		File output = new File(filepath);
 
		if (!output.isFile()) {
			// ���ָ���ļ������ڣ��򴴽����ļ�
			output.createNewFile();
			WritableWorkbook writeBook = Workbook.createWorkbook(output);
			// createSheet(Sheet���ƣ��ڼ���Sheet)
			WritableSheet Sheet = writeBook.createSheet("������", 0);
			// ������ʽ
			WritableFont headFont = new WritableFont(
					WritableFont.createFont("����"), 11, WritableFont.BOLD);
			WritableCellFormat headwcf = new WritableCellFormat(headFont);
			// ��ɫ
			headwcf.setBackground(Colour.GRAY_25);
			// Sheet.setColumnView(�к�,���)
			Sheet.setColumnView(0, 10);
			Sheet.setColumnView(1, 30);
			Sheet.setColumnView(2, 30);
			Sheet.setColumnView(3, 10);
			Sheet.setColumnView(4, 50);
			Sheet.setColumnView(5, 40);
			Sheet.setColumnView(6, 50);
			Sheet.setColumnView(7, 10);
			
			// �������־��ж��뷽ʽ
			headwcf.setAlignment(Alignment.CENTRE);
			// ���ô�ֱ����
			headwcf.setVerticalAlignment(VerticalAlignment.CENTRE);
			// Label(�к�,�к�,����)
			Label label00 = new Label(0, 0, "�������", headwcf);
			Label label10 = new Label(1, 0, "��������", headwcf);
			Label label20 = new Label(2, 0, "url", headwcf);
			Label label30 = new Label(3, 0, "����ʽ", headwcf);
			Label label40 = new Label(4, 0, "��������", headwcf);
			Label label50 = new Label(5, 0, "Ԥ�ڽ��", headwcf);
			Label label60 = new Label(6, 0, "ʵ�ʽ��", headwcf);
			Label label70 = new Label(7, 0, "ִ�н��", headwcf);
			
			Sheet.addCell(label00);
			Sheet.addCell(label10);
			Sheet.addCell(label20);
			Sheet.addCell(label30);
			Sheet.addCell(label40);
			Sheet.addCell(label50);
			Sheet.addCell(label60);
			Sheet.addCell(label70);
			
			writeBook.write();
			writeBook.close();
		}
 
		return filepath;
	}
}
