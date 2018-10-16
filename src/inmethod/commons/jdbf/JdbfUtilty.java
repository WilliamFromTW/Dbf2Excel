package inmethod.commons.jdbf;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Vector;

import inmethod.commons.rdb.DataSet;
import inmethod.jakarta.excel.CreateXLSX;
import net.iryndin.jdbf.core.DbfField;
import net.iryndin.jdbf.core.DbfMetadata;

public class JdbfUtilty {
	private static JdbfUtilty aJdbfUtilty = null;

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			JdbfUtilty.getInstance().exportDbfToExcel("cp1251", "c:\\tmp\\cp1251.dbf", "c:/tmp/sss.xlsx");
			//exportDbfToExcel("big5", "c:\\tmp\\cp1251.dbf", "c:/tmp/sss.xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private  JdbfUtilty() {}
	
	public static JdbfUtilty getInstance() {
		if( aJdbfUtilty==null)
			aJdbfUtilty = new JdbfUtilty();
		return aJdbfUtilty;
	}
	/**
	 * @param sCharset dbf record charset
	 * @param sDbfFile dbf file location
	 * @param sOutputExcelFile
	 * @throws Exception
	 */
	public void exportDbfToExcel(String sCharset, String sDbfFile, String sOutputExcelFileName)
			throws Exception {
		Charset stringCharset = Charset.forName(sCharset);
		File file = new File(sDbfFile);
		net.iryndin.jdbf.reader.DbfReader reader = new net.iryndin.jdbf.reader.DbfReader(file);
		DbfMetadata meta = reader.getMetadata();
		// System.out.println(meta); 

		Collection<DbfField> aFields = meta.getFields();

		DataSet aDSHeader = new DataSet();
		Vector<String> aDataCell = new Vector<String>();
		for (DbfField aField : aFields) {
			if (aField.getName() != null && !aField.getName().trim().equals("")) {
				//System.out.println("table field name = " + aField.getName());
				aDataCell.add(aField.getName());
			}
		}
		aDSHeader.addData(aDataCell);

		net.iryndin.jdbf.core.DbfRecord rec = null;

		DataSet aDSData = new DataSet();
		Vector<String> aDSDataRecord =null;
		while ((rec = reader.read()) != null) {
			rec.setStringCharset(stringCharset);
			 aDSDataRecord = new Vector<String>();
			for(String sFieldName:aDataCell) {
				if( rec.getString(sFieldName)!=null ) {
					aDSDataRecord.add(rec.getString(sFieldName));
			//	System.out.print(rec.getString(sFieldName));
				}else {
					aDSDataRecord.add("");
				}
			}
			aDSData.addData(aDSDataRecord);

		}
		reader.close();
		FileOutputStream aFO = new FileOutputStream(new File(sOutputExcelFileName));
		CreateXLSX aFE = new CreateXLSX(aFO);
		aFE.setCurrentSheet();
		aFE.setPrintResultSetHeader(false);
		aFE.calculateExcel(aDSHeader);
		aFE.calculateExcel(aDSData);

		aFE.buildExcel();

	}
}
