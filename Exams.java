package com.mannir.servlets;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItemIterator;
import org.apache.commons.fileupload.FileItemStream;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.fileupload.util.Streams;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import com.google.appengine.api.datastore.Entity;
import com.google.appengine.api.datastore.Key;
import com.google.appengine.api.datastore.KeyFactory;
import com.google.appengine.api.datastore.Query;
import com.google.appengine.api.datastore.DatastoreService;
import com.google.appengine.api.datastore.DatastoreServiceFactory;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Exams extends HttpServlet {
  public String id="", uid="", pincode="", regno="", password="", mobileno="", fullname="", school="", department="", programme="", session="", courses1="", courses2="", bankname="", tellerno="", amount="", datereg="", mail="", created="", filename="";
	private Connection cn;
	
    public void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException {
    	response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=ExamResults.xls");
        
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sh = wb.createSheet("ExamResults");
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] {"Reg No.", "Fullname", "Course", "CU", "CA", "Exam", "Total", "Grade", "Point", "GP", "Remarks"});
        
    	DatastoreService ds = DatastoreServiceFactory.getDatastoreService();
    	
    	boolean isMultipart = ServletFileUpload.isMultipartContent(request);
    	
    	// Create a new file upload handler
    	ServletFileUpload upload = new ServletFileUpload();

    	// Parse the request
    	
    	try {
    	FileItemIterator iter = upload.getItemIterator(request);
    	while (iter.hasNext()) {
    	    FileItemStream item = iter.next();
    	    String name = item.getFieldName();
    	    InputStream stream = item.openStream();
    	    if (item.isFormField()) {
    	        System.out.println("Form field " + name + " with value "
    	            + Streams.asString(stream) + " detected.");
    	    } else {
    	        System.out.println("File field " + name + " with file name "
    	            + item.getName() + " detected.");
    	        // Process the input stream
    	        
    			try {
    				//DatastoreService ds = DatastoreServiceFactory.getDatastoreService();  
    			    //FileInputStream file = new FileInputStream(new File("exams.xls"));
    			     
    			    //Get the workbook instance for XLS file 
    			    HSSFWorkbook workbook = new HSSFWorkbook(stream);
    			 
    			    //Get first sheet from the workbook
    			    HSSFSheet sheet = workbook.getSheetAt(0);
    			     
    			    //Iterate through each rows from first sheet
    			    Iterator<Row> rowIterator = sheet.iterator();
    			    
    		
    			    int a=2;
    			    while(rowIterator.hasNext()) {
    			    	
    			    	
    			        Row row = rowIterator.next();
    			        
    			        if(row.getRowNum()>0) {
    			        	
            			String fn = row.getCell(0).toString();
            		    String rn = row.getCell(1).toString();
        			    String cs = row.getCell(2).toString();
        			    int cu = Integer.parseInt(row.getCell(3).toString());   			        
    			        int ca = Integer.parseInt(row.getCell(4).toString());
    			        int ex = Integer.parseInt(row.getCell(5).toString());

    			        
    			        int tt = Integer.parseInt(gd("tt",cu, ca,ex));
    			        String gr = gd("gr",cu, ca,ex);
    			        double pt = Double.parseDouble(gd("pt",cu, ca,ex));
    			        double gp = Double.parseDouble(gd("gp",cu, ca,ex));
    			        String rm = gd("rm", cu, ca,ex);
    			        
    			        Entity e1 = new Entity("EXAMS");
    			        e1.setProperty("regno",rn);
    			        e1.setProperty("course",cs);
    			        e1.setProperty("cu",cu);
    			        e1.setProperty("ca",ca);
    			        e1.setProperty("exam",ex);
    			        e1.setProperty("total",tt);
    			        e1.setProperty("grade",gr);
    			        e1.setProperty("point",pt);
    			        e1.setProperty("gp",gp);
    			        e1.setProperty("remarks",rm);
    			        e1.setProperty("fullname",fn);

    			        ds.put(e1);
    			        
    			      	 data.put(a+"", new Object[] {fn, rn, cs, cu+"", ca+"", ex+"", tt+"", gr, pt, gp, rm});

    			        
    			       a++;
    			        
    			        }
    			       // System.out.println(row.getCell(0)+"=="+row.getCell(1));
    			         
    			        //For each row, iterate through each columns
    			      ///  Iterator<Cell> cellIterator = row.cellIterator();
    			      ///  while(cellIterator.hasNext()) {
    			             
    			         // /  Cell cell = cellIterator.next();
    			             

    			            /**
    			            switch(cell.getCellType()) {
    			                case Cell.CELL_TYPE_BOOLEAN:
    			                    System.out.print(cell.getBooleanCellValue() + "\t\t");
    			                    break;
    			                case Cell.CELL_TYPE_NUMERIC:
    			                    System.out.print(cell.getNumericCellValue() + "\t\t");
    			                    break;
    			                case Cell.CELL_TYPE_STRING:
    			                    System.out.print(cell.getStringCellValue() + "\t\t");
    			                    break;
    			            }
    			            */
    			        }
    			    
    			        System.out.println("");
    			  ///  }
    			    //file.close();
    			   // FileOutputStream out =  new FileOutputStream(new File("test.xls"));
    			  //  workbook.write(out);
    			    //out.close();
    			     
} catch (FileNotFoundException e) { e.printStackTrace(); } catch (IOException e) { e.printStackTrace();	} 	        
     
}}} catch(Exception e1) { System.out.println(e1);  }	
    	

/**
        // create a workbook , worksheet
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("ExamResults");
        CreationHelper createHelper = wb.getCreationHelper();

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow((short)0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue( createHelper.createRichTextString("This is a string") );
        row.createCell(3).setCellValue(true);

        //write workbook to outputstream
        ServletOutputStream out = response.getOutputStream();
        wb.write(out);
        out.flush();
        out.close();   	
 */
    	 
    	 
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
		    Row row = sh.createRow(rownum++);
		    Object [] objArr = data.get(key);
		    int cellnum = 0;
		    for (Object obj : objArr) {
		        Cell cell = row.createCell(cellnum++);
		        if(obj instanceof Date) cell.setCellValue((Date)obj);
		        else if(obj instanceof Boolean) cell.setCellValue((Boolean)obj);
		        else if(obj instanceof String)  cell.setCellValue((String)obj);
		        else if(obj instanceof Double) cell.setCellValue((Double)obj);
		    }
		}
		 
		try {
			
		   // FileOutputStream out = new FileOutputStream(new File("new.xls"));
		   // wb.write(out);
		  //  out.close();
		    
	        ServletOutputStream out = response.getOutputStream();
	        wb.write(out);
	        out.flush();
	        out.close(); 
	        
		    System.out.println("Excel written successfully..");
		     
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}   	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    }
    
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
    	///String path = request.getPathInfo();
    	///String id = path.substring(1);
    	///System.out.println(id);
    	///System.out.println(getServletContext().getMajorVersion());
    	///String regno = "99999";
    ///	String courses1 = "ENG011,ENG012,ENG013,LIT011,LIT012,ISS011,ISS012,ISS013,ISS014,MAT011,MAT012,BED121,BED112,CHE111,CHE112,";
    	///String courses2 =  "BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,BEA311,";
    	
    	String id = request.getParameter("id");
    	
    	//System.out.println(getServletContext().getServerInfo());
    	if(getServletContext().getMajorVersion()==3) {
    		try {       	
			cn = DriverManager.getConnection("jdbc:derby://localhost:1527/agpdb", "admin", "13131");
			//PreparedStatement ps = cn.prepareStatement("SELECT * FROM ADMIN.APPLICATION WHERE UID='"+uid+"'");
			ResultSet rs = cn.createStatement().executeQuery("SELECT * FROM REG WHERE UID="+id); rs.next();        	
			id = rs.getString("id");
			uid = rs.getString("uid");
			pincode = rs.getString("pincode");
			regno = rs.getString("regno");
			password = rs.getString("password");
			mobileno = rs.getString("mobileno");
			fullname = rs.getString("fullname");
			school = rs.getString("school");
			department = rs.getString("department");
			programme = rs.getString("programme");
			session = rs.getString("session");
			courses1 = rs.getString("courses1");
			courses2 = rs.getString("courses2");
			bankname = rs.getString("bankname");
			tellerno = rs.getString("tellerno");
			amount = rs.getString("amount");
			datereg = rs.getString("datereg");
			mail = rs.getString("mail");
			created = rs.getString("created");
			filename = rs.getString("filename");		
    		} catch(Exception e1) { System.out.println(e1); }
    	}
    	else { 

			DatastoreService ds = DatastoreServiceFactory.getDatastoreService();	
			int kid = Integer.parseInt(id);
			Key ky = KeyFactory.createKey("REG", kid);	
			try {
			Entity result = ds.get(ky);			
			//id = (String) result.getProperty("id");
			uid = (String) result.getProperty("uid");
			pincode = (String) result.getProperty("pincode");
			regno = (String) result.getProperty("regno");
			password = (String) result.getProperty("password");
			mobileno = (String) result.getProperty("mobileno");
			fullname = (String) result.getProperty("fullname");
			school = (String) result.getProperty("school");
			department = (String) result.getProperty("department");
			programme = (String) result.getProperty("programme");
			session = (String) result.getProperty("session");
			courses1 = (String) result.getProperty("courses1");
			courses2 = (String) result.getProperty("courses2");
			bankname = (String) result.getProperty("bankname");
			tellerno = (String) result.getProperty("tellerno");
			amount = (String) result.getProperty("amount");
			datereg = (String) result.getProperty("datereg");
			mail = (String) result.getProperty("mail");
			created = (String) result.getProperty("created");
			filename = (String) result.getProperty("filename");
			} catch(Exception e) { System.out.println(); }
    		}
    	
    	//if(application.getMajorVersion()==3) { }
        try {
        	Document document = new Document(PageSize.A4, 15f, 15f, 15f, 15f);   
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            PdfWriter.getInstance(document, baos);
            document.open();
            
            ServletContext cntx= getServletContext();
            
            //String mime = cntx.getMimeType(filename);
            
            Image logo = Image.getInstance(cntx.getRealPath("images/logo.png"));
            logo.setAbsolutePosition(20f, 735f);
            logo.scaleAbsolute(80f, 90f);
            
            Image photo = Image.getInstance(cntx.getRealPath("images/nopic.jpg"));
            photo.setAbsolutePosition(490f, 735f);
            photo.scaleAbsolute(80f, 90f);
            
            Paragraph p1 = new Paragraph("",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 18));
            p1.add("ABDU GUSAU POLYTECHNIC");
            p1.add(Chunk.NEWLINE);
            p1.add("TALATA-MAFARA, ZAMFARA STATE");
            p1.add(Chunk.NEWLINE);
            p1.add(" ");
            p1.add("ONLINE EXAMINATION RESULTS");
            p1.setAlignment(Element.ALIGN_CENTER);
            


            
            Paragraph tt1 = new Paragraph("STUDENT INFORMATION",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10));
            Paragraph tt2 = new Paragraph("FIRST SEMESTER RESULTS",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10));
            Paragraph tt3 = new Paragraph("SECOND SEMESTER RESULTS",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10));
            Paragraph tt4 = new Paragraph("CREDITS UNIT INFORMATION",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10));
            Paragraph tt5 = new Paragraph("PROGRAMMES INFORMATION",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10));
            Paragraph tt6 = new Paragraph("RESULT INFORMATION",FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10));
            tt1.setAlignment(Element.ALIGN_CENTER);
            tt2.setAlignment(Element.ALIGN_CENTER);
            tt3.setAlignment(Element.ALIGN_CENTER);
            tt4.setAlignment(Element.ALIGN_CENTER);
            tt5.setAlignment(Element.ALIGN_CENTER);
            tt6.setAlignment(Element.ALIGN_CENTER);
            
            PdfPTable tb1 = new PdfPTable(2);
          // for(int i=1;i<=20;i++) { tb1.addCell(new PdfPCell(new Paragraph(rs.getString(i)))); }
            tb1.addCell(new PdfPCell(new Phrase("Table ID: " + id)));
            tb1.addCell(new PdfPCell(new Phrase("User ID: " + uid)));
            tb1.addCell(new PdfPCell(new Phrase("Registration No: " + regno)));
            tb1.addCell(new PdfPCell(new Phrase("Fullname: " + fullname)));
            tb1.addCell(new PdfPCell(new Phrase("School: " + school)));
            tb1.addCell(new PdfPCell(new Phrase("Department: " + department)));
            tb1.addCell(new PdfPCell(new Phrase("Programme: " + programme)));
            tb1.addCell(new PdfPCell(new Phrase("Session: " + session)));
            tb1.addCell(new PdfPCell(new Phrase("Bank Name: " + bankname)));
            tb1.addCell(new PdfPCell(new Phrase("Teller Number: " + tellerno)));
            tb1.addCell(new PdfPCell(new Phrase("Amount Paid: " + amount)));
            tb1.addCell(new PdfPCell(new Phrase("PIN Sn: " + pincode)));
            tb1.addCell(new PdfPCell(new Phrase("PIN Code: " + pincode)));
            tb1.addCell(new PdfPCell(new Phrase("Mobile Number: " + mobileno)));
            tb1.addCell(new PdfPCell(new Phrase("Date Registered: " + datereg)));
            tb1.addCell(new PdfPCell(new Phrase("Time Stamp: " + created)));        
           tb1.setWidthPercentage(100);
           tb1.setSpacingBefore(5f);
           //tb1.setSpacingAfter(5f);
           
           String[] c1 = courses1.split(",");
           String[] c2 = courses2.split(",");
           
  
           

           
           
           

           
           
           
           PdfPTable t1 = new PdfPTable(new float[] { 1, 2, 7, 1, 1, 1, 1, 1, 1, 1, 1 });
           t1.setWidthPercentage(100f);
           t1.getDefaultCell().setUseAscender(true);
           t1.getDefaultCell().setUseDescender(true);
           t1.getDefaultCell().setBackgroundColor(BaseColor.LIGHT_GRAY);
           //for (int i = 0; i < 1; i++) {
               t1.addCell("SN");
               t1.addCell("Course");
               t1.addCell("Course Title");
               t1.addCell("CU");
               t1.addCell("CA");
               t1.addCell("Exm");
               t1.addCell("Total");
               t1.addCell("Grd");
               t1.addCell("Point");
               t1.addCell("GP");
               t1.addCell("Rm");
          // }
           t1.getDefaultCell().setBackgroundColor(null);
           //t1.setHeaderRows(2);
           //t1.setFooterRows(1);

           for (int i=0;i<c1.length;i++) {
               //movie = screening.getMovie();
               t1.addCell(""+(i+1));
               t1.addCell(c1[i]);
               t1.addCell(course(c1[i],"title"));
               t1.addCell(course(c1[i],"cu"));
               t1.addCell(course(c1[i],"status"));
               t1.addCell(course(c1[i],"status"));
               t1.addCell(course(c1[i],"status"));
               t1.addCell(course(c1[i],"status"));
               t1.addCell(course(c1[i],"status"));
               t1.addCell(course(c1[i],"status"));
               t1.addCell(course(c1[i],"status"));
             //  table.addCell(String.valueOf(movie.getYear()));
           }
 
           PdfPTable t2 = new PdfPTable(new float[] { 1, 2, 7, 1, 1, 1, 1, 1, 1, 1, 1 });
           t2.setWidthPercentage(100f);
           t2.getDefaultCell().setUseAscender(true);
           t2.getDefaultCell().setUseDescender(true);
           t2.getDefaultCell().setBackgroundColor(BaseColor.LIGHT_GRAY);
          // for (int i = 0; i < 2; i++) {
           t2.addCell("SN");
           t2.addCell("Course");
           t2.addCell("Course Title");
           t2.addCell("CU");
           t2.addCell("CA");
           t2.addCell("Exm");
           t2.addCell("Total");
           t2.addCell("Grd");
           t2.addCell("Point");
           t2.addCell("GP");
           t2.addCell("Rm");
         //  }
          t2.getDefaultCell().setBackgroundColor(null);
         //  t2.setHeaderRows(2);
         //  t2.setFooterRows(1);

           for (int i=0;i<c2.length;i++) {
               //movie = screening.getMovie();
               t2.addCell(""+(i+1));
               t2.addCell(c2[i]);
               t2.addCell(course(c2[i],"title"));
               t2.addCell(course(c2[i],"cu"));
               t2.addCell(course(c1[i],"status"));
               t2.addCell(course(c1[i],"status"));
               t2.addCell(course(c1[i],"status"));
               t2.addCell(course(c1[i],"status"));
               t2.addCell(course(c1[i],"status"));
               t2.addCell(course(c1[i],"status"));
               t2.addCell(course(c1[i],"status"));             //  table.addCell(String.valueOf(movie.getYear()));
           }
           
           t1.setSpacingBefore(3f);
           
           t2.setSpacingBefore(3f);

           

           PdfPTable tb4 = new PdfPTable(2);
         // for(int i=1;i<=20;i++) { tb1.addCell(new PdfPCell(new Paragraph(rs.getString(i)))); }
           tb4.addCell(new PdfPCell(new Phrase("First Semester Total Credits: " + "")));
           tb4.addCell(new PdfPCell(new Phrase("Second Semester Total Credits: " + "")));
           tb4.addCell(new PdfPCell(new Phrase("First Semester Total Points: " + "")));
           tb4.addCell(new PdfPCell(new Phrase("Second Semester Total Points: " + "")));
           tb4.addCell(new PdfPCell(new Phrase("First Semester Total Grade Points: " + programme)));
           tb4.addCell(new PdfPCell(new Phrase("Second Semester Total Grade Points: " + session)));
           tb4.addCell(new PdfPCell(new Phrase("First Semester GPA: " + amount)));
           tb4.addCell(new PdfPCell(new Phrase("Second Semester GPA: " + "")));
           tb4.addCell(new PdfPCell(new Phrase("Current GPA: " + pincode)));
           tb4.addCell(new PdfPCell(new Phrase("Remarks: " + mobileno)));    
           tb4.setWidthPercentage(100);
           tb4.setSpacingBefore(5f);
           
           document.add(p1);
           document.add(logo);  
           document.add(photo);
           document.add(tt1);         
           document.add(tb1);
           document.add(tt2);
           document.add(t1);
           document.add(tt3);
           document.add(t2);
           document.add(tt6);
           document.add(tb4);
            
        	/**       

         
            if (uid == null || uid.trim().length() == 0) { uid = "Invalid ID"; }
            Document document = new Document(PageSize.A4, 20f, 20f, 20f, 20f);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            PdfWriter.getInstance(document, baos);
            document.open();

//            document.add(new Paragraph(String.format("23456You have submitted the following text using the %s method:",request.getMethod())));
//            document.add(new Paragraph(uid));                        

            

           // document.add(Chunk.NEWLINE);            
  

            
            
            

          




*/
           
           
document.add(new Phrase("ABDU GUSAU POLYTECHNIC TALATA MAFARA, ZAMFARA STATE [Cloud System Developed by: MANNIR ESYSTEMS LIMITED 2011-2013 (www.mannir.net)]", FontFactory.getFont(FontFactory.TIMES, 8)));           
            document.close();            
            response.setHeader("Expires", "0");
            response.setHeader("Cache-Control","must-revalidate, post-check=0, pre-check=0");
            response.setHeader("Pragma", "public");
            response.setContentType("application/pdf");
            response.setContentLength(baos.size());
            OutputStream os = response.getOutputStream();
            baos.writeTo(os);
            os.flush();
            os.close();
        }
        catch(Exception e3) { System.out.println(e3); }
       // }




    }

    public String gd(String type, int cu, int ca, int ex) {
    	String rt = null;
    	int tt = (ca+ex);
    	String gr = null;
    	double pt;
    	double gp;
    	String rm;  	
    	
    	int marks = tt;

        	if (marks >= 70) {gr="A"; pt =4.0; gp = pt*cu; rm = "DISTINCTION";} 
        else if (marks >= 60) {gr="B"; pt =3.0; gp = pt*cu; rm = "CREDIT";} 
        else if (marks >= 50) {gr="C"; pt =3.0; gp = pt*cu; rm = "VERY GOOD";} 
        else if (marks >= 45) {gr="D"; pt =2.0; gp = pt*cu; rm = "GOOD";} 
        else if (marks >= 40) {gr="E"; pt =1.0; gp = pt*cu; rm = "PASS";} 
        				else  {gr="F"; pt =0.0; gp = pt*cu; rm = "C/O"; }

    	
    	

    	switch(type) {
		case "tt":  rt = tt+""; break;
		case "gr": rt = gr; break;			
		case "pt": rt = pt+""; break;
		case "gp": rt = gp+""; break;
		case "rm": rt = rm; break;
		
    	}

    	return rt;
    }
    
    public String course(String cd, String tp) {
    	String vl = null;
    	try {
    	ResultSet r = cn.createStatement().executeQuery("SELECT * FROM COURSES WHERE CODE='"+cd+"'"); r.next();
    	vl = r.getString(tp);
    	} catch (Exception e3) {System.out.println(e3); }
    	return vl;
    }
}
