<%@include file="/mannir/mannir.jsp" %>
<%@include file="/header.jsp" %>
<%@include file="/menus.jsp" %>
<h1 class="title" id="page-title" align="center">Examination Records</h1>
<a href="<%=this.getServletContext().getContextPath() %>/exams.xls">Sample Excel File</a>

<form method="POST" enctype="multipart/form-data" action="/exams">
  <fieldset class="container-inline collapsible collapsed form-wrapper"><legend><span class="fieldset-legend">Upload Examination Records</span></legend>
<div class="fieldset-wrapper"><div style="float:left;margin-right:2px;">
<div style="float:left;margin-right:2px;"><label>Select File to Upload</label><input type="file" name="upfile"></div>
<div style="float:left;margin-right:2px;"><label>Type Course Code</label><input type="text" name="note"></div>
<input type="submit" value="Upload" class="form-submit">
  </div></div>
</fieldset>
</form>



<table class="sticky-enabled" style="width:80%">
<thead><tr><th>SN</th><th>Reg No</th><th>Fullname</th><th>Course</th><th>CU</th><th>CA</th><th>Exam</th><th>Total</th><th>Grade</th><th>Point</th><th>GP</th><th>Remarks</th><th>View Result</th><th>Print Result</th></tr></thead>
<tbody>

<%

if(application.getMajorVersion()==3) {
  
ResultSet rs = cn.createStatement().executeQuery("SELECT * FROM EXAMS"); //rs.next();	
int sn = 1;	
while(rs.next()){
	int m = sn % 2;
	if(m==1) { out.println("<tr class='odd'>"); }
	else { out.println("<tr class='even'>"); }
	
	out.println("<td style='border:1px solid black;'>"+sn+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("regno")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("code")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("ca")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("exam")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("total")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("grade")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("point")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("gp")+"</td>");
	out.println("<td style='border:1px solid black;'>"+rs.getString("remarks")+"</td>");
	out.println("<td style='border:1px solid black;'><a href='/application/"+rs.getString("regno")+"' target='new'>Print</a></td>");
	out.println("</tr>");
	sn++;
	}			
}

	else {
		DatastoreService datastore = DatastoreServiceFactory.getDatastoreService();
	    Query query = new Query("EXAMS"); //, guestbookKey).addSort("date", Query.SortDirection.DESCENDING);
	    List<Entity> el = datastore.prepare(query).asList(FetchOptions.Builder.withLimit(50));
	    if (el.isEmpty()) { out.println("No Records!"); }
	    
	    int sn = 1;	
		for (Entity en : el) {
			int m = sn % 2; if(m==1) { out.println("<tr class='odd'>"); } else { out.println("<tr class='even'>"); }
	        out.println("<td style='border:1px solid black;'>"+en.getKey().getId()+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("regno")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("fullname")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("course")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("cu")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("ca")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("exam")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("total")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("grade")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("point")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("gp")+"</td>");
	        out.println("<td style='border:1px solid black;'>"+en.getProperty("remarks")+"</td>");
			out.println("<td style='border:1px solid black;'><a href='/viewresult.jsp?pin="+en.getProperty("regno")+"'>View</a></td>");
			out.println("<td style='border:1px solid black;'><a href='/exams?id="+en.getProperty("regno")+"' target='new'>Print</a></td>");
	        out.println("</tr>");
	        sn++;
	        }
		
	    
	}

%>



</tbody>
</table>

<%@include file="footer.jsp" %>
