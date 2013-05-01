
var excel = {};
	function handleFileSelect(evt) {
		var files = evt.target.files; // FileList object
		var dataTable;
			// Loop through the FileList
			for (var i = 0, z = 1, f; f = files[i]; i++) {
			  fileCount = (i+1);
			  var reader = new FileReader();

			  // Closure to capture the file information.
			  reader.onload = (function(theFile) {
				return function(e) {
				  // Print the contents of the file
				  var span = document.createElement('span');
				  //console.log(e);
				  var tester = $(e.target.result).find("table");
				  var len = tester.length;
				  
				  for (var j = 0, n = tester.length; j < n; j++ ) {
					  if (tester[j+1] != null) {
						  if ( $( tester[j] ).html().length > $( tester[j+1] ).html().length ) {
							  dataTable = $(tester[j]);
						  } else {
							  dataTable = $( tester[j+1] );
						  }
					  }
				  }
				  
				 excel.rows = dataTable.find('tr').length - 1;
				 excel.sheet = "<?xml version=\"1.0\"?> \n\
<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" \n\
 xmlns:o=\"urn:schemas-microsoft-com:office:office\" \n\
 xmlns:x=\"urn:schemas-microsoft-com:office:excel\" \n\
 xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" \n\
 xmlns:html=\"http:\/\/www.w3.org\/TR\/REC-html40\"> \n\
 <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\"> \n\
  <Version>12.0<\/Version> \n\
 <\/DocumentProperties> \n\
 <OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\"> \n\
  <AllowPNG\/> \n\
 <\/OfficeDocumentSettings> \n\
 <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\"> \n\
  <WindowHeight>18480<\/WindowHeight> \n\
  <WindowWidth>30320<\/WindowWidth> \n\
  <WindowTopX>180<\/WindowTopX> \n\
  <WindowTopY>160<\/WindowTopY> \n\
  <ProtectStructure>False<\/ProtectStructure> \n\
  <ProtectWindows>False<\/ProtectWindows> \n\
 <\/ExcelWorkbook> \n\
 <Styles> \n\
  <Style ss:ID=\"Default\" ss:Name=\"Normal\"> \n\
   <Alignment ss:Vertical=\"Bottom\"\/> \n\
   <Borders\/> \n\
   <Font ss:FontName=\"Verdana\"\/> \n\
   <Interior\/> \n\
   <NumberFormat\/> \n\
   <Protection\/> \n\
  <\/Style> \n\
  <Style ss:ID=\"m310936552\"> \n\
   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"\/> \n\
   <Borders> \n\
	<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
   <\/Borders> \n\
   <Font ss:FontName=\"Verdana\" ss:Size=\"18.0\" ss:Bold=\"1\"\/> \n\
  <\/Style> \n\
  <Style ss:ID=\"m310936572\"> \n\
   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"\/> \n\
   <Borders> \n\
	<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
   <\/Borders> \n\
   <Font ss:FontName=\"Verdana\" ss:Size=\"14.0\"\/> \n\
  <\/Style> \n\
  <Style ss:ID=\"m310936592\"> \n\
   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"\/> \n\
   <Borders> \n\
	<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
   <\/Borders> \n\
   <Font ss:FontName=\"Verdana\" ss:Size=\"14.0\" ss:Bold=\"1\"\/> \n\
  <\/Style> \n\
  <Style ss:ID=\"s21\">  \n\
   <Borders> \n\
	<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
	<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" \
	 ss:Color=\"#C0C0C0\"\/> \n\
   <\/Borders> \n\
  <\/Style> \n\
  <Style ss:ID=\"s37\"> \n\
   <Borders\/> \n\
  <\/Style> \n\
  <Style ss:ID=\"s38\"> \n\
   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"\/> \n\
   <Borders\/> \n\
   <Font ss:FontName=\"Verdana\" ss:Size=\"14.0\" ss:Bold=\"1\"\/> \n\
  <\/Style> \n\
 <\/Styles> \n\
 <Worksheet ss:Name=\"Sheet1\"> \n\
  <Table ss:ExpandedColumnCount=\"10\" ss:ExpandedRowCount=\"%EXPANDEDROWCOUNT%\" x:FullColumns=\"1\" \
   x:FullRows=\"1\"> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"165.0\"\/> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"98.0\" ss:Span=\"1\"\/> \n\
   <Column ss:Index=\"4\" ss:AutoFitWidth=\"0\" ss:Width=\"103.0\"\/> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"113.0\"\/> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"96.0\"\/> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"128.0\"\/> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"93.0\"\/> \n\
   <Column ss:AutoFitWidth=\"0\" ss:Width=\"94.0\"\/> \n\
   \n\
   <Row ss:AutoFitHeight=\"0\" ss:Height=\"23.0\" ss:StyleID=\"s21\">  \n\
	<Cell ss:MergeAcross=\"9\" ss:StyleID=\"m310936552\"><Data \  ss:Type=\"String\">%TITLE_GOES_HERE%<\/Data><\/Cell> \n\
   <\/Row> \n\
   <Row ss:AutoFitHeight=\"0\" ss:Height=\"18.0\" ss:StyleID=\"s21\"> \n\
	<Cell ss:MergeAcross=\"9\" ss:StyleID=\"m310936572\"><Data ss:Type=\"String\">REPORT LIMITED BY \ READER<\/Data><\/Cell>\n\
   <\/Row> \n\
   <Row ss:AutoFitHeight=\"0\" ss:Height=\"18.0\" ss:StyleID=\"s21\"> \n\
	<Cell ss:MergeAcross=\"9\" ss:StyleID=\"m310936592\"><Data ss:Type=\"String\">%X_NUMBER% entries were found \
	matching your search criteria.<\/Data><\/Cell> \n\
   </Row>\n";
				  
				  
				  excel.prefix = $("#excelPrefix").val();
				  excel.sheet += "<Row ss:AutoFitHeight=\"0\" ss:Height=\"18.0\" ss:StyleID=\"s37\">\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<Cell ss:StyleID=\"s38\"\/>\n\
									<\/Row>\n";
				  excel.sheet += dataTable.html().replace(/<th>/g,"<Cell ss:StyleID=\"s38\"><Data ss:Type=\"String\">");
				  excel.title = $("#excelTitle").val();
				  console.log("File number: " + z);
				  if (excel.title == "") {
					  if ( excel.prefix !== "" ) {
						  excel.title = excel.prefix + "_" + z;
						  z++;
					  } else {
						  excel.title = "Spreadsheet_" + z;
						  z++;
					  }
				  }
				  excel.sheet = excel.sheet.replace(/%TITLE_GOES_HERE%/g, excel.title);
				  excel.sheet = excel.sheet.replace(/%X_NUMBER%/g, excel.rows);
				  excel.sheet = excel.sheet.replace(/%EXPANDEDROWCOUNT%/, excel.rows + 50); //50 extra for wiggle room
				  excel.sheet = excel.sheet.replace(/<\/th>/g,"<\/Data><\/Cell>\t");
				  excel.sheet = excel.sheet.replace(/<tbody>/g,"");
				  excel.sheet = excel.sheet.replace(/<\/tbody>/g, "");
				  excel.sheet = excel.sheet.replace(/<tr>/g, "<Row>");
				  excel.sheet = excel.sheet.replace(/<\/tr>/g, "<\/Row>");
				  excel.sheet = excel.sheet.replace(/<td>/g, "<Cell><Data ss:Type=\"String\">");
				  excel.sheet = excel.sheet.replace(/<\/td>/g, "<\/Data><\/Cell>");

				  excel.sheet += "<\/Table> \n\
								 <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n\
								 <PageLayoutZoom>0<\/PageLayoutZoom>\n\
								 <Selected\/>\n\
								 <Panes>\n\
								 <Pane>\n\
								 <Number>3<\/Number>\n\
								 <ActiveRow>8<\/ActiveRow>\n\
								 <ActiveCol>9<\/ActiveCol>\n\
								 <\/Pane>\n\
								 <\/Panes>\n\
								 <ProtectObjects>False<\/ProtectObjects>\n\
								 <ProtectScenarios>False<\/ProtectScenarios>\n\
								 <\/WorksheetOptions>\n\
								 <\/Worksheet>\n\
								 <\/Workbook>\n";
				  console.log(encodeURIComponent(excel.sheet));
				  uriContent = "data:application/vnd.ms-excel;charset=utf-8," + encodeURIComponent(excel.sheet);
				  $("#downloadHere").append("<a href=\""+uriContent+"\" target=\"_blank\" download=\""+excel.title+".xls\" id=\""+excel.title+"\">Download "+excel.title+" now<\/a><br>\n");
				  //Append for later version <input type=\"text\" placeholder=\"Name of File Here\" data-bind=\""+excel.title+"\" id=\"bind_"+excel.title+"\"> 
				  
				  $("#list").append($(span).html(e.target.result));
				  $("#excelTitle").val("");
				};
			  })(f);
			  
			  reader.readAsText(f, "UTF-8");
			}
		  }
		  /*
		  $(document).ready(function(){
			  $("#downloadHere").on("change", "[type='text']", function(event) {
				  alert("changed");
			  });
		  });
		  */
		  document.getElementById('files').addEventListener('change', handleFileSelect, false);