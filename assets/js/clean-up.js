var Excel = function() {

	this.structure =
	{
		"doctype": '<?xml version="1.0"?>',
		"workbook":
		{
			"open": [
					'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"',
					'xmlns:o="urn:schemas-microsoft-com:office:office"',
					'xmlns:x="urn:schemas-microsoft-com:office:excel"',
					'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"',
					'xmlns:html="http://www.w3.org/TR/REC-html40">'
				].join('\n'),
			"close": '</Workbook>'
		},
		"OfficeDocumentSettings":
		{
			"open": '<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">',
			"allowPNG": '<AllowPNG/>',
			"close": '</OfficeDocumentSettings>'
		},
		"DocumentProperties":
		{
			"open": '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">',
			"close": '</DocumentProperties>'
		},
		"worksheet":
		{
			"open": [
					'<Worksheet ss:Name="Sheet1">',
					'<Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="%EXPANDEDROWCOUNT%" x:FullColumns="1" x:FullRows="1">'
				].join('\n'),
			"close": [
						'</Worksheet>',
						'</Table>'
				].join('\n')
		},
		"row":
		{
			"open":'<Row>',
			"close":'</Row>'
		},
		"cellTypes": ['string', 'number', 'date'],
		"cell":
		{
			"open": function(styleID, type){
				if (styleID !== undefined && type !== undefined){
					return '<Cell ss:StyleID="'+styleID+'"><Data ss:Type="'+type+'">';
				} else if (styleID !== undefined && type === undefined) {
					return '<Cell ss:StyleID="'+styleID+'"><Data ss:Type="String">';
				} else {
					return '<Cell><Data ss:Type="String">';
				}
				return this.open();
			},
			"close":[
					'</Cell>',
					'</Data>'
				].join('\n')

		}

	};
	this.styles =
	{
		"open": '<Styles>',
		"close": '</Styles>'
	};
};
function handleFileSelect(event) {
	var files = event.target.files; // FileList object
	var dataTable;
	for (var i = 0, z = 1, n = files.length; i <= n; i++, z++) {
		var reader = new FileReader();
	}
}