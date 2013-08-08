// parses the excel files
var htmlFiles = [];
var TextParser = function TextParser() {
	var _this = this;

	// private methods
	// this is a queue of files that are queued to be parsed
	this._fileQueue = [];

	this._parseFile = function(file) {
		var reader = new FileReader(),
			rFilter = /^(?:text\/html|text\/plain)$/i,
			defer = jQuery.Deferred(),
			textFile = {};

		// throw an error if it is not an instance of File
		if( !(file instanceof File) ) {
			alert('You must give me a file!');
			throw new Error('There was a problem and the file was not read correctly');
		}
		// throw an error if it isn't html or plain text
		if( !rFilter.test(file.type) ) {
			alert('You must select a valid html file!');
			throw new Error('must select a valid html file');
		}

		// pass off name, size and type
		textFile.fileName = file.name;
		textFile.size = file.size;
		textFile.type = file.type;
		// change to more MV*-esque
		textFile.updatedAt = file.lastModifiedDate;

		// resolve the promise when the file is finished being read
		reader.onload = function() {
			textFile.contents = reader.result;
			console.log('textFile ready');
			defer.resolve(textFile);
		};

		// reject promise if something goes wrong
		reader.onerror = function() {
			defer.reject("Something went wrong while trying to read the File.");
		};

		// notify the promise that it is still reading the file every second
		reader.onprogress = function working() {
			if ( defer.state() === "pending" ) {
				defer.notify( "Reading file..." );
				setTimeout( working, 1000 );
			}
		};

		// read the file
		reader.readAsText(file);

		// return a promise to parse file
		return defer.promise();
	};

	// public methods
	this.textFiles = [];

	// handle the promise that _parseFile returns, and push it into a done queue
	this.parse = function(file) {
		jQuery.when( _this._parseFile(file) ).then(
			// success
			function( data ) {
				_this.textFiles.push(data);
			},
			// error
			function( error ){
				throw new Error(error);
			},
			// in progress
			function( status ) {
				console.log( status );
			}
		);
	};

	this.queueUp = function(filesArray) {
		if( !(filesArray instanceof Array) ) {
			filesArray = [filesArray];
		}
		while( filesArray.length > 0 ) {
			_this._fileQueue.push( filesArray.shift() );
		}
	};

	this.parseAll = function() {
		while( _this._fileQueue.length > 0 ) {
			_this.parse( _this._fileQueue.shift() );
		}
	};

};

// view controller
var queueController = function(selector) {
	var _this = this,
		parser = new TextParser();

	this.el = document.querySelectorAll(selector);
	this.$el = $(selector);

	this.queueFiles = function() {

	};

	this.renderQueue = function() {

	};
};

//document.getElementById('drop-files').addEventListener('change', parser.queueUp(), false);

var Excelerator = function Excelerator() {
	var excelSheet = {};

	if(!window.File && !window.FileList && !window.FileReader) {
		alert('This Browser is incapable of using this program. Please use the latest version of Mozilla Firefox or Google Chrome.');
		throw new Error('This Browser is incapable of using this program. Please use the latest version of Mozilla Firefox or Google Chrome.');
	}

	// read file on initialize

	// call generateXLSX to spit out XLSX format

	this.generateXLSX = function generateXLSX() {
		var excelFile = excelStructure.docHeader + excelStructure.styles + excelStructure.tableHeader + excelStructure.spacerRow;

	};

};

excelStructure = {
	docHeader: '<?xml version="1.0"?>' +
		'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' +
			'xmlns:o="urn:schemas-microsoft-com:office:office"' +
			'xmlns:x="urn:schemas-microsoft-com:office:excel"' +
			'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' +
			'xmlns:html="http://www.w3.org/TR/REC-html40">' +
		'<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">' +
			'<Version>12.0</Version>' +
		'</DocumentProperties>' +
		'<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office"> ' +
			'<AllowPNG/> ' +
		'</OfficeDocumentSettings> ' +
		'<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel"> ' +
			'<WindowHeight>18480</WindowHeight> ' +
			'<WindowWidth>30320</WindowWidth> ' +
			'<WindowTopX>180</WindowTopX> ' +
			'<WindowTopY>160</WindowTopY> ' +
			'<ProtectStructure>False</ProtectStructure> ' +
			'<ProtectWindows>False</ProtectWindows> ' +
		'</ExcelWorkbook>',

	styles: '<Styles>' +
		'<Style ss:ID="Default" ss:Name="Normal">' +
			'<Alignment ss:Vertical="Bottom"/>' +
			'<Borders/>' +
			'<Font ss:FontName="Verdana"/>' +
			'<Interior/>' +
			'<NumberFormat/>' +
			'<Protection/>' +
		'</Style>' +
		'<Style ss:ID="m310936552">' +
			'<Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>' +
			'<Borders> ' +
				'<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"' +
					' ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"' +
					' ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"' +
					' ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"' +
					' ss:Color="#C0C0C0"/> ' +
			'</Borders>' +
			'<Font ss:FontName="Verdana" ss:Size="18.0" ss:Bold="1"/>' +
		'</Style>' +
		'<Style ss:ID="m310936572">' +
			'<Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/> ' +
			'<Borders> ' +
				'<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
			'</Borders> ' +
			'<Font ss:FontName="Verdana" ss:Size="14.0"/> ' +
		'</Style> ' +
		'<Style ss:ID="m310936592"> ' +
			'<Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/> ' +
			'<Borders> ' +
				'<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
			'</Borders> ' +
			'<Font ss:FontName="Verdana" ss:Size="14.0" ss:Bold="1"/>' +
		'</Style>' +
		'<Style ss:ID="s21">' +
			'<Borders>' +
				'<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
				'<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"' +
					'ss:Color="#C0C0C0"/> ' +
			'</Borders>' +
		'</Style>' +
		'<Style ss:ID="s37">' +
			'<Borders/> ' +
		'</Style>' +
		'<Style ss:ID="s38"> ' +
			'<Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>' +
			'<Borders/>' +
			'<Font ss:FontName="Verdana" ss:Size="14.0" ss:Bold="1"/>' +
		'</Style>' +
	'</Styles>',

	tableHeader: '<Worksheet ss:Name="Sheet1">' +
		'<Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="%EXPANDEDROWCOUNT%" x:FullColumns="1"' +
			'x:FullRows="1"> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="165.0"/> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="98.0" ss:Span="1"/> ' +
			'<Column ss:Index="4" ss:AutoFitWidth="0" ss:Width="103.0"/> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="113.0"/> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="96.0"/> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="128.0"/> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="93.0"/> ' +
			'<Column ss:AutoFitWidth="0" ss:Width="94.0"/> ' +
			'<Row ss:AutoFitHeight="0" ss:Height="23.0" ss:StyleID="s21">  ' +
				'<Cell ss:MergeAcross="9" ss:StyleID="m310936552"><Data ss:Type="String">%TITLE_GOES_HERE%</Data></Cell> ' +
			'</Row>' +
			'<Row ss:AutoFitHeight="0" ss:Height="18.0" ss:StyleID="s21"> ' +
				'<Cell ss:MergeAcross="9" ss:StyleID="m310936572"><Data ss:Type="String">REPORT LIMITED BY READER</Data></Cell>' +
			'</Row>' +
			'<Row ss:AutoFitHeight="0" ss:Height="18.0" ss:StyleID="s21"> ' +
				'<Cell ss:MergeAcross="9" ss:StyleID="m310936592"><Data ss:Type="String">%X_NUMBER% entries were found' +
					'matching your search criteria.</Data></Cell> ' +
			'</Row>',
	spacerRow: '<Row ss:AutoFitHeight="0" ss:Height="18.0" ss:StyleID="s37">' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
		'<Cell ss:StyleID="s38"/>' +
	'</Row>',
	docFooter: '</Table>' +
		'<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">' +
			'<PageLayoutZoom>0</PageLayoutZoom>' +
			'<Selected/>' +
			'<Panes>' +
				'<Pane>' +
					'<Number>3</Number>' +
					'<ActiveRow>8</ActiveRow>' +
					'<ActiveCol>9</ActiveCol>' +
				'</Pane>' +
			'</Panes>' +
			'<ProtectObjects>False</ProtectObjects>' +
			'<ProtectScenarios>False</ProtectScenarios>' +
		'</WorksheetOptions>' +
	'</Worksheet>' +
	'</Workbook>'

};
