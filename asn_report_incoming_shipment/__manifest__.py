{
	"name"			: "Report Incoming Shipment",
	"version"		: "10.0",
	"author"		: "ajeng039e@gmail.com",
	"category"		: "Report",
	'website'		: '',
	"description"	: """\
       					 This module provide report for incoming shipment with lot number
    				  """,
	"summary"		: "Report incoming shipment with Lot Number",
    "depends"       : ['xlsx_report_template','stock'],
	"data"			: ['wizard/report_incoming_shipment_view.xml'],
	"installable"	: True,
	"auto_install"	: False,
    "application"	: True,
}
