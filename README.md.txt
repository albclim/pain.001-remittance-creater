PAIN.001 Generation Macro
This repository contains an Excel-based tool (PAIN001 generationv2.xlsm) designed to automate the creation of ISO 20022 compliant payment initiation messages.

Project Overview
The macro leverages VBA (Visual Basic for Applications) to transform structured spreadsheet data into PAIN.001 XML files. This format is the global standard for initiating credit transfers with financial institutions.

Technical Components
	VBA Engine: Contained within xl/vbaProject.bin, the code handles data parsing, XML node construction, and file system operations.
	Data Structure: The workbook utilizes four primary worksheets to manage payment instructions, party identifiers, and bank metadata.
	Form Controls: Includes embedded ActiveX/Form controls (defined in ctrlProps) to provide a user interface for triggering the generation process.

Key Features
	Automated XML Schema Mapping: Maps Excel cell ranges to specific ISO 20022 tags (e.g., <GrpHdr>, <PmtInf>, <CdtTrfTxInf>).
	Validation: Typically includes logic to ensure mandatory fields like IBAN, BIC, and Currency codes meet formatting requirements before export.
	Batch Processing: Capable of generating multiple payment instructions within a single XML message container.