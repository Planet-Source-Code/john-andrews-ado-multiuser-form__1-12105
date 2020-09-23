Instructions for the ADO Data Form demo.

DESCRIPTION
This Visual Basic 6 project contains a form for browsing an ADO recordset. It is based on 
the 'single-record form' created by the VB wizard, but includes:

* Multiuser capabilities (with a client-side recordset)
* Enhanced error handling
* Improved user interface

The demo allows several instances of the form to be opened, simulating a 'pretend' multiuser
network.

The form is intended for use as a template, and as a learning tool.

USAGE
To use the form in your own project:
1: Add a copy of the form to your project
2: Set the ADO connection string in the form_load code to match your database
3: Place databound controls on the form (if you drop controls onto the form from the
data environment, set their datasource and datamember properties to blank)
4: You may need to adjust your Project/References if you are using a different version of ADO to v2.5, or Data Binding Collection older than SP4 

Hope you find the form useful! - john.andrews@angelfire.com