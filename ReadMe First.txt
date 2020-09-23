Stock Controlling System - Initial Instructions
-----------------------------------------------

Note:
 Stock Controlling System needs few initial configurations before the use of it as I developed this programme
 to install on computer with a setup programme. Here you will get only the source Code of Stock Controlling System.
 Therefore, please follow the below instructions for a proper execution of the programme.

1. As Stock Controlling System uses custom-created ActiveX dynamic-link library (DLL) file for the Clear routine of it's database path configuration,
    first you go to 'Db Clear Config Dll' folder and make the project as 'db_cls_config.dll' there. 
    (From the File menu --> Make db_cls_config.dll).

    * db_cls_config.dll file is used with Stock Controlling System because Clear routine of the database path configuration is also
      performed outside the main executable file, which is 'STOCK.exe' by using the Clsdbconfig.exe, from which you may create its executable
      file from 'Db Clear Config Exe' folder. 

2. Make a reference to db_cls_config.dll you created.

3. See the HELP SCS.htm file in 'Stock Programme - to be published' folder to get an extensive idea of the programme.

		=======================================================================
		= Initial password for Administrator: password 			      =
		= Password of databases: king@#$%^12sam2009                           =
	        = Backdoor of login to the system (command line argument): bypassadmn =		
		=======================================================================

4. Some features of this program has been developed to be used in a multi-user environment. But, if you need to test the program on a single PC just use 	      separate executable files of the Main Execurable file (STOCK.exe).

   * Even though this program uses Microsoft Access, which is not a multi-user DBMS, as the backend, you can use Microsoft SQL Server instead with some 	     modifications to the program.

I developed this programme nearly 2 years back for Singapore Informatics Computer Institute(Pvt) Ltd. I think this programme will certainly lead you to think
beyond the traditional database programming. Please send me your comments to pgbsoft@gmail.com. I really appreciate them.

----------------------------------------------------------------------------------------------------------------------------------------------------------

P. G. B. Prasanna
pgbsoft@gmail.com
