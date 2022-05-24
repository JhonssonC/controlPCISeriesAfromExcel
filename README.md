# controlPCISeriesAfromExcel
Control of the program Personal Communications I Series Access for windows through macro programmed in excel with VBA.


Execution Test:

[Imgur](https://i.imgur.com/Ip28SCC.gifv)


Prerequisites (only for windows 7 - 10):
Office 2012-2019 (32-bit)
Personal Communications iSeries Access for Windows

Instructions:
* Open or create a macro-enabled Excel file.
* Create a table of contents to a specific sheet called VAR

[Imgur1](https://i.imgur.com/w8SWzkm.png)

* Import the .bas module

![Imgur2](https://i.imgur.com/doXrknC.png)

* In another Sheet (it can be Sheet1) build the following table in an empty sheet, paying special attention to the columns specified in the VAR sheet in the previous step, the columns must match the headers, not textually, but they must be the data that is specified on the VAR sheet.

[Imgur3](https://i.imgur.com/rhakXs7.png)

* The data object of search are the codes, these are taken as a reference to locate the rest of the data in the system based on a specific logic of pressing keys and obtaining data.

* Open the Personal Communications iSeries Access for Windows program and log in, navigate to the search for customer information based on the code (depends on the program).

[Imgur4](https://i.imgur.com/JS9F7k8.png)

* Enter codes to search, select the codes in the table and execute the macro.

[Imgur](https://i.imgur.com/Ip28SCC.gifv)

note:
The selection can be one or several elements and it also supports elements only from a specified filter (previously the table data must be filtered in excel and it will only execute the macro on the selection without considering hidden rows).


Bibliography:
https://www.ibm.com/docs/en/personal-communications/12.0?topic=sseq5y-12-0-0-com-ibm-pcomm-doc-books-html-host-access08-htm