# controlPCISeriesAfromExcel
Control of the program Personal Communications I Series Access for windows through macro programmed in excel with VBA.

Prueba de Ejecución:

![Imgur](https://i.imgur.com/Ip28SCC.gif)


Prerrequisitos (solo para windows 7 - 10):
Office 2012-2019 (32 bits)
Personal Communications iSeries Access para Windows

Instrucciones:
* Abrir o crear un archivo Excel habilitado para macros.
* Cree una tabla de contenidos a una hoja especifica llamada VAR

![Imgur1](https://i.imgur.com/w8SWzkm.png)

* Importe el modulo .bas

![Imgur2](https://i.imgur.com/doXrknC.png)


* En otra Hoja (puede ser Hoja1) construya en una hoja vacia la siguiente tabla poniendo especial atencion a las columnas especificadas en la hoja VAR en el paso anterior las columnas deben concordar con los encabezados, no textualmente pero si deben ser los datos que se especificaron el la hoja VAR.


![Imgur3](https://i.imgur.com/rhakXs7.png)


* Los datos objeto de busqueda son los codigos, estos se toman como referencia para ubicar el resto de datos en el sistema en base a una logica especifica de pulsacion de teclas y obtencion de datos.

* Abrir el programa Personal Communications iSeries Access para Windows y loguearse, desplazarse hasta la busqueda de informacion de clientes en base al codigo (depende del programa).

![Imgur4](https://i.imgur.com/JS9F7k8.png)

* Digitar codigos a buscar, seleccionar los codigos en la tabla y ejecutar la macro.

![Imgur](https://i.imgur.com/Ip28SCC.gif)

nota:
La seleccion puede ser uno o varios elementos y soporta tambien elementos solo de un filtro especificado (previamente se deben filtrar los datos de la tabla en excel y unicamnete ejecutara la macro a la seleccion sin considerar filas ocultas).


Bibliografia:
https://www.ibm.com/docs/es/personal-communications/12.0?topic=sseq5y-12-0-0-com-ibm-pcomm-doc-books-html-host-access08-htm
