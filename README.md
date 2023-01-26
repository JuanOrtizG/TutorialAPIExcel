# TutorialAPIExcel

Tutoriales para Excel
Para interactuar con la API de Excel se dispone de dos maneras: 
API de JavaScript de Excel: Presentada con Office 2016, la API de JavaScript de Excel proporciona objetos fuertemente tipados que puede usar para acceder a hojas de cálculo, rangos, tablas, gráficos y más.

API common: introducida con Office 2013, la API common se puede usar para acceder a funciones como la interfaz de usuario, los cuadros de diálogo y la configuración del cliente que son comunes en varios tipos de aplicaciones de Office.

![image](https://user-images.githubusercontent.com/12565341/214921918-58bd1df5-98c1-4948-8470-82f76d0e28e3.png)

# Modelo de objetos de Excel
Si imaginamos una casa y queremos utilizar un cuchillo, no podemos utilizarlo directamente, debemos ingresar a la casa, ingresar a la cocina, luego ingresar al armario, y abrir el cajón de cubiertos, y ahí obtenemos el cuchillo, no hay forma de obtenerlo directamente. 
en Excel hacemos lo mismo, debemos ingresar primero a la aplicación que será el libro de trabajo, luego debemos seleccionar una hoja, luego debemos elegir un rango y posteriormente podemos acceder a  una celda especificia: 
```js
LibroDeTrabajo.Hoja.Rango.Celda
```
 
Un libro de trabajo *workbook* contiene una o más hojas de trabajo *worksheet*.
Un worksheet contiene colecciones de objetos de datos que están presentes en la hoja individual y da acceso a las celdas a través  de objetos de rango.
Un rango *range* representa un grupo de celdas contiguas. 
Los rangos se utilizan para crear y colocar tablas, gráficos, formas y otros objetos de organización o visualización de datos. 
```c
workbook.worksheet.getRange('A1').value
```
# Rangos
![image](https://www.aulaclic.es/estadistica-excel/graficos/Fig-1-2.png)
![image](https://cdn.exceltotal.com/wp-content/uploads/2013/02/celdas-y-rangos-en-excel-2013-01.png)
Un rango es un grupo de celdas contiguas en el libo de trabajo. Para obtener un rango se suele usar la notacion A1: que es usar numeros para las filas y letras para las columnas.
los rangos tienen tres propiedades principales: Valores, Formulas y formato. 
  - Obtener y establecer valores de las celdas
  - evaluar las fórmulas 
  - formato visual de las celdas.


```JS
function main(workbook: ExcelScript.Workbook) {
// Datos que cargaremos,3 filas y 3 columnas...
let data = [
	[1,2,3],
	[4,5,6],
	[7,8,9]
];

// Accedemos al workbook y a la hoja actual, obtenemos el rango y guardamos los datos en dicho rango.
workbook.getActiveWorksheet().getRange('a1:c3').setValues(data);

//Para cambiar el formato, accedamos en cascada hasta el rango
//workbook.hoja.rango
//en rango podemos acceder a los metodos de formato, cambiaremos el color del relleno.
//...rango.formato.relleno.color.cyan

	workbook.getActiveWorksheet().getRange('a1:c3').getFormat().getFill().setColor('Cyan');


}
```
