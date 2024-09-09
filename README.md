# SPLITX_ExcelFunction
Complemento de Excel para dividir texto en celdas verticales u horizontales usando múltiples delimitadores.

## Descripción
Este complemento de Excel permite dividir texto en celdas usando múltiples delimitadores.

## Idiomas de Función
IMAGEX (Inglés) / IMAGENX (Español).

## Características Principales
- Versatilidad: SPLITX puede manejar hasta tres delimitadores diferentes, lo que la hace extremadamente flexible para diversas necesidades de procesamiento de texto.
- Dirección Personalizable: Puedes elegir si deseas que las partes del texto se distribuyan horizontalmente (en filas) o verticalmente (en columnas).
- Fácil de Usar: Con una sintaxis clara y sencilla, SPLITX se integra perfectamente en tus hojas de cálculo de Excel, mejorando tu flujo de trabajo sin complicaciones.

## Descargar el Complemento
Puedes descargar el complemento SPLITX.xlam desde el siguiente enlace:
[Descargar SPLITX.xlam]([#](https://github.com/danraxai/SPLITX_ExcelFunction/blob/main/SPLITX.xlam))

## Instalación
1. Descargar el Archivo:
   - Descarga el archivo .xlam desde el enlace proporcionado.
2. Abrir Excel:
   - Abre Excel y ve a Archivo > Opciones > Complementos.
3. Administrar Complementos:
   - En la parte inferior de la ventana, selecciona Complementos de Excel en el menú desplegable y haz clic en Ir....
4. Agregar el Complemento:
   - Haz clic en Examinar... y navega hasta el archivo .xlam que descargaste.
   - Selecciona el archivo y haz clic en Aceptar.
5. Activar el Complemento:
   - Asegúrate de que la casilla junto a Splitx esté marcada y haz clic en Aceptar.

## Sintaxis
```
=SPLITX(texto, direccion, delimitador1, [delimitador2], [delimitador3])
```
- texto: El texto completo que deseas dividir en partes más pequeñas.
- direccion: 0 para distribuir las partes en filas, 1 para distribuirlas en columnas.
- delimitador1: El primer delimitador que se utilizará para dividir el texto.
- delimitador2 (Opcional): Un segundo delimitador adicional.
- delimitador3 (Opcional): Un tercer delimitador adicional.

## Ejemplos de Uso
### Ejemplo 1: Dividir Texto en Columnas
Supongamos que tienes el texto "manzana,pera;plátano" en la celda A1 y quieres dividirlo en columnas usando la coma y el punto y coma como delimitadores. Puedes usar la función SPLITX de la siguiente manera:
```
=SPLITX(A1, 1, ",", ";")
```
Esto dividirá el texto en "manzana", "pera" y "plátano" y los colocará en celdas adyacentes verticalmente.

### Ejemplo 2: Dividir Texto en Filas
Si deseas dividir el mismo texto en filas, puedes usar:
```
=SPLITX(A1, 0, ",", ";")
```
Esto colocará "manzana", "pera" y "plátano" en celdas adyacentes horizontalmente.

## Beneficios
- Ahorro de Tiempo: Automatiza la tarea de dividir texto, eliminando la necesidad de hacerlo manualmente.
- Precisión: Reduce el riesgo de errores humanos al procesar grandes volúmenes de datos.
- Flexibilidad: Adapta la función a tus necesidades específicas con múltiples delimitadores y opciones de dirección.

## Conclusión
La función SPLITX es una adición invaluable a tu arsenal de herramientas de Excel, proporcionando una solución eficiente y flexible para la manipulación de texto. Ya sea que estés trabajando con listas de productos, datos de clientes o cualquier otro tipo de información textual, SPLITX te ayudará a organizar y gestionar tus datos de manera más efectiva.
