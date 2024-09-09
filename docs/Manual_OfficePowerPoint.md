



# Office PowerPoint
  
Modulo para controlar Microsoft Office Power Point.  

*Read this in other languages: [English](Manual_OfficePowerPoint.md), [Português](Manual_OfficePowerPoint.pr.md), [Español](Manual_OfficePowerPoint.es.md)*
  
![banner](imgs/Banner_OfficePowerPoint.png o jpg)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Descripción de los comandos

### Nueva presentación
  
Crea un nueva presentación Power Point
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Abrir presentación
  
Abre una presentación Power Point
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo||archivo.pptx|

### Guardar presentación
  
Guarda la presentación de Power Point abierta
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar archivo||archivo.pptx|

### Obtener tipo de diapostiva
  
Obtiene una lista de todos los tipos de diapositivas que existen
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado||res |

### Insertar diapositiva
  
Inserta una nueva diapositiva a la presentación
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Tipo|||

### Escribir en presentación
  
Escribe en una presentación Power Point.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de slide||0|
|Id o nombre del elemento donde se escribirá||Title 1|
|Tamaño de fuente||8|
|Alineación||Center|
|Negrita||True|
|Cursiva||True|
|Subrayar||True|
|Escriba texto||Lorem ipsum |

### Cerrar presentación
  
Cierra la presentación que se está ejecutando
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Agregar imagen
  
Agrega una imagen a la presentación
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la imagen||imagen.jpg|
|Numero de slide||0|
|Posición||2,4.5 |
|Alto||5|

### Agregar cuadro de texto
  
Agrega un cuadro de texto a la presentación.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto||lorem ipsum|
|Posición||x, y|
|Dimensiones||ancho, alto|
|Numero de slide||0|

### Editar texto
  
Modifica el texto de una diapositiva existente
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de slide||0|
|Forma||Title 1|
|Tamaño de fuente||8|
|Alineación||Center|
|Negrita||True|
|Cursiva||True|
|Subrayar||True|
|Texto||Lorem ipsum|

### Obtener elementos de la presentacion
  
Lista los elementos que estan en la presentacion
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de slide||0|
|Guardar resultado||res|

### Editar propiedades de un elemento
  
Edita las propiedades de un elemento en una diapositiva
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de slide||0|
|Id o nombre del elemento||Title 1|
|Posición||x, y|
|Dimensiones||ancho, alto|
|Rotación||45.0|
