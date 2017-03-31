#Libro Excel CLIENTE con VBA
>Este Libro excel fue creado para poder agilizar Registros, Gestion y Reportes de los servicios ofrecidos por mi empleador, a nuestros clientes.
>
>En ese entonces todo el historial de servicios se tenían que registrar en **Libros Excel**, y nos referimos a miles de **servicios** (trabajos dentales) para muchos **clientes** (Clinicas, Hospitales, Entidades Gubernametales, Facultades de Universidades, Laboratorios Dentales y Consultas particulares ) que podían tener muchas sucursales en todo el país.
>
>Ante esto, inicialmente me ayude de C# para poder gestionar la información de los libros excel  y generar reportes mensuales para nuestros clientes; debido que las tablas excel no son un Motor de Base de Datos, y al tratarse de miles de registros mensuales, y ante la obligación de tener que seguir usando los libros excel como fuente de información, emepecé a usar **Visual Basic para Aplicaciones (VBA)** que al escribirse en el mismo libro, era mas rápido que C# para editar y manejar datos al momento de Ingresar, Analizar y Reportar la cantidad masiva de Información en los Libros.
>
>
###Estructura del Libro:
* #####Hoja "CONFIG"
>El libro sí o sí tiene que contener una hoja de configuración "CONFIG", en la cual se inicializan datos especificos para el correcto funcionamiento del libro y para que éste sea escalable. basta con modificar esta hoja "CONFIG" para poder adaptarlo al tipo de servicio, tipo de facturación y tipo de reporte del cliente específico; generalmente se recomienda usar un libro excel para cada cliente que de poseer varias sucursales, estas iran en distintas hojas; o también es posible agrupar en el mismo libro, distitnos clientes que tengan algo en común (por ejemplo, para cierto tipo de clientes a quienes se aplica el mismo arancel, estos pueden agruparse en un LIBRO EXCEL, cada cliente en una hoja).
>
* #####Hoja "DATOS"
>Acá irá la tabla con los datos del cliente, con los datos de las sucursales de cliente o con los datos de clientes que tienen algun funcionamiento en común. la tabla deberá ser creada con los datos que usted considere, con las columnas que usted requiera (por lo general datos de facturación y de contacto del cliente)
>
* #####Hoja "PLANTILLA"
>Esta hoja es el modelo para crear hojas de ingreso de datos; si se trata de una sola hoja de registros (solo un cliente por libro sin sucursales) hay que modificar esta hoja plantilla, cambiando el nombre y la estructura de la **TABLA DE INGRESOS** (las columnas que usted considere, con las formulas que usted especifique, es decir, cree su propia tabla). Si se trata de un cliente con sucursales o un grupo de clientes con algun funcionamiento en común, esta hoja plantilla siempre debe existir (por si se van agregando hojas) y solamente hay que editarla inicialmente cambiando la estructura de la **TABLA DE INGRESOS** y después copiarla para cada sucursal o cliente nuevo que se va agregando al libro.
>en las primeras filas de esta hoja, irán datos de Facturación, de contacto u otros datos que  usted deberá especificar en la hoja **"DATOS"**, algunas celdas de estas filas, deberán relacionarse (con fórmulas) a la hoja "DATOS" para mostrar información de cada sucursal o cliente, en cada hoja de ingreso.
>
* #####Hoja "ARANCEL"
>Si en la hoja "PLANTILLA" usted creó una tabla de ingreso de trabajos, existirá una o mas columnas de esta tabla que se relacionarán al código y al precio del servicio ofrecido. Es muy buena práctica, crear esta hoja "ARANCEL" con una tabla que contenga la relación, con tipo, codigo y precio, de los distinos servicios que ofrecemos; además con esto conseguirá que sus tablas sean dinámicas y relacionadas.

###FORMULARIOS DEL LIBRO
>A continuación detallaremos cada FORMULARIO del libro:
##form_reporte (Reporte Mensual de Servicios o Trabajos realizados)
![form_reporte](https://raw.githubusercontent.com/ofaber82/VBA/master/IMAGENES/form_reporte.png)
>Formulario que permite hacer reporte mensual de Servicios o trabajos no facturados  Según lo que especifique la configuración de la la hoja CONFIG (si existe una condición para el reporte, si se debe incluir la cabecera, si incluímos o no el desglose con el IVA, del total bruto, etc) y según las opciones que elija en el formulario (el periodo o mes a facturar, si el reporte es de esta sucursal o todas las sucursales). Esta opción generará 1 o mas archivos excel con la tabla de trabajos reportados, la cantidad dependerá si elije imprimr la sucursal actual o todas las sucursales.
>
>##form_revisión (Revisión e impresión de un periodo facturado)
![form_revision](https://raw.githubusercontent.com/ofaber82/VBA/master/IMAGENES/form_revision.png)
>Formulario que muestra (en un libro nuevo generado), la relación de trabajos correspondientes a un periodo facturado. Se generará y abrirá in libro con la selección adecuada para la impresion.

>##form_actualizar_sucursales (Actualización de montos de facturación
![form_actualizar_sucursales](https://raw.githubusercontent.com/ofaber82/VBA/master/IMAGENES/form_actualizar_sucursales.png)
>Actualiza los montos facturados en la hoja Facturación, para lo cual recorre todas las filas de la tabla obteniendo el monto total de cada facturación y actualizando con estos montos la hoja FACTURACIÓN. Esto aplica para una o para todas las sucursales.

###Ficha con Menú personalizado
>El libro cuenta con una ficha personalizada para poder elegir opciones avazadas y llamar a los formularios descritos anteriormente.
![ficha_personalizada](https://raw.githubusercontent.com/ofaber82/VBA/master/IMAGENES/ficha_personalizada.png)