#Libro Excel CLIENTE con VBA
>Este Libro excel fue creado para poder agilizar Registros, Gestion y Reportes de los servicios ofrecidos por mi empleador, a nuestros clientes.
>
>En ese entonces todo el historial de servicios se tenían que registrar en **Libros Excel**, y nos referimos a miles de **servicios** (trabajos dentales) para muchos **clientes** (Clinicas, Hospitales, Entidades Gubernametales, Facultades de Universidades, Laboratorios Dentales y Consultas particulares ) que podían tener muchas sucursales en todo el país.
>
>Ante esto, inicialmente me ayude de C# para poder gestionar la información de los libros excel  y generar reportes mensuales para nuestros clientes; debido que las tablas excel no son un Motor de Base de Datos, y al tratarse de miles de registros mensuales, y ante la obligación de tener que seguir usando los libros excel como fuente de información, emepecé a usar **Visual Basic para Aplicaciones (VBA)** que al escribirse en el mismo libro, era mas rápido que C# para editar y manejar datos al momento de Ingresar, Analizar y Reportar la cantidad masiva de Infrmación en los Libros.
>
>
###Estructura del Libro:
* #####Hoja "CONFIG"
>El libro sí o sí tiene que contener una hoja de configuración "CONFIG", en la cual se especifican datos especificos para el correcto funcionamiento. basta con modificar esta hoja "CONFIG" para poder adaptarlo al tipo de servicio, tipo de facturación y tipo de reporte del cliente específico; generalmente se recomienda usar un libro excel para cada cliente que de poseer varias sucursales, estas iran en distintas hojas para cada una de ellas; o también es posible agrupar en el mismo libro, distitnos clientes que tengan algo en común (por ejemplo, para cierto tipo de clientes a quienes se aplica el mismo arancel, estos pueden agruparse en un LIBRO EXCEL, cada cliente en una hoja).
>
* #####Hoja "DATOS"
>Acá ira la tabla con los datos del cliente, de las sucursales de cliente o de los clientes que tienen algun funcionamiento en común. la tabla deberá ser creada con los datos que usted considere, con las columnas que usted requiera (por lo general datos de facturación y de contacto del cliente)
>
* #####Hoja "PLANTILLA"
>Esta hoja es el modelo para crear hojas de ingreso de datos; si se trata de una sola hoja de registros (solo un cliente por libro sin sucursales) hay que modificar esta hoja plantilla, cambiando el nombre y la estructura de la **TABLA DE INGRESOS** (las columnas que usted considere, con las formulas que usted especifique, es decir, cree su propia tabla). Si se trata de un cliente con sucursales o un grupo de clientes con algun funcionamiento en común, esta hoja plantilla siempre debe existir (por si se van agregando hojas) y solamente hay que editarla inicialmente cambiando la estructura de la **TABLA DE INGRESOS** y después copiarla para cada sucursal o cliente nuevo que se va agregando al libro.
>en las primeras filas de esta hoja, irán datos de Facturación, de contacto u otros datos que  usted deberá especificar en la hoja **"DATOS"**, algunas celdas de estas filas, deberán relacionarse (con fórmulas) a la hoja "DATOS" para mostrar información de cada sucursal o cliente, en cada hoja de ingreso.
>
* #####Hoja "ARANCEL"
>Si en la hoja "PLANTILLA" usted creó una tabla de ingreso de trabajos, existirá una o mas columnas de esta tabla que se relacionarán al código y al precio del servicio ofrecido. Es muy buena práctica, crear esta hoja "ARANCEL" con una tabla que contenga la relación, con tipo, codigo y precio, de los distinos servicios que ofrecemos; además con esto conseguirá que sus tablas sean dinámicas y relacionadas.
