# Taller de Productividad Basada en Herramientas Tecnológicas

## Descripción
"Los Perrillos" es un restaurante especializado en la venta de distintos tipos de hot dogs. Con solo una semana y media de operación, 
ha identificado la necesidad de estandarizar y digitalizar sus procesos para mejorar la eficiencia y reducir errores humanos. 
Esta solución tiene como objetivo transformar las operaciones análogas, como registrar ventas en una libreta, 
en un sistema digitalizado que agiliza y optimiza estas tareas.

## Problema identificado
El restaurante registra manualmente todas sus transacciones y gestiones de inventario. 
Este proceso, además de ser propenso a errores, implica una duplicación de esfuerzos cuando se transfiere la información a una computadora. 
El objetivo es eliminar el retrabajo y reducir los errores humanos en el proceso.

## Solución
A través de MS Excel con la ayuda de macros y programación en VBA, se desarrolló una solución que permite:

- Capturar pedidos.
- Generar tickets.
- Actualizar el inventario en tiempo real.
- Actualizar el balance de ventas en tiempo real.

## Arquitectura

La solución se basa en un archivo Excel compuesto por cinco hojas:

CAJA: Contiene el botón "CREAR UN NUEVO PEDIDO", el cual a través de una macro genera un formulario para que el cajero pueda llenarlo con los datos dependiendo del pedido del cliente, este formulario cuenta con botones para agregar distintos productos, para eliminar un pedido y para cancelarlo también, por otro lado, cuenta con el botón "Aceptar",
el cual manda la información a distintas hojas, cuando se le da clic al botón "Aceptar", automáticamente se genera un ticket y se almacena en la hoja TICKETS, también, a través
de la hoja COMPOSICION, el programa "lee" por que está constituido cada platillo y lo resta de la hoja INVENTARIO, por ejemplo, si se hace un pedido de un "Perrillo Clásico",
al darle al botón aceptar registrará el ticket, luego irá a la hoja COMPOSICION y en esa hoja está registrado que el "Perrillo Clásico" está constituido por un pan chico y  una salchicha de pavo chica, entonces esos elementos va y los resta al inventario y finalmente registra la venta en términos de dinero en la hoja BALANCE. 
Imagen del formulario: 
![image](https://github.com/Meinhof-code/TPBHT/assets/68880191/ed7da184-3cff-4bff-a4e0-2f75d01e938e)

TICKETS: Almacena los tickets generados en la hoja CAJA.
Imagen de la hoja: 
![image](https://github.com/Meinhof-code/TPBHT/assets/68880191/d85caeff-d2f6-405a-a20b-803bfb56efb9)


INVENTARIO: Almacena los datos del inventario y se actualiza constantemente por los pedidos generados en la hoja CAJA
Imagen de la hoja: 
![image](https://github.com/Meinhof-code/TPBHT/assets/68880191/2d832e22-c3bf-467b-b3f0-a0b00bba78dc)


BALANCE: Almacena los datos de las ventas y se actualiza constantemente por los pedidos generados en la hoja CAJA
Imagen de la hoja: 
![image](https://github.com/Meinhof-code/TPBHT/assets/68880191/cabd7331-d774-4da8-9a26-187adf073a10)


COMPOSICION: Sirve como referencia para el programa para ver la composición de los productos para después ir a descontarlos a la hoja INVENTARIO
Imagen de la hoja: 
![image](https://github.com/Meinhof-code/TPBHT/assets/68880191/d596672e-9af6-489c-ba0b-24d0b9c5505d)

## Requisitos del Sistema

- Microsoft Excel (última versión).
- Habilitación de macros para permitir la ejecucuón de VBA

## Instalación
1. Descargar el archivo Excel del repositorio.
2. Abrir el archivo y habilitar las macros para permitir la ejecución de VBA.

## Uso
Para registrar un nuevo pedido:

1. Presionar el botón "Crear un nuevo pedido".
2. Rellenar el formulario con la petición del cliente.
3. Presionar el botón "Aceptar".

## Contribución

Para contribuir al desarrollo y mejora de la herramienta:

1. Descargar el archivo Excel.
2. Implementar correcciones o mejoras.
3. Compartir las actualizaciones para ser revisadas e integradas.
