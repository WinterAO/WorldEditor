# WinterMapEditor

Editor de Mapas oficial de Comunidad Winter

# Como usar las Capas

Las capas son principalmente, para que lo entiendan más facil, la posición del grafico
con respecto a los personajes.

Capa 1: Pisos, paredes, y cualquier decoración similar a un piso que se use de Piso.

Capa 2: Costas, piedras.
Extra: 
- Las costas, al ser capa 2, reproducen un sonido dinstínto de los pasos.
- Las piedras siempre van en capa 2, ya que si le pasamos por arriba a una pierda, el personaje tambien le pasa por arriba.

Capa 3: Entradas de cavernas, Árboles (que no son objetos, porque los árboles pueden ser añadidos a un mapa, como objetos (arboles talables) y como capa 3 (no se pueden talar))
Extra: Normalmente, todo lo que sea una estatua, un monumento, un castillo, o una piramide, va en capa 3, porque si el personaje pasa por detrás de esa imagen, el body del pj desaparece. Como si fuera 3D.

Capa 4: Techos.

Todo lo que sea un techo, o que esté realizado con el fin de que tape todo. (Como una zona oscura por ejemplo)

# Como usar los objetos en un mapa

Entre los objetos normales del juego, podemos encontrar cosas como:

Yunkes (Que sirven para poder construír ítems con herrero)
Fragua (Lo mismo pero para hacer lingotes)
Árboles (Que sirven para ser talados por usuarios o trabajadores excepto en Winter)

Es fundamental saber porqué no debemos llenar los mapas de objetos, y esas razones son:

Que el servidor cuando abre, carga todos los objetos de cada mapa. E imaginen que un mapa
tiene normalmente: Piso, capas2, capas 3, capas 4, triggers, y algunos tienen montones de
objetos además.

Lo bueno, es que el servidor no carga los graficos de los mapas.. eso solo lo hace el cliente
para ahorrar espacio en memoria ^^. Si tota, ¿para que cargar los graficos en el servidor?
¿Quien los va a ver? xD nadie :P

Entonces, el servidor carga: Bloqueos, triggers (importantisimo saberlo), objetos, traslados
y las dimensiones del mapa :P osea 100x100 siempre.

Entonces, como talar se puede talar en zona segura, evitamos el hecho de poner objetos como
"arboles" en zonas inseguras, y con esto evitamos que el servidor carge objetos al pedo.
Y de esta manera optimizamos a nivel "técnicas" los mapas.

# Como usar los Triggers

Los triggers, son los TILES del mapa.

Para entender este tutorial, es fundamental haber leido el de Capas.

Por ejemplo:

Existen 7 tipos de triggers.

Trigger 0
Trigger 1
Trigger 2
Trigger 3
Trigger 4
Trigger 5
Trigger 6

Trigger 0 = NADA.

Trigger 1 = No afecta la lluvia, y la capa 4 desaparece cuando estamos pisando ese tile con nuestro personaje.

Trigger 2 = Lo mismo que el trigger 1, pero conbinado con el trigger 3.

Trigger 3 = Posición inválida para los NPCS. Los Npcs no pueden pisar ese tile, ni respawnear en él.

Trigger 4 = Lo mismo que el trigger 1, pero en este tile no se puede atacar ni aunque sea mapa inseguro.

Trigger 5 = En este trigger, si permanecemos mucho tiempo parados, nos lleva a la carcel por 10 minutos (sin pena en prontuario) por bloqueo de la libre circulación.

Trigger 6 = RING DE TORNEO. En este trigger vale todo. Las muertes de Criminales y Ciudadanos no cuentan. Cuando hay un espacio entre una arena y la otra (osea, que los triggers no son contínuos) no te deja atacar. Es para evitar que de una arena a la otra se ataquen.

# Tecnicas Fundamentales para Mapear

Como se explicó en el tutorial de "Como usar los objetos en un mapa" el uso de técnicas
de mapeo, no esta exento en Bender, ya que gran parte de la optimización de los mapas, va por
cuenta de los Mappers y que sepan hacer buenos mapas, de la manera más optima.

Esto no es muy dificil.

En los límites de un mapa (por ejemplo) podemos ver que no estan bloqueados todos los limites.
Solo esta bloqueado 1 tile entre mapa y mapa. Eso se debe, a que no hace falta en realidad
bloquear un área en la cual un personaje jamás pudiera llegar. Como así tampoco los npcs pueden
llegar.
Entonces, sabiendo que en los bordes del mapa, solo se puede "ver" lo que está, pero nunca se
podrá pisar con nuestro personaje, tenemos que evitar fundamentalmente, el hecho de poner:
Bloqueos
Triggers
Objetos
Traslados inalcanzables.

Y todo lo que posiblemente me esté olvidando ahora de ponerles, pero que ustedes van a dar
cuenta con el tiempo, mirando como ejemplo, los mapas que ya estan realizados por mi, o por
cualquiera de los mappers de BenderAO a cuales ya les trasmití estas técnicas de mapeo.

Con esto por ejemplo, se logró mejorar muchísimo la cantidad de FPS del cliente, y parece
extraño.. pero en las primeras versiones del AO, (e inclusive hasta alkon hoy en día) se
encontraban mapas que tenian en todos sus bordes, bloqueos, traslados, objetos, triggrs,
siendo estas cosas, imposibles de utlizar mientras consumían espacio en memoria reverendamente
al pedo.
