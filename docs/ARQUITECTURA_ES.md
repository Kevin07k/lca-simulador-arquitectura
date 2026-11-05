# P02 — Arquitectura y contratos del simulador (propuesta para revisión)

**Estado:** propuesta de Javier para revisión de Kevin en la issue [#2](https://github.com/Kevin07k/lca-simulador-arquitectura/issues/2). La aprobación del PR establece el acuerdo del equipo. Plataforma: Microsoft Excel con VBA (`.xlsm`). El libro, los módulos VBA y la interfaz corresponden a tareas posteriores.

## 1. Parámetros y mapa

| Elemento | Contrato |
| --- | --- |
| Dato y bus de datos | 8 bits sin signo (00–FFh); las operaciones aritméticas envuelven módulo 256. |
| Dirección virtual y PC | 9 bits (000–1FFh, 512 bytes). `PC` apunta al comienzo de una instrucción. |
| Dirección física y bus de direcciones | 8 bits (00–FFh, 256 bytes). |
| Registros generales | `AX`, `BX`, `CX`, `DX`: 8 bits; códigos 0, 1, 2, 3. |
| Registros de control | `PC`, `MAR_V`: 9 bits; `MAR_P`, `MDR`, `IR`, `FLAGS`: 8 bits; `IR` contiene el opcode. `ZF`, `CF`, `SF` son bits de `FLAGS`. |
| Paginación | 32 bytes/página; 16 páginas virtuales (`VPN` 0–15) y ocho marcos físicos (`PFN` 0–7). `VPN = VA >> 5`, desplazamiento `VA & 1Fh`; `PA = PFN*32 + desplazamiento`. |

| Virtual | Uso | Política |
| --- | --- | --- |
| 000–0FFh (VPN 0–7) | Código: 64 instrucciones de 4 bytes | Fetch permitido; escritura CPU prohibida. Carga inicial por el cargador. |
| 100–1DFh (VPN 8–14) | Datos: 224 bytes | Lectura y escritura CPU permitidas; fetch prohibido. |
| 1E0–1E7h | Puerto/buffer de teclado de ocho bytes | Solo dispositivo mediante `IN`, no direccionable por `LD/ST`. |
| 1E8–1EFh | Puerto/buffer LCD de ocho bytes | Solo dispositivo mediante `OUT`, no direccionable por `LD/ST`. |
| 1F0–1FFh | Reserva E/S | Acceso CPU directo prohibido. |

| Físico | Uso |
| --- | --- |
| 00–DFh (PFN 0–6) | Siete marcos reemplazables para VPN 0–14. |
| E0–E7h (PFN 7) | Ocho bytes de la cola circular del teclado. |
| E8–EFh (PFN 7) | Ocho bytes del buffer de video LCD 8×1. |
| F0–FFh (PFN 7) | Reserva E/S; nunca se asigna a páginas de usuario. |

La VPN 15 está reservada, no posee entrada presente asignable ni participa en reemplazo; sus desplazamientos 0–15 se reflejan en el marco físico 7 **solo por el controlador E/S**. No se usa traducción de páginas normal ni caché para E/S. No se suponen regiones virtuales lineales en RAM: toda VPN 0–14 puede ocupar cualquiera de los siete marcos disponibles.

## 2. ISA y estado visible

P04 concreta la distribución de FLAGS: bit 0 = ZF, bit 1 = CF, bit 2 = SF; bits 3–7 en cero. Ver [API y verificación CPU/ALU](P04_CPU_ALU.md).

**Instrucción fija de cuatro bytes**, en orden de direcciones crecientes: `[opcode, R, X, H]`. `R` es un código de registro 0–3 o cero cuando no se utiliza; `X` es inmediato de 8 bits, código de segundo registro o byte bajo de dirección; `H` es cero excepto como byte alto de una dirección virtual, donde solo se permite 00h o 01h. Dirección operando = `256*H + X`. Todos los campos no utilizados valen cero; un campo no canónico produce `INVALID_ENCODING`. El ensamblador acepta hexadecimales de dos o tres dígitos con sufijo `h` y etiquetas resueltas a dirección absoluta de 9 bits. La imagen de código se llena con `HLT` (`01 00 00 00`) después de la última instrucción validada; el código no puede superar 256 bytes. El punto de entrada predeterminado es 000h.

| Opcode | Mnemónico | R | X / H | Efecto al completar |
| --- | --- | --- | --- | --- |
| 00h / 01h | `NOP` / `HLT` | 0 | 0 / 0 | NOP avanza; HLT detiene y deja PC en el HLT. |
| 10h | `MOV R,#imm` | destino | inmediato / 0 | R = inmediato. |
| 11h | `MOV R,R2` | destino | registro fuente / 0 | R = R2. |
| 12h | `LD R,[VA]` | destino | byte bajo / alto | R = byte de datos en VA. |
| 13h | `ST [VA],R` | fuente | byte bajo / alto | byte de datos en VA = R. |
| 20h / 21h / 22h | `ADD R,R2` / `SUB R,R2` / `CMP R,R2` | primer registro | segundo registro / 0 | ADD/SUB guardan resultado en R; CMP solo actualiza flags con R−R2. |
| 23h / 24h | `INC R` / `DEC R` | registro | 0 / 0 | R = R±1. |
| 30h / 31h / 32h / 33h / 34h | `JMP VA` / `JZ VA` / `JNZ VA` / `JC VA` / `JS VA` | 0 | byte bajo / alto | PC = destino si se cumple la condición. |
| 40h / 41h | `IN R` / `OUT R` | registro | 0 / 0 | Consume teclado / escribe LCD; puertos implícitos. |

Una instrucción comienza en una dirección múltiplo de cuatro y termina dentro de la misma página de 32 bytes. Antes de ejecutar, el decodificador valida opcode, campos, códigos de registro y direcciones: `LD/ST` solo 100–1DFh; saltos solo 000–0FCh y alineados a cuatro. Un `PC` fuera de código o desalineado es error. No hay acceso indirecto ni instrucciones de tamaño variable. El cargador valida programa completo antes de sustituir la imagen vigente.

`ZF = 1` si el resultado aritmético de ocho bits es cero; `SF` copia el bit 7; `CF` indica acarreo en ADD/INC o préstamo en SUB/CMP/DEC. Los cinco opcodes aritméticos actualizan los tres flags, incluso CMP, que no escribe R. MOV, LD, ST, saltos, IN, OUT, NOP y HLT no modifican flags. Un salto que no se toma avanza cuatro bytes. En finalización normal `PC += 4`, salvo salto tomado y HLT. La suma fuera del espacio de código genera `PC_OUT_OF_CODE` en el próximo fetch.

## 3. Traducción, reemplazo y errores

La tabla contiene 16 entradas `VPN, P (presente), PFN, dirty, ref`. Las VPN 0–14 comienzan con `P=0`, `dirty=ref=0`, `PFN` vacío; VPN 15 se marca `RESERVED`. La imagen de respaldo («swap») almacena exactamente 32 bytes por cada VPN 0–14. El cargador inicializa código y datos en ese respaldo. Al acceder a una VPN normal: validar región/permisos → consultar PTE → si no está presente, tratar `PAGE_FAULT` interno → traducir → consultar caché → RAM. `PAGE_FAULT` es observable en la traza y luego se reintenta la **misma** microoperación; nunca se repite un efecto ya comprometido.

Con marcos libres se toma el PFN libre menor. Si no hay, se emplea CLOCK sobre PFN 0–6, cursor inicial 0: si `ref=1`, ponerlo en 0 y avanzar; el primer `ref=0` es víctima y el cursor queda en el marco siguiente. Antes de reutilizar un PFN, invalidar todas sus líneas en L1I, L1D y L2; si `dirty=1`, copiar exactamente 32 bytes de RAM a respaldo de la VPN víctima; invalidar su PTE. Leer 32 bytes de la VPN entrante desde respaldo al PFN, marcar `P=1`, `PFN`, `dirty=0`, `ref=0`, y reintentar; el acceso completado marca `ref=1` y una escritura `dirty=1`. Esto exige siete marcos reemplazables y permite provocar reemplazos con ocho VPN distintas.

Errores fatales: `VA_OUT_OF_RANGE` (VA <0 o >511), `PC_OUT_OF_CODE`, `UNALIGNED_PC`, `PROTECTION` (fetch fuera de código, datos fuera de datos o acceso directo E/S), `INVALID_OPCODE`, `INVALID_ENCODING`, `INVALID_REGISTER`, `INVALID_BRANCH_TARGET` y `BUS_CONTENTION`. Un error detiene RUN/STEP, deja `PC` en la instrucción infractora, muestra código y dirección relevantes y no compromete la instrucción. Una falta de página válida no es fatal. `IN_EMPTY` es espera, no error: se conserva PC/registro/flags y RUN se pausa hasta nueva entrada; el botón STEP puede reintentar. El acceso directo a direcciones E/S no se convierte en `IN/OUT` automáticamente.

## 4. Cachés y contratos de módulos

El camino normal es **CPU → traductor → L1I (fetch) o L1D (dato) → L2 unificada → RAM física**. Direcciones de etiqueta de caché siempre físicas. Las tres cachés usan línea de cuatro bytes, mapeo directo y reemplazo por sobrescritura: L1I = 2 líneas (8 B), L1D = 2 líneas (8 B), L2 = 8 líneas (32 B). `línea = PA >> 2`; `índice = línea mod número_de_líneas`; `etiqueta = línea div número_de_líneas`. Ninguna línea cruza un límite de página/marco. En fallo de L1 se consulta L2; en fallo de L2 se lee la línea completa de RAM. L1I y L1D mantienen hit/miss separados; L2 cuenta una consulta por fallo de L1, clasificada como hit o miss. Los contadores se incrementan por acceso de **un byte solicitado por CPU**, no por cada byte de relleno.

Las escrituras ST a datos son *write-through* hasta L2 y RAM: en miss se carga la línea (*write-allocate*), después se actualiza L1D, L2 y RAM. RAM es la fuente autoritativa para swap. El código es inmutable durante RUN; el cargador reconstruye respaldo e invalida todas las líneas antes de ejecutar un programa nuevo. No hay caché de E/S; el acceso con `IO=1` usa directamente E0–EFh. Reset e intercambio de marcos invalidan líneas apropiadas antes de que se observe RAM reutilizada.

| Módulo y operación lógica | Entradas → salidas/efectos |
| --- | --- |
| CPU `Fetch(PC)` | PC válido → 4 bytes en orden, IR=primer byte, `MAR_V=VA`, `MAR_P=PA`, `MDR=byte` por lectura; no adelantar PC hasta terminar instrucción. |
| Traductor `Translate(VA, tipo)` | VA 9 bits, `tipo=FETCH/READ/WRITE` → PA 8 bits o `PAGE_FAULT` recuperable o error de protección/rango; marca ref al completar, dirty tras escritura. |
| Caché `Read/Write(PA,tipo,valor)` | PA física, `tipo=FETCH/DATA`; retorna byte/estado y evento hit/miss; nunca recibe VA ni E/S. |
| RAM `Read/Write(PA,valor)` | PA 00–DFh para páginas; E0–EFh solo controlador E/S; retorna byte o estado. Acceso a F0–FFh prohibido. |
| Bus `Transfer(origen,destino,PA,valor,señales)` | Dirección física y dato, un único emisor activo, receptor habilitado; registra cada ciclo `READ/WRITE`, `Clock`, `IO`, `Enable/Load`, y estado `High-Z` si nadie conduce; dos emisores → `BUS_CONTENTION` antes de escribir. |
| E/S `IN(R)` / `OUT(R)` | `IN` lee una vez la cabeza FIFO y carga R; `OUT` escribe R en el video y avanza cursor; señales `En_KBD`/`Ld_LCD` y eventos con `IO=1`. |

En cada flanco de `Clock`, el bus registra un solo ciclo de transferencia. En reposo: `Read=Write=IO=Enable=Load=0`, sin emisor, dato `High-Z`. Lectura RAM (incluido relleno de línea): `Read=1, Write=0, IO=0`, emisor RAM, receptor caché/MDR; escritura ST: `Read=0, Write=1, IO=0`, emisor CPU/caché, receptor RAM. IN: `Read=1, Write=0, IO=1, En_KBD=1`, emisor teclado, receptor registro; OUT: `Read=0, Write=1, IO=1, Ld_LCD=1`, emisor registro, receptor LCD/RAM. `Enable` habilita exclusivamente al emisor seleccionado y `Load` al receptor; `Read` y `Write` nunca son 1 simultáneamente. Un hit en caché no crea ciclo de RAM externa y deja ese bus en High-Z. La traza identifica microetapa, VA si corresponde, PA, dato, origen, destino y señales; el arbitraje detecta dos emisores antes de escribir.

La cola de teclado es FIFO circular de ocho bytes físicos E0–E7h, con `head`, `tail` y `count` (0–8) como estado del controlador. Al estar llena se **rechaza** el byte nuevo y se muestra `KBD_FULL` sin modificar la cola. En `IN` vacío, no se mueve PC ni `head`; con dato, se lee E0h+head y solo al comprometer R se avanza `head`, se reduce `count` y se avanza PC. `En_KBD=1` durante la lectura. El LCD tiene ocho bytes E8–EFh, inicializados a espacio (20h), y un `cursor` 0–7: `OUT` escribe E8h+cursor, pone `Ld_LCD=1`, actualiza la celda visible y avanza cursor módulo 8. El octavo carácter llena la última posición; el noveno reemplaza la primera. `OUT` no cambia la cola ni flags. El estado del puerto y su RAM física permanecen coherentes.

Reset/carga validada: borrar registros generales y flags; PC = punto de entrada 000h; desactivar señales y buses a High-Z; limpiar cola, punteros y LCD (ocho espacios, cursor 0); invalidar cachés y poner contadores a cero; inicializar 00–DFh a cero y PTE no presentes; restaurar respaldo inicial del programa/datos (código cargado y datos a cero). HLT y errores paran RUN; Reset los limpia. Una falta recuperable conserva PC y no produce escritura duplicada al reanudar.

## 5. Casos manuales para revisión

**Traducción.** PTE[10] = `(P=1, PFN=3, dirty=0, ref=0)`; VA `14Dh` (333) da VPN `10`, desplazamiento `0Dh` (13), PA `6Dh` (109). Es lectura de datos válida; `ref` queda 1. VA `1E8h` se reserva a LCD; `LD` debe producir PROTECTION, mientras `OUT` escribe a E8h+cursor sin buscar PTE[15]. VA `200h` produce VA_OUT_OF_RANGE.

**Fetch.** PTE[1] = `(P=1, PFN=2)`, PC=`020h`, RAM `40–43h = 10 00 2A 00`. Suponiendo cachés vacías, los cuatro fetch de bytes traducen `020–023h → 040–043h`. En el primero: miss L1I, miss L2, relleno de cuatro bytes desde RAM; en los tres restantes: hit L1I. Decodifica `MOV AX,#2Ah`; al completar, AX=`2Ah`, PC=`024h`, flags intactos. Contadores derivados: L1I hit=3/miss=1; L2 hit=0/miss=1. La traza de bus muestra cuatro lecturas físicas de relleno (`040–043h`) y señales READ=1, IO=0, nunca dos emisores. Esta es una expectativa manual, no una prueba del libro todavía inexistente.

**E/S.** Estado inicial: cola vacía, LCD ocho espacios, cursor=0. Ingresar `A` (`41h`) escribe E0h y deja count=1. `IN AX` con `En_KBD=1`, READ=1, IO=1 lee E0h; al comprometer AX=`41h`, count=0, head=1. `OUT AX` con `Ld_LCD=1`, WRITE=1, IO=1 escribe E8h=`41h`; celda LCD[0]=`A`, cursor=1. Un segundo IN sin tecla queda en espera y no consume ni adelanta PC. Estos eventos se podrán comprobar al implementar #13 y #14.

## 6. Alcance y decisión de equipo

Este archivo fija los **contratos propuestos** para las issues #3–#14; las capacidades de memoria/caché y los ocho bytes de teclado/video provienen de sus criterios de aceptación. Las elecciones de ISA fija, rangos exactos, PFN reservado, CLOCK, escritura write-through, control de FIFO/cursor y códigos de error se proponen aquí y deben revisarse con Kevin antes de marcar #2 como acordada. Si el «documento de planeación 02» contiene valores distintos, ajustar este contrato y su versión inglesa **antes** de aprobar el PR. No se afirma haber ejecutado VBA ni capturas inexistentes.

[English version](ARCHITECTURE.md)
