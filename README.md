# Tablero de gestión y rotación – Arepas Sary

Tablero web con ingresos, retiros y rotación temprana por mes y por empresa temporal.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | El tablero. |
| `data.json` | Los datos del tablero, **cifrados con contraseña**. Solo contienen conteos por mes y empresa, sin datos personales. |
| `actualizar.html` | Convierte la carpeta de Excel en un `data.json` nuevo. Los Excel se leen en el navegador y no se suben a ningún lado. |
| `procesar.js` | La lógica de cálculo que usa `actualizar.html`. |
| `cifrado.js` | Cifra y descifra el `data.json` con la contraseña (PBKDF2 + AES-256). |
| `logo.png` | (Opcional) Logo que aparece arriba a la izquierda. Si no está, se muestra un texto. |

## Contraseña

El tablero pide una contraseña. Sin ella, el `data.json` publicado es ilegible, aunque alguien lo descargue desde GitHub.
La página y su diseño sí son visibles para cualquiera con el enlace; lo protegido son los números.

- Usa siempre la misma contraseña al actualizar, para que quienes ya la tienen sigan entrando.
- Para cambiarla, genera el `data.json` con la nueva y compártela con tu equipo.
- Si la olvidas no se pierde nada: genera de nuevo el `data.json` desde los Excel con una contraseña nueva.
- La contraseña se recuerda solo mientras la pestaña esté abierta. El enlace "Bloquear tablero" la borra.

## Publicar por primera vez

1. Abre `actualizar.html` con doble clic desde tu computador, arrastra la carpeta del año, escribe la contraseña dos veces y descarga el `data.json`. Ponlo en esta misma carpeta.
2. En GitHub, crea un repositorio nuevo (por ejemplo `indicadores-coordinacion`), entra a **Add file → Upload files**, arrastra todos estos archivos y confirma con **Commit changes**.
3. Ve a **Settings → Pages**. En *Branch* elige `main` y la carpeta `/ (root)`. Guarda.
4. En uno o dos minutos aparece el enlace del tablero, con la forma `https://TU-USUARIO.github.io/tablero-sary/`.

## De dónde sale cada dato

- **Ingresos:** del Excel mensual de cada subcarpeta (hoja o tabla INGRESOS).
- **Planta de personal (total de empleados):** del archivo `TOTAL EMPLEADOS` (fila TOTAL debajo de cada título "MES DE AÑO"). Si un mes no está ahí, se usa la hoja "Indicador de rotación" del Excel mensual. Lo que escribas a mano en `actualizar.html` tiene prioridad sobre ambos.
- **Retiros:** del consolidado `INFO_RETIROS_INDICADOR.xlsx` en todos los meses que contenga (hoy, agosto y septiembre). Los meses que no estén ahí se toman de la hoja RETIROS de la carpeta (enero a julio).
- **Retiros por subproceso:** solo del consolidado, así que este detalle empieza en agosto de 2026.
- Si en la carpeta quedan varias versiones del consolidado (por ejemplo `(2)` y `(3)`), se usa solo la más reciente según su fecha de modificación y las demás se ignoran. Igual es mejor dejar una sola.
- El nombre de los archivos no importa: el consolidado se reconoce por sus columnas y el mes de cada Excel mensual por sus ingresos.

## Actualizar cada mes

1. Dentro de la carpeta general (por ejemplo `INDICADORES JULIANA`): agrega la subcarpeta del nuevo mes con el Excel de ingresos y retiros, reemplaza el consolidado `INFO RETIROS INDICADOR` por la versión nueva y actualiza `TOTAL EMPLEADOS` con el bloque del mes.
2. Abre `https://TU-USUARIO.github.io/tablero-sary/actualizar.html`.
3. Arrastra la carpeta general completa. Revisa el resumen, la columna **Planta** (corrígela si hace falta) y la sección "Revisa estas filas".
4. Usa **Ver tablero con estos datos** si quieres revisarlo antes de publicarlo.
5. Escribe la contraseña dos veces, descarga el `data.json` y súbelo al repositorio con **Add file → Upload files** (reemplaza el anterior).

## Nunca subas los Excel ni el .pbix

Contienen nombres, documentos y salarios de empleados. Aunque el repositorio sea privado, el sitio de GitHub Pages es público.
El archivo `.gitignore` bloquea los `.xlsx`, `.xls` y `.pbix` si algún día usas Git desde el computador, pero al subir por la web de GitHub ese bloqueo no aplica: revisa qué arrastras.

## Reglas de cálculo (iguales al Power BI "MES A MES")

- Se leen las tablas de Excel `INGRESOS` y `RETIROS` de cada archivo.
- Se ignoran los archivos con "ANTERIOR" en el nombre y los temporales de Excel (`~$`).
- SARY INNOVATION y GRANEL FOOD se cuentan como SARY. COMPLEMENTOS HUMANOS se muestra como COMPLEMENTOS.
- Los ingresos van al mes de ingreso; los retiros, al mes de retiro.
- Rotación temprana: retiros ocurridos entre 0 y 60 días después del ingreso. El % se calcula sobre los retiros del mes.
- Índice de rotación: retiros del mes ÷ planta total de empleados del mes. El voluntario cuenta solo renuncias (RENUNCIA VOLUNTARIA y RENUNCIA CON ANTECEDENTES); el involuntario, las terminaciones.
- Índice acumulado del año: suma de retiros ÷ planta promedio de los meses con planta registrada.
- Subprocesos graficados: EMPAQUE, PISO 3 y MERCADEO Y VENTAS de Medellín, Bogotá y Cali.
- Los gráficos de líneas siempre muestran todos los meses; el mes elegido se resalta en amarillo.
