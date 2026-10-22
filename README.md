<img width="70" height="70" alt="ViBa7" src="https://github.com/user-attachments/assets/f3761f3e-f757-4595-a493-8cbfbcbd4c36" /> # ViBa7 IDE (ViBa7-x86) 


[![Build & Test](https://img.shields.io/badge/.NET-10.0%20LTS-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20x86%20(WOW64)-0078D6?logo=windows)](https://microsoft.com)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/Tests-349%20Passing-success.svg)](tests/ViBa7.Tests)
[![Version](https://img.shields.io/badge/Version-00.001.0636-informational.svg)](src/ViBa7.IDE/AppVersion.cs)
[![Architecture](https://img.shields.io/badge/Architecture-Dual%20Toolchain%20(x32%20%7C%20x64)-orange.svg)](docs/TOOLCHAIN_SPEC.md)

<img width="800" height="600" alt="{63E68446-0F25-41F4-B23F-D057198AFB24}" src="https://github.com/user-attachments/assets/fb8d8b25-379c-426b-8690-f5a528af5f5a" />
  
</p>

---

## 1. Visión y Propósito del Proyecto

**ViBa7 IDE** es un entorno de desarrollo integrado (IDE) y diseñador visual moderno desarrollado en **C# y WPF sobre .NET 10 LTS**, ejecutándose como un proceso nativo en **32 bits (Win-x86)** con soporte de **4 GB de memoria virtual (`LARGEADDRESSAWARE`)**. 

Su misión central es proporcionar **100% de retrocompatibilidad en diseño y codificación con Microsoft Visual Basic 6.0 SP6**, permitiendo abrir, diseñar, editar, depurar y compilar proyectos clásicos (`.vbp`, `.vbg`, `.frm`, `.ctl`, `.pag`, `.bas`, `.cls`) sin alterar formatos, saltos de línea ni páginas de códigos ANSI.

### Estrategia de Compilación Bifurcada (Dual Toolchain)
El IDE desacopla por completo la interfaz del proceso de generación de código máquina, adoptando el estándar formal [TOOLCHAIN_SPEC.md](docs/TOOLCHAIN_SPEC.md):
1. **Compilador x32 (`vbc32.exe`)**: Desarrollo independiente para generar binarios PE32 nativos 100% compatibles con VB6 SP6 enlazados directamente a `msvbvm60.dll`.
2. **Compilador x64 (`vbc64.exe`)**: Desarrollo independiente para compilar proyectos nuevos creados desde cero en ViBa7 a ejecutables nativos de 64 bits (PE64), con soporte para tipos extendidos (`LongPtr`, `LongLong`, `PtrSafe`, `#If Win64`).

---

## 2. Fundamentación Arquitectónica: ¿Por qué el IDE en 32 Bits?

```
┌────────────────────────────────────────────────────────────────────────┐
│                   ESPACIO DE MEMORIA DEL PROCESO                       │
├───────────────────────────────────┬────────────────────────────────────┤
│   Proceso IDE x64                 │   Proceso ViBa7 IDE x86            │
│   ❌ NO puede cargar DLLs/OCX x86 │   ✔ Carga directa en memoria de    │
│   ❌ Requiere IPC / Out-of-proc   │     todos los OCX y TypeLibs       │
│   ❌ Pérdida de eventos nativos   │   ✔ OLE In-Place Activation nativo │
│   ❌ Incompatible con msvbvm60    │   ✔ Compatible directo con runtime │
└───────────────────────────────────┴────────────────────────────────────┘
```

1. **Alojamiento Nativo de Controles ActiveX (.OCX)**:  
   La inmensa mayoría de aplicaciones comerciales de VB6 utilizan controles `.ocx` de terceros (FarPoint Spread, MSComCtl, Sheridan, TrueDBGrid, etc.) cuyos códigos fuente están perdidos. En Windows, un proceso de 64 bits no puede cargar librerías de 32 bits en su espacio de memoria. Al ejecutarse en 32 bits, ViBa7 aloja, dibuja e interactúa directamente con cualquier control OCX mediante las interfaces estándar COM/OLE (`IOleObject`, `IOleClientSite`, `IOleInPlaceSite`).
2. **Capacidad de Orquestación x64**:  
   En Windows de 64 bits, un proceso de 32 bits dispone de 4 GB de memoria virtual gracias a `LARGEADDRESSAWARE` y puede iniciar, comunicar y depurar procesos y compiladores de 64 bits (`vbc64.exe`) sin ninguna restricción.

---

## 3. Características Implementadas

### 🎨 Diseñador Visual de Formularios y Controles (WYSIWYG)
* **Lienzo de Diseño Bidireccional**: Renderizado fidedigno a partir de las cabeceras `Begin VB.Form...End` de archivos `.frm`.
* **Soporte de Contenedores Jerárquicos**: Inserción y arrastre de controles dentro de contenedores (`Frame`, `PictureBox`), manteniendo su árbol visual y coordenadas relativas.
* **Portapapeles Completo en Diseño**: Operaciones de **Copiar**, **Cortar** y **Pegar** (<kbd>Ctrl+C</kbd>, <kbd>Ctrl+X</kbd>, <kbd>Ctrl+V</kbd>) con cálculo automático de sufijos numéricos unívocos para nombres duplicados.
* **Matrices de Controles (`Control Arrays`)**: Soporte para la propiedad `Index`, duplicación de controles con el mismo nombre y despacho del argumento `Index As Integer` en eventos compartidos.
* **Herramienta Gotero de Color (*EyeDropper*)**: Captura de color interactiva de cualquier píxel de la pantalla con asignación directa a propiedades de color (`BackColor`, `ForeColor`).
* **Manejo Integral de Recursos Binarios (`.frx`)**: Carga, visualización y persistencia de streams de imágenes, mapas de bits e iconos vinculados a formularios y controles sin desfasar sus offsets.
* **Rejilla Magnética (*Grid & Snap to Grid*)**: Rejilla de puntos ajustable con alineación magnética para posicionamiento preciso de controles.
* **Alternancia Rápida Código ↔ Objeto**: Atajos estándar (<kbd>F7</kbd> para Código, <kbd>Shift+F7</kbd> para Diseñador visual).

<p align="center">
  <img width="600" height="510" alt="Diseñador de Formularios" src="https://github.com/user-attachments/assets/55b15372-5360-405d-8ef9-a94e40fea98f" />
</p>

---

### 🎛️ Soporte Completo de UserControls (`.ctl` / `.ctx`)
* **Integración en Toolbox**: Los controles de usuario presentes en el proyecto se registran automáticamente en la Caja de Herramientas con el distintivo `🎛️` y su identificador calificado (`Proyecto.MiUserControl`).
* **Composición Visual Recursiva**: Al colocar un UserControl sobre un formulario, el diseñador lee recursivamente su archivo `.ctl` y dibuja con fidelidad todos sus controles constitutivos (etiquetas, botones, cajas de texto).
* **Extracción Automática de Propiedades y Enums**: Analizador estático que detecta declaraciones `Property Get/Let/Set`, bloques de enumeración públicos (`Enum`) y persistencia `PropertyBag`, exponiéndolos de forma interactiva en la Ventana de Propiedades.
* **Escala HiMetric ↔ Twips**: Conversión bidireccional exacta de dimensiones de diseño HiMetric (`_ExtentX`, `_ExtentY`) a Twips y píxeles ($1\text{ pulgada} = 2540\text{ HiMetric} = 1440\text{ Twips}$).
* **Contenedor de Pruebas (*Test Container*)**: Al pulsar <kbd>F5</kbd> sobre un archivo `.ctl` activo, se lanza en un contenedor de prueba interactivo para ensayar su ciclo de vida y eventos de forma aislada.

---

### 🧩 Soporte para Controles ActiveX / OCX de Terceros
* **Lector Nativo TypeLib COM**: Invocación P/Invoke a `oleaut32.dll!LoadTypeLibEx` con la bandera `REGKIND_NONE`, permitiendo inspeccionar metadatos de archivos `.ocx`, `.dll` y `.tlb` sin requerir permisos de administrador.
* **Diálogo Avanzado de Componentes (<kbd>Ctrl+T</kbd>)**:
  * Escaneo exhaustivo del Registro de Windows (vistas de 32 y 64 bits) y carpetas de sistema para listar todos los controles y TypeLibs registrados.
  * Casillas de verificación para activar/desactivar componentes en el proyecto activo.
  * Autoregistro de componentes en disco (`regsvr32.exe /s` y `LoadTypeLibEx REGKIND_REGISTER`).
  * Persistencia bidireccional de directivas `Object={GUID}#Ver#0; Archivo.ocx` en el archivo `.vbp`.
* **Páginas de Propiedades de Fabricante `(Custom)`**: Exposición de la propiedad `(Custom)` para invocar los cuadros de diálogo nativos del fabricante del control OCX.
* **Incorporación en Toolbox e IntelliSense Dinámico**: Los controles del OCX se suman automáticamente a la caja de herramientas con el glifo `🧩` y expanden el catálogo de miembros y eventos expuestos.

---

### 📋 Editor de Menús Clásico (<kbd>Ctrl+E</kbd>)
* **Diálogo Jerárquico de Menús**: Réplica moderna del editor de menús de VB6.
* **Visualización de Niveles**: Identificación visual de jerarquía mediante sangrías con puntos guía (`....Menú`) controlados mediante los botones **►** y **◄**.
* **Configuración de Propiedades**: Edición de `Caption`, `Name`, `Index`, `Shortcut` (<kbd>Ctrl+O</kbd>, <kbd>F1</kbd>, etc.), `Checked`, `Enabled` y `Visible`.
* **Reordenación y Sincronización**: Botones Subir (**▲**), Bajar (**▼**), Insertar, Eliminar y Siguiente, con serialización y deserialización automática de las estructuras `Begin VB.Menu...End` en el archivo `.frm`.

<p align="center">
  <img width="590" height="370" alt="Editor de Menús" src="https://github.com/user-attachments/assets/7f7ff856-be5b-47a5-9062-19128916caef" />
</p>

---

### 💡 Servicios de Lenguaje e IntelliSense
* **Auto List Members (<kbd>Ctrl+J</kbd> / `.`)**: Despliegue de listas de miembros con iconos tipados para propiedades, métodos y eventos al escribir el punto (`.`) o bajo demanda.
* **Auto Parameter Info (<kbd>Ctrl+I</kbd>)**: Tooltip interactivo con firmas completas de funciones y métodos, con resaltado en negrita del parámetro actualmente activo.
* **Auto Complete Word (<kbd>Ctrl+Espacio</kbd>)**: Autocompletado inteligente de palabras reservadas, tipos, variables locales y globales.
* **IntelliSense para Enumeraciones**: Detección contextual y sugerencia automática de miembros enum (ej. constantes `vbInformation`, `vbCritical`, `vbYesNoCancel`, `vbNormalFocus`).
* **Auto-Corrector Canónico de Mayúsculas/Minúsculas**: Corrección automática del estilo y casing oficial de palabras reservadas e identificadores reconocidos al confirmar una línea.
* **Catálogo Exhaustivo del Runtime**: Base de datos completa con funciones intrínsecas de cadena, matemáticas, conversión, fecha y tiempo (`Mid`, `InStr`, `MsgBox`, `Format`, etc.).

<p align="center">
  <img width="800" alt="IntelliSense y Autocompletado" src="https://github.com/user-attachments/assets/050198c3-ea1c-491c-bd68-390b7abf3557" />
</p>

---

### 📝 Editor de Código y Productividad Visual
* **Editor AvalonEdit Personalizado**: Motor de renderizado de texto optimizado para sintaxis VB6/ViBa7.
* **Plegado y Desplegado de Código (*Code Folding*)**: Colapso visual de rutinas `Sub`, `Function`, `Property`, bloques `Type...End Type` y `Enum...End Enum`.
* **Guías de Estructura de Control (*Branch Lines*)**: Líneas verticales guía que conectan la apertura y cierre de bloques `If...End If`, `Select Case`, `For...Next`, `Do...Loop` y `With...End With`.
* **Detalles de Fin de Bloque (*Inlay Hints / Ghost Text*)**: Anotaciones descriptivas automáticas al final de bloques extensos (ej. `End If  ' [If total > 0]`).
* **Operaciones de Edición Rápida**: Comentar/descomentar en bloque (<kbd>Ctrl+K, Ctrl+C</kbd> / <kbd>Ctrl+K, Ctrl+U</kbd>), inversión de asignaciones (`A = B` ⇄ `B = A`) y cambio de mayúsculas/minúsculas.

<img width="405" height="560" alt="{F3EE3937-B1D2-4FCE-8A25-9DDB1D69403E}" src="https://github.com/user-attachments/assets/468d4a4e-36c4-4358-b8db-c60fab850104" />

---

### 🔍 Exploración y Navegación Avanzada
* **Explorador de Proyectos Dual**:
  * **Modo Estándar**: Árbol tradicional agrupado por carpetas funcionales (Formularios, Módulos, Clases, Controles de usuario).
  * **Modo Avanzado**: Desglose profundo de miembros internos (Sub, Function, Property, Eventos, Constantes, Variables de módulo y Enums).
* **Búsqueda de Referencias y Usos (*Find References* - <kbd>Shift+F12</kbd>)**: Localiza todas las apariciones de un símbolo en el archivo, proyecto o grupo `.vbg`.
* **Historial de Navegación y Marcadores (*Bookmarks*)**:
  * Historial Atrás / Adelante (<kbd>Alt+←</kbd> / <kbd>Alt+→</kbd>).
  * Marcadores visuales en el margen (`🔖`) con navegación cíclica (<kbd>F2</kbd> / <kbd>Shift+F2</kbd>) y alternancia rápida (<kbd>Ctrl+F2</kbd>).
* **Explorador de Tareas (*Task List*)**: Detección automática de directivas `' TODO:`, `' BUG:`, `' FIXME:`, `' NOTE:` y `' HACK:`.
<img width="300" height="312" alt="{01E95DF5-CE1C-4676-8CEE-780BDEB583FD}" src="https://github.com/user-attachments/assets/181ee42f-efca-4972-8148-de8abc8a08ac" />

---

### ⚡ Asistentes de Código y Snippets
* **Asistente de Propiedades (*Property Builder*)**: Generación interactiva de bloques `Property Get/Let` (valores primitivos) y `Property Get/Set` (referencias a objetos).
* **Asistente de Procedimientos (*Procedure Builder*)**: Diálogo interactivo para declarar métodos con visibilidad, parámetros tipados (`ByVal`/`ByRef`/`Optional`) y plantillas de captura de errores.
* **Biblioteca de Snippets (AutoText con <kbd>Tab</kbd>)**: Expansión rápida para estructuras de control frecuentes (`if`, `ife`, `sel`, `for`, `dow`, `dou`, `whi`, `with`, `prop`, `sub`, `fn`, `err`).
* **Portapapeles Múltiple (*Clipboard Ring* - <kbd>Ctrl+Shift+V</kbd>)**: Historial dinámico de los últimos fragmentos copiados con selector emergente e inserción directa.

---

### 🛠️ Mantenimiento y Calidad de Código
* **Formateador de Código (*Code Formatter*)**: Indentación basada en anidación, alineación estética de cláusulas `As` en declaraciones `Dim` y normalización de saltos de línea.
* **Ordenador de Procedimientos (*Sort Procedures*)**: Clasificación alfabética o por visibilidad preservando el encabezado de declaraciones.
* **Comparador de Cambios (*File Diff Viewer*)**: Visor de diferencias unificado y lado a lado frente a la versión guardada en disco.
* **Estadísticas de Proyecto (*Project Statistics*)**: Métricas de líneas efectivas (SLOC), líneas de comentarios, densidad y total de procedimientos.
* **Motor de Manejo de Errores y Trazas**: Inserción o retiro masivo de esquemas de error y sentencias de trazabilidad con marcadores únicos (`' [VIBA7_TRACE]`).
* **Diseñador de Orden de Tabulación (*Tab Order Designer*)**: Visualización y reordenamiento interactivo o geográfico de la propiedad `TabIndex`.
* **Verificador de Teclas Rápidas (*Hotkey Designer*)**: Detección y corrección de colisiones de aceleradores `&` en etiquetas y botones.

---

### 🚀 Ejecución en Caliente (Hot-Run / <kbd>F5</kbd>) y Ventana Inmediata
* **Ejecución Instantánea en Memoria**: Lanzamiento inmediato de la interfaz y lógica del proyecto sin requerir compilación a disco.
* **Anfitrión Visual Interactivo (`LiveFormWindow`)**: Instanciación dinámica de controles con proxies bidireccionales (`LiveControlProxy`, `LiveFormProxy`).
* **Temporizadores Funcionales (`Timer`)**: Ejecución continua del temporizador con respeto a la propiedad `Interval`, disparo del evento `_Timer` y persistencia de variables estáticas.
* **Ventana Inmediata (*Immediate Window* - <kbd>Ctrl+G</kbd>)**:
  * Evaluación de expresiones mediante `? expr` o `Print expr`.
  * Asignación y ejecución de sentencias en tiempo de ejecución (`Variable = Valor`, llamadas a rutinas).
  * Consola de salida de instrucciones `Debug.Print`.
* **Controles de Depuración**: Iniciar (<kbd>F5</kbd>), Pausa (<kbd>Ctrl+Break</kbd>), Detener (<kbd>Shift+F5</kbd>), Paso a paso (<kbd>F8</kbd> / <kbd>Shift+F8</kbd>) y Puntos de Interrupción (<kbd>F9</kbd>).

---

### 📦 Arquitectura de Compilación: Toolchain Satélite x32 y x64
Conforme a [docs/TOOLCHAIN_SPEC.md](docs/TOOLCHAIN_SPEC.md), el IDE interactúa con compiladores satélites independientes mediante CLI y diagnósticos estructurados en JSON:
* **Compilador x32 (`vbc32.exe`)**:
  * Salida: PE32 (i386) nativo.
  * Runtime: Enlace canónico a `msvbvm60.dll!ThunRTMain`.
  * Estructuras: Grafo binario oficial `VB5!` (`VBHeader`, `ProjInfo`, `ObjectTable`).
  * 100% de compatibilidad binaria para ejecutar en cualquier Windows (desde Windows 98 hasta Windows 11).
* **Compilador x64 (`vbc64.exe`)**:
  * Salida: PE64 (x86-64 / AMD64) nativo.
  * Runtime: Enlace a `vbrunt64.dll` (Modern VB64 Runtime).
  * Sintaxis extendida: Tipos `LongPtr`, `LongLong` (entero 64 bits con sufijo `^`), `Declare PtrSafe Sub...`, `#If Win64 Then`.
  * Diseñado para proyectos nuevos creados desde cero que requieren superar el límite de 4 GB de memoria.

---

## 4. Estructura de la Solución y Base de Conocimiento

```text
ViBa7-x86.slnx
├── src/
│   ├── ViBa7.IDE/             # Aplicación WPF x86 (Shell, Diseñador, Editor, Diálogos)
│   ├── ViBa7.ActiveXHost/     # Anfitrión COM/OLE In-Place x86 para controles ActiveX (.OCX)
│   └── ViBa7.ProjectModel/    # Biblioteca de dominio (Parser .vbp, TypeLib, Persistencia, IntelliSense)
├── tests/
│   └── ViBa7.Tests/           # Suite automatizada xUnit (349 pruebas en verde)
├── LIBRARY VB6/               # Base de conocimiento condensada e indexada (VB6 / COM / Win32 / msvbvm60)
│   ├── VB6_LIBRARY_INDEX.md             # Índice maestro de ingeniería y referencia cruzada
│   ├── VB6_LIBRARY_DATABASE.json        # Base de datos estructurada de capítulos y tópicos
│   ├── TYPELIB_AND_RUNTIME_SYMBOLS.json # Catálogo de símbolos e interfaces COM de msvbvm60.dll y VB6.OLB
│   ├── CODE_SAMPLES_INDEX.json          # Índice de arquitectura empresarial (418 módulos VB6)
│   └── query_library.py                 # Herramienta CLI de consulta instantánea de la biblioteca
└── docs/
    ├── TOOLCHAIN_SPEC.md                # Especificación formal del contrato CLI/JSON con compiladores
    ├── PLAN_IMPLEMENTACION_ACTIVEX_VB6.md # Diagnóstico y hoja de ruta ActiveX/OCX al 100%
    └── adr/                             # Registros de decisiones arquitectónicas (ADR-001, ADR-002)
```

---

## 5. Instrucciones de Compilación y Ejecución

### Requisitos del Sistema
* **Sistema Operativo**: Windows 10 o Windows 11 (edición de 64 bits o 32 bits).
* **SDK Requerido**: [.NET 10.0 SDK LTS](https://dotnet.microsoft.com/download/dotnet/10.0) con componente de escritorio de Windows.

### 1. Compilar la Solución Completa
```bash
dotnet build ViBa7-x86.slnx
```

### 2. Ejecutar la Suite de Pruebas Automatizadas
El proyecto cuenta con **349 pruebas automatizadas** que validan la preservación de archivos, el parser `.vbp`/`.frm`, el motor de UserControls, el anfitrión ActiveX (`ViBa7.ActiveXHost`), la extracción de TypeLibs y la sincronización HiMetric:
```bash
dotnet test ViBa7-x86.slnx
```

### 3. Iniciar el IDE
```bash
dotnet run --project src/ViBa7.IDE
```

---

## 📄 Licencia y Reconocimientos

El código fuente de **ViBa7 IDE** se distribuye bajo los términos de la licencia **GNU General Public License v3.0 (GPL-3.0)**. Consulta el archivo [LICENSE](LICENSE) para conocer todos los detalles.

* **Autor**: David Rojas [Axio.UK] (Axiomedes).
* **Aviso Legal**: Este proyecto es un desarrollo independiente de código abierto y no incluye binarios ni código propietario descompilado de Microsoft Corporation. Visual Basic es una marca registrada de Microsoft Corporation. Las referencias a Visual Basic 6.0 se utilizan exclusivamente con propósitos de identificación técnica y compatibilidad de formatos.
