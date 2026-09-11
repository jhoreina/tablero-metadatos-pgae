# Dashboard de Metadatos y Gobernanza de Datos PGAE (SDA)

Este repositorio contiene la plataforma interactiva web del **Modelo de Metadatos y Gobernanza de Datos PGAE v1.0** de la **Secretaría Distrital de Ambiente (SDA)** de la Alcaldía Mayor de Bogotá D.C.

---

## 🌿 Características de la Plataforma
- **Línea Gráfica Institucional**: Desarrollada bajo los colores y estilos institucionales de la SDA (Verde Oscuro `#577515`, Verde Brillante `#89BC24` y fondo neutro `#F9F9F9`).
- **Datos Reales e Integrales**: Cargados directamente del archivo final `Modelo_Metadatos_PGAE_SDA-vf.xlsx` (230 atributos gobernados en 3NF con metadatos completos).
- **Mecanismos de Actualización Dinámica (Automatización)**:
  1. **Generador Python (`generar_dashboard.py`)**: Script que lee cualquier versión del archivo Excel y compila nuevamente `index.html`.
  2. **Cargador Web en Vivo**: Botón **"📁 Cargar Excel (.xlsx)"** dentro de la interfaz web (`index.html`) mediante SheetJS para actualizar la matriz, KPIs, inventario y gráficos directamente en el navegador en tiempo real.
- **Navegación por Pestañas Interactivas**:
  1. 📊 **Visión General**: Métricas KPI, Gráficos dinámicos, Matriz "Firma de Datos" (Eslabón × Campo) con filtro por clic y Hoja de Ruta OTI.
  2. 🗂️ **Inventario de Atributos**: Catálogo gobernado con buscador en tiempo real, filtros multinivel y visor modal detallado.
  3. 🎯 **Campos de Acción y Eslabones**: Fichas explicativas de los 9 vectores temáticos y 8 procesos institucionales.
  4. ❓ **Preguntas Estratégicas**: Demanda analítica de información (PE-01 a PE-15).
  5. 🧩 **Dimensiones Maestras**: Claves de homologación y gobernanza OTI (DM-01 a DM-07).
  6. 📈 **Indicadores Derivados**: Fórmulas analíticas (CEA, CEE, RESPEL, IDAE, Predial, Renta).
  7. 💻 **Catálogo de Fuentes**: Inventario técnico de sistemas maestros y accesos OTI.
- **Exportación Abierta**: Posibilidad de exportar el inventario completo en formatos CSV y JSON.

---

## 🔄 ¿Cómo actualizar el Dashboard si cambian los datos en Excel?

Tienes **dos opciones** muy sencillas para actualizar la información:

### Opción 1: Mediante el Script de Python (Recomendado para GitHub)
Si modificas o reemplazas el archivo `Modelo_Metadatos_PGAE_SDA-vf.xlsx` (o creas una nueva versión en Excel), ejecuta en tu terminal:

```bash
python generar_dashboard.py
```
O si deseas especificar la ruta de otro archivo Excel:
```bash
python generar_dashboard.py "ruta/a/tu_nuevo_archivo.xlsx"
```
El script extraerá automáticamente todas las pestañas, recalculará los conteos de la Matriz y volverá a generar el archivo `index.html` actualizado.

### Opción 2: Carga Directa en el Navegador Web
1. Abre el archivo `index.html` en tu navegador web.
2. En la esquina superior derecha del encabezado institucional, haz clic en el botón **"📁 Cargar Excel (.xlsx)"**.
3. Selecciona tu archivo Excel actualizado. La plataforma procesará los nuevos atributos y recalculará la Matriz, Gráficos y Tablas al instante.

---

## 🚀 Guía de Publicación en GitHub Pages

Para publicar este dashboard en tu cuenta de GitHub y compartir el enlace público:

### Paso 1: Crear un Repositorio en GitHub
1. Ingresa a tu cuenta de [GitHub](https://github.com/) y haz clic en **New Repository**.
2. Asigna un nombre al repositorio (ej: `modelo-metadatos-pgae-sda`).
3. Selecciona la visibilidad **Public** y crea el repositorio.

### Paso 2: Subir los Archivos al Repositorio
En tu consola local (Git Bash / PowerShell), ejecuta:

```bash
git init
git add .
git commit -m "Actualización del Dashboard de Metadatos PGAE SDA desde archivo -vf.xlsx"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/modelo-metadatos-pgae-sda.git
git push -u origin main
```

*(Reemplaza `TU_USUARIO` y `modelo-metadatos-pgae-sda` por tus datos reales).*

### Paso 3: Activar GitHub Pages
1. Ve a **Settings -> Pages** en tu repositorio de GitHub.
2. En **Build and deployment / Source**, selecciona `Deploy from a branch`.
3. Selecciona la rama `main` y guarda.
4. En 1-2 minutos tu portal estará en vivo en:  
   `https://TU_USUARIO.github.io/modelo-metadatos-pgae-sda/`

---

## 📄 Estructura de Archivos del Repositorio
```
.
├── index.html                      # Portal Web Interactivo (SPA)
├── generar_dashboard.py            # Generador automático Python desde Excel
├── Modelo_Metadatos_PGAE_SDA-vf.xlsx # Modelo fuente de metadatos original en Excel
└── README.md                       # Documentación y guía de despliegue
```

---
*Alcaldía Mayor de Bogotá D.C. — Secretaría Distrital de Ambiente (SDA)*  
*Subdirección de Ecourbanismo y Gestión Ambiental Empresarial / Equipo Análisis Sectorial y POMCAS*
