# sistema-automatizado-recibos-inquilinos
Sistema automatizado en Google Apps Script para extracción OCR de recibos PDF (luz/agua), cálculo dinámico de liquidaciones por local e integración de mensajes para WhatsApp.
# Sistema Automatizado de Extracción OCR y Liquidación de Servicios para Inquilinos

Proyecto de automatización *End-to-End* desarrollado con **Google Apps Script** para solucionar la distribución de costos de servicios básicos (luz, agua) y alquiler entre múltiples locales comerciales o inquilinos.

---

## 🎯 Problema resuelto
El cálculo manual de consumos de luz/agua a partir de recibos emitidos por las empresas distribuidoras suele ser un proceso propenso a errores humanos, especialmente al desglosar cargos afectos (IGV), cargos no afectos, lecturas de medidores en fechas distintas y tarifas por kWh. Este proyecto automatiza todo el flujo de trabajo: desde la lectura del recibo en PDF hasta la generación del mensaje final de cobro.

---

## 🛠️ Tecnologías y Herramientas
- **Lenguaje:** JavaScript (ES6) en Google Apps Script.
- **Servicios Integrados:** Google Drive API (Advanced Services), Google Docs OCR Engine, Google Sheets.
- **Procesamiento de Texto:** Expresiones Regulares (RegEx) para parsing de recibos y extracción de importes/fechas.
- **Lógica de Negocio:** Algoritmo de sincronización de lecturas históricas vía Google Forms y redondeo de precisión financiera.

---

## 🚀 Funcionalidades Clave

1. **Lectura e Ingesta OCR Automática:**
   - Detecta y lee recibos en PDF cargados en Google Drive.
   - Extrae automáticamente precios unitarios por kWh, fechas de lectura, totales de agua/luz y desglose de cargos afectos/no afectos.

2. **Cálculo Dinámico y Distribución de Costos:**
   - Asigna prorrateos fijos y proporcionales para el servicio de agua según las reglas de cada local.
   - Aplica fórmulas financieras exactas (Costo de consumo + Proporcional de cargos + IGV + Adicionales) para calcular el monto neto de energía.

3. **Algoritmo de Búsqueda y Sincronización de Lecturas:**
   - Busca en el historial de lecturas (procedentes de Google Forms) las fechas más cercanas de inicio y fin de periodo, ajustando desfases de días en caso de variaciones.

4. **Generación de Reporte y Mensajería:**
   - Vuelca los resultados tabulados en una hoja de resumen.
   - Redondea importes a 1 decimal para precisión de cobro.
   - Genera automáticamente la plantilla de mensaje formateado con sintaxis de WhatsApp para cada inquilino.

---

## 📊 Estructura de Datos
- `Calculadora`: Hoja de paso donde se inyectan los datos extraídos por el OCR.
- `Form_Responses`: Base de datos de lecturas registradas periódicamente.
- `Resumen_Cobro`: Panel principal de salida con importes finales por local y plantilla para WhatsApp.

---

## 👨‍💻 Habilidades Demostradas
- Automatización de procesos de negocio (BPA).
- Manipulación de APIs y servicios avanzados de Google Workspace.
- Parsing estructurado de documentos no estructurados (PDF/OCR).
- Manejo de estructuras de datos y algoritmia aplicada en JavaScript.
