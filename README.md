# 📊 Control de Partidas Abiertas SAP - Cuenta EM/RF (MLAR, MMAR, U003 & U365)

Aplicación web desarrollada en **Streamlit** para el análisis, conciliación y control operativo de la cuenta de compensación de compras **EM/RF (Entrada de Mercancías / Recepción de Facturas - Cuenta `2101011001`)** en SAP (Transacción **FBL3N**), cruzada con las Órdenes de Compra (OC) y sus fechas comprometidas de entrega.

---

## 💾 Bases de Datos Integradas

La aplicación utiliza como fuentes de datos directas cuatro archivos Excel en la raíz del proyecto:
1. **`PARTIDAS MLAR.xlsx`**: Datos contables y de compras correspondientes a **MLAR**.
2. **`PARTIDAS MMAR.xlsx`**: Datos contables y de compras correspondientes a **MMAR**.
3. **`PARTIDAS U003.xlsx`**: Datos contables y de compras correspondientes a **U003**.
4. **`PARTIDAS U365.xlsx`**: Datos contables y de compras correspondientes a **U365**.

Todas las bases pueden ser analizadas de manera **Consolidada (Todas)** o filtrando específicamente por sociedad (**MLAR**, **MMAR**, **U003** o **U365**) desde el panel lateral.

---

## 🎯 Problema de Negocio que Resuelve

En SAP, la cuenta EM/RF concilia lo que el depósito/planta recibe físicamente (**MIGO / Clase WE / Clave 96 - Haber**) contra lo que el proveedor factura (**MIRO / Clase RE, GR, HR / Clave 86 - Debe**).

Esta herramienta detecta automáticamente:
1. **🔴 URGENTE - Factura Contabilizada con OC Vencida Sin Recepción:** Se cargó o pagó la factura del proveedor, la fecha prometida de entrega ya se cumplió, pero no hay registro de ingreso de mercadería en almacén (riesgo financiero / reclamo a almacén).
2. **🟡 RECLAMO - Mercadería Ingresada con OC Vencida Sin Factura:** El almacén dio ingreso físico a los materiales o servicios, pero el proveedor aún no envió la factura o Compras no la registró (pasivo acumulado).
3. **🟢 COMPENSABLE - Saldo $0,00:** Posiciones que ya concilian de manera exacta (Recepción = Factura), listas para depurar y cerrar en SAP con la transacción **`F.13`**.

---

## 🚀 Funcionalidades Principales en Streamlit

- **Panel de Control y Filtros (Sidebar):**
  - **Selector de Base de Datos:** `Todas (Consolidado)`, `MLAR` o `MMAR`.
  - **Búsqueda Rápida:** Filtrado instantáneo por número de Pedido (OC) o nombre de Proveedor.
  - **Filtros Operativos:** Estado contable EM/RF, Estado de vencimiento, Proveedor, Operador de Compras (`OPERADOR OC`), Operador de Verificación (`OPERADOR DE VA`), Grupo de Compras y Slider de Días de Atraso mínimos.
- **Tabla Directa de Operaciones:**
  - `📑 Todas las Posiciones`: Tabla consolidada a nivel de Pedido y Posición con formato de monedas, fechas y estados (sin pestañas adicionales).
- **Exportación a Excel:** Descarga con un solo clic de un archivo `.xlsx` con las posiciones filtradas y las columnas visibles.

---

## ⚡ Cómo Iniciar la Aplicación

### Opción 1: Con el Lanzador de Windows (Recomendada)
Haz doble clic sobre el archivo:
```text
INICIAR_DASHBOARD.bat
```
El script iniciará el servidor de Streamlit y abrirá automáticamente tu navegador en `http://localhost:8501`.

### Opción 2: Desde la Terminal / Consola
1. Asegúrate de tener instaladas las dependencias:
   ```bash
   pip install -r requirements.txt
   ```
2. Ejecuta Streamlit:
   ```bash
   streamlit run app.py
   ```

---

## 📂 Estructura del Proyecto

```text
├── app.py                         # Aplicación web completa en Streamlit
├── requirements.txt               # Dependencias de Python (streamlit, pandas, openpyxl, plotly)
├── INICIAR_DASHBOARD.bat          # Lanzador automático de un solo clic para Windows
├── PARTIDAS MLAR.xlsx             # Base de datos SAP para MLAR
├── PARTIDAS MMAR.xlsx             # Base de datos SAP para MMAR
├── PARTIDAS U003.xlsx             # Base de datos SAP para U003
├── PARTIDAS U365.xlsx             # Base de datos SAP para U365
├── .streamlit/
│   └── config.toml                # Configuración de tema visual y servidor Streamlit
├── index.html                     # Versión estática autónoma anterior
└── README.md                      # Documentación del proyecto
```
