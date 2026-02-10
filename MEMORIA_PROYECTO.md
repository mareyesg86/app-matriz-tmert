# 📋 Memoria del Proyecto: App Resumen Matriz TMERT

**Autor:** Mauricio Reyes González  
**Fecha de última actualización:** 10 de febrero de 2026  
**Repositorio:** https://github.com/mareyesg86/app-matriz-tmert

---

## 🎯 Objetivo del Proyecto

Crear una aplicación web en Streamlit que analice archivos Excel de la **Matriz TMERT versión 7 de ACHS** y genere un resumen ejecutivo con:
- Información general de la empresa
- Nivel de riesgo máximo por factor
- Conteo de casos por nivel de riesgo
- Validación de cumplimiento de planes de acción
- Alertas de validación para identificar errores o datos incompletos

---

## 🗂️ Estructura del Proyecto

```
matriz tmert/
├── app_resumen_tmert.py      # Aplicación principal
├── requirements.txt          # Dependencias Python
├── .gitignore               # Archivos ignorados por Git
└── MEMORIA_PROYECTO.md      # Este archivo
```

---

## 📊 Estructura del Excel (Matriz TMERT v7 ACHS)

### Hojas del Excel

| Hoja | Nombre/Propósito |
|------|------------------|
| **Hoja 1** | Información general de la empresa |
| **Hoja 2** | (No utilizada en el resumen) |
| **Hoja 3** | Identificación Inicial de Factores de Riesgo |
| **Hoja 4** | Evaluación de Repetitividad |
| **Hoja 5** | Evaluación de Postura |
| **Hoja 6** | Evaluación de MMC - Levantamiento/Descenso/Transporte (LDT) |
| **Hoja 7** | Evaluación de MMC - Empuje/Arrastre (EA) |
| **Hoja 8** | Evaluación de MMP (Manipulación Manual de Personas) |
| **Hoja 9** | Evaluación de Vibración Mano-Brazo (MB) |
| **Hoja 10** | Evaluación de Vibración Cuerpo Completo (CC) |

### Hoja 1: Información General (celdas clave)

| Dato | Celda |
|------|-------|
| Razón Social | D5 |
| RUT Empresa | D6 |
| Nombre Centro de Trabajo | D8 |
| Total Trabajadores | D12 |

### Hoja 3: Identificación Inicial

- **Rango de datos:** Filas 14 a 3013
- **Columna B:** Número de caso/puesto
- **Columnas C y D:** Datos del puesto (vinculados a Hoja 1)
- **Columnas E-K:** Identificación de factores de riesgo (SI/NO)

| Columna | Factor de Riesgo | Hoja de Evaluación |
|---------|-----------------|-------------------|
| E | Repetitividad | Hoja 4 |
| F | Postura | Hoja 5 |
| G | MMC LDT | Hoja 6 |
| H | MMC EA | Hoja 7 |
| I | MMP | Hoja 8 |
| J | Vibración CC | Hoja 10 |
| K | Vibración MB | Hoja 9 |

### Hojas 4-8: Factores con Evaluación en Dos Pasos

Estos factores tienen dos niveles de evaluación:
1. **Primera evaluación (Columna Q o similar):** Resultado "aceptable" o "no aceptable"
2. **Segunda evaluación (Columna X o similar):** Si la primera es "no aceptable", se evalúa si es "intermedio/no crítico" o "crítico"

| Factor | Hoja | Col. 1ª Eval | Col. 2ª Eval | Rango Filas |
|--------|------|-------------|-------------|-------------|
| Repetitividad | 4 | Q (17) | X (24) | 16-116 |
| Postura | 5 | AE (31) | AW (49) | 17-116 |
| MMC LDT | 6 | AG (33) | BD (56) | 18-118 |
| MMC EA | 7 | X (24) | AO (41) | 17-117 |
| MMP | 8 | Y (25) | AO (41) | 17-117 |

### Hojas 9-10: Vibraciones (Evaluación Directa)

Solo tienen un resultado: "aceptable" o "no aceptable" (equivale a crítico).

| Factor | Hoja | Col. Resultado | Rango Filas |
|--------|------|----------------|-------------|
| Vibración MB | 9 | S (19) | 16-116 |
| Vibración CC | 10 | V (22) | 16-116 |

### Flujo de Datos entre Hojas

Los casos/puestos de trabajo fluyen automáticamente entre hojas mediante fórmulas de Excel:

```
Hoja 1 (Datos Base) → Hoja 2 (Lista Puestos) → Hoja 3 (Identificación Inicial)
                                                        ↓
                                            Si hay "SI" en columna E-K
                                                        ↓
                                            Hojas 4-10 (Evaluación por Factor)
```

**IMPORTANTE:** Si las fórmulas en las hojas 4-10 son borradas accidentalmente, los casos no aparecerán aunque tengan "SI" en la Hoja 3.

### Fórmulas de Vinculación (Columna C, Caso 1)

| Hoja | Factor | Fila Caso 1 | Fórmula Col C |
|------|--------|-------------|---------------|
| 4 | Repetitividad | 16 | `=Hoja1!O5` |
| 5 | Postura | 17 | `=Hoja1!W5` |
| 6 | MMC LDT | 18 | `=Hoja1!AE5` |
| 7 | MMC EA | 17 | `=Hoja1!AM5` |
| 8 | MMP | 17 | `=Hoja1!AU5` |
| 9 | Vibración MB | 16 | `=Hoja1!BK5` |
| 10 | Vibración CC | 16 | `=Hoja1!BC5` |

**Patrón de fórmulas:** Para el caso N en la fila F:
- Fórmula: `=Hoja1![COLUMNA][5 + N - 1]`
- Ejemplo: Caso 2 en Hoja 4, fila 17 → `=Hoja1!O6`

---

## ✅ Validaciones Implementadas

### 1. Validación de Identificación Inicial (Hoja 3)

**Función:** `validar_identificacion_inicial()`

**Lógica:** Si las columnas C o D tienen datos (distintos de "0"), entonces las columnas E a K deben contener "SI" o "NO".

**Alerta generada:** "Caso X: Falta completar identificación inicial"

### 2. Validación de Evaluaciones Pendientes

**Función:** `validar_evaluaciones_pendientes()`

**Lógica:** Si en Hoja 3 un factor está marcado como "SI", debe existir el caso en la hoja de evaluación correspondiente.

**Alerta generada:** "Caso X: Marcado SI en [Factor] pero sin evaluación en Hoja Y"

### 3. Validación de Identificación Avanzada Completa

**Función:** `validar_identificacion_avanzada_completa()`

**Lógica:** 
- Si en Hoja 3 hay "SI" para un factor, la evaluación avanzada debe estar completa
- Para factores de 2 pasos: Columna Q debe tener valor; si Q="no aceptable", Columna X también debe tener valor
- Para vibraciones: la columna de resultado debe tener valor

**Alertas generadas:**
- "Caso X: [Factor] - Identificación avanzada incompleta (1ª evaluación vacía)"
- "Caso X: [Factor] - Identificación avanzada incompleta (2ª evaluación vacía)"

### 4. Validación de Planes de Acción

**Función:** `validar_plan_accion_general()`

**Lógica:** Si el nivel de riesgo es CRÍTICO, deben existir datos en las columnas del plan de acción.

**Columnas de Plan de Acción por Factor:**

| Factor | Hoja | Columnas Requeridas |
|--------|------|---------------------|
| Repetitividad | 4 | Y, Z, AB, AC, AD |
| Postura | 5 | AX, AY, BA, BB, BC |
| MMC LDT | 6 | BE, BF, BH, BI, BJ |
| MMC EA | 7 | AP, AQ, AS, AT, AU |
| MMP | 8 | AP, AQ, AS, AT, AU |
| Vibración MB | 9 | T, U, W, X, Y |
| Vibración CC | 10 | W, X, Z, AA, AB |

---

## 🔧 Configuración Técnica

### Dependencias (requirements.txt)

```
streamlit
openpyxl
pandas
python-calamine
xlrd
```

### Estrategia de Lectura de Excel

La aplicación intenta cargar el archivo Excel en este orden para manejar archivos problemáticos:

1. **pandas + calamine** (más robusto)
2. **openpyxl con data_only=True** (lee valores calculados)
3. **openpyxl con data_only=False** (lee fórmulas)
4. **pandas + openpyxl**
5. **pandas + xlrd** (para archivos .xls antiguos)

### Función auxiliar: Acceso a celdas

```python
def get_cell_value(num_hoja, fila, columna):
    # Devuelve el valor de una celda como string
    # Maneja tanto openpyxl (ws.cell) como pandas (DataFrame)
```

---

## 🖥️ Interfaz de Usuario

### Secciones del Resumen Ejecutivo

1. **Información de Empresa:** Razón social, RUT, centro de trabajo, trabajadores
2. **Estado de la Revisión:** Total alertas, casos críticos, casos intermedios, planes de acción
3. **Factores de Riesgo:** Tabla con conteo por nivel (crítico, intermedio, aceptable)
4. **Alertas de Validación:** Expanders con detalle de cada tipo de alerta
5. **Detalle por Factor:** Tabla con nivel máximo y estado del plan de acción
6. **Información General:** Expander con datos completos de la empresa

### Colores por Nivel de Riesgo

| Nivel | Color | Emoji |
|-------|-------|-------|
| CRÍTICO | #ffcccc (rojo claro) | 🔴 |
| INTERMEDIO | #fff3cd (amarillo) | 🟡 |
| ACEPTABLE | #d4edda (verde claro) | 🟢 |
| AUSENTE | #f8f9fa (gris) | ⚪ |

---

## 🚀 Despliegue

### Streamlit Community Cloud

1. Repositorio GitHub: `mareyesg86/app-matriz-tmert`
2. Rama: `main`
3. Archivo principal: `app_resumen_tmert.py`
4. URL de la app: (se genera automáticamente al desplegar)

### Comandos Git útiles

```bash
# Agregar cambios
git add .

# Commit
git commit -m "Descripción del cambio"

# Push a GitHub
git push origin main
```

### Ejecución local

```bash
cd "c:\Users\mauro\matriz tmert"
streamlit run app_resumen_tmert.py
```

---

## 📝 Historial de Cambios Principales

| Fecha | Cambio |
|-------|--------|
| Inicial | Creación de app básica con resumen de factores |
| - | Agregada validación de planes de acción para casos críticos |
| - | Implementada carga robusta de Excel (múltiples engines) |
| - | Agregada validación de identificación inicial (Hoja 3) |
| - | Implementada validación cruzada Hoja 3 vs Hojas de factores |
| - | Corregido conteo de casos por nivel de riesgo |
| - | Mostrar nombres completos de factores en tabla |
| 10/02/2026 | Agregada validación de identificación avanzada completa |
| 10/02/2026 | Corregida fila de inicio de Repetitividad (16 en vez de 14) |
| 10/02/2026 | Documentado flujo de datos y fórmulas de vinculación entre hojas |

---

## 🔮 Posibles Mejoras Futuras

- [ ] Exportar resumen a PDF
- [ ] Agregar gráficos de distribución de riesgos
- [ ] Comparar múltiples matrices TMERT
- [ ] Histórico de evaluaciones por empresa
- [ ] Filtros por factor de riesgo o nivel

---

## 📞 Contacto

**Desarrollador:** Mauricio Reyes González  
**GitHub:** mareyesg86

