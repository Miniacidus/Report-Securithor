# Securithor Report Automator

### Descripción
Herramienta de automatización desarrollada en Python para la optimización de reportes en centros de monitoreo que utilizan el software **Securithor**. Este programa procesa archivos CSV de señales y genera reportes de asistencia y fallas técnicas en formato Excel profesional.

### Problema Solucionado
La generación manual de reportes tomaba varias horas de trabajo administrativo y era propensa a errores humanos. Esta herramienta reduce el tiempo de procesamiento a **segundos**, garantizando precisión en el mapeo de nombres y filtrado de clientes.

### Funciones Principales
- **Mapeo Automático:** Vincula números de cuenta con nombres de clientes mediante un diccionario externo.
- **Filtro de Bajas:** Excluye automáticamente cuentas inactivas.
- **Detección de Fallas:** Identifica clientes sin señales de test (señal 88) y separa los resultados en pestañas de colores.
- **Formato de Impresión:** Genera archivos Excel listos para imprimir (50 filas por hoja, orientación horizontal).

### Tecnologías
- **Lenguaje:** Python 3.x
- **Librerías:** Pandas (Manejo de datos), Tkinter (Interfaz gráfica), OpenPyxl (Formato Excel).
