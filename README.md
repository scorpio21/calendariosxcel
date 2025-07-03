# 📅 Sistema Integral de Calendario y Turnos 2025

Este repositorio proporciona una solución completa y automatizada en **VBA para Excel** para la gestión anual de calendario, turnos de personal, cálculo de ganancias y generación de informes visuales.

---

## ✨ Características principales

- **📆 Calendario 2025 automatizado**
  - Visualización mensual compacta.
  - Días festivos nacionales, autonómicos y locales resaltados en celeste, con el nombre bajo el número.
  - Domingos destacados en rojo claro.
  - Días no laborables resaltados en azul.
  - Leyenda clara y detallada con explicación de colores y lista de todos los festivos.

- **👥 Gestión avanzada de turnos**
  - Asignación diaria para 5 empleados con reglas personalizadas.
  - Ciclo de turnos configurable y cambio automático en verano (a partir del 28 de julio).
  - Hoja dedicada con turnos diarios, horarios y observaciones.

- **💰 Resumen de ganancias**
  - Cálculo automático de ganancias semanales por empleado y totales.
  - Visualización clara de la información financiera semanal.

- **📊 Informes y gráficos automáticos**
  - Gráfica de turnos por empleado (barras).
  - Gráfica de evolución de ganancias semanales (línea).

- **⚡ Automatización total**
  - Ejecución integral mediante el macro principal `GenerarTodo`, ideal para asociar a un botón en Excel.

---

## 🚀 Instalación y uso

1. **Importa el módulo VBA**
   - Abre Excel y presiona `Alt+F11` para acceder al editor VBA.
   - Ve a `Archivo > Importar archivo...` y selecciona `CalendarioTurnos2025.bas`.

2. **Vincula el macro principal a un botón**
   - Inserta un botón de formulario en cualquier hoja de Excel.
   - Haz clic derecho sobre el botón y selecciona "Asignar macro...".
   - Elige la macro `GenerarTodo`.

3. **Ejecuta**
   - Al pulsar el botón, se generarán automáticamente: el calendario anual, la hoja de turnos, el resumen de ganancias y las gráficas.

---

## ⚙️ Personalización

- **Festivos y empleados:** Puedes editar fácilmente las listas de festivos y los nombres de empleados en el módulo VBA para adaptarlo a tus necesidades.
- **Reglas de turnos:** El sistema está diseñado para ser fácilmente modificable en cuanto a reglas y ciclos de turnos.

---

## 🖥️ Requisitos

- Microsoft Excel para Windows (con soporte para VBA)
- No requiere complementos adicionales

---

## 🙌 Créditos

Desarrollado por [scorpio21](https://github.com/scorpio21), con el apoyo de GitHub Copilot.

---

## 📄 Licencia

Este proyecto se distribuye bajo la licencia MIT. Consulta el archivo [LICENSE.md](LICENSE.md) para más detalles.

---