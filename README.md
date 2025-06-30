# Calendario y Turnos 2025

Este repositorio contiene un sistema automatizado en VBA para la gestión de:

- Calendario anual 2025 (con festivos, domingos y leyenda explicativa)
- Turnos de empleados con cambio de ciclo en julio
- Resumen semanal de ganancias por empleado y totales
- Gráficas automáticas de turnos y ganancias
- Todo ejecutable desde un solo botón en Excel

---

## ¿Qué incluye?

1. **Calendario 2025**  
   - Festivos autonómicos/locales destacados en celeste y con nombre bajo el número.
   - Domingos resaltados en rojo claro.
   - Días no laborables en azul.
   - Leyenda completa con explicación de colores y listado de todos los festivos.

2. **Turnos**  
   - Turnos diarios de 5 empleados, con reglas y ciclo que cambia el 28 de julio.
   - Hoja de turnos con detalle diario y observaciones.

3. **Resumen de Ganancias**  
   - Cálculo automático de ganancias semanales por empleado y totales.

4. **Gráficas automáticas**  
   - Turnos por empleado (barras).
   - Ganancias semanales (línea).

5. **Macro principal**  
   - Puedes ejecutar todo con la macro `GenerarTodo`.

---

## Instalación y uso

1. **Importa el módulo VBA**  
   - Abre Excel y pulsa `Alt+F11` para entrar al editor VBA.
   - Menú: `Archivo > Importar archivo...` y selecciona `CalendarioTurnos2025.bas`.

2. **Asigna la macro principal a un botón**  
   - En Excel, inserta un botón de formulario.
   - Haz clic derecho sobre el botón y selecciona "Asignar macro...".
   - Elige la macro `GenerarTodo`.

3. **Pulsa el botón y listo**  
   - El sistema generará todo automáticamente: calendario, turnos, resumen y gráficas.

---

## Personalización

- Puedes modificar la lista de festivos o empleados directamente en el módulo `.bas`.
- Para otros años, solo cambia la lógica de fechas y festivos.

---

## Requisitos

- Microsoft Excel para Windows (VBA habilitado)
- No requiere complementos externos

---

## Créditos

Desarrollado por [scorpio21](https://github.com/scorpio21) con ayuda de GitHub Copilot.

---

## Licencia

MIT License