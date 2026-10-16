# Aluwal · Sistema

Apps internas de Aluwal Aberturas conectadas por un mismo número de presupuesto (B-0001, B-0002...).

- `presupuestador/` presupuestador interno (+ calculador a medida)
- `online/` presupuestador para clientes
- `obra/` vidrios, cortes de perfiles y optimizador de barras
- `ventas/` control de ventas y seguimiento (medidas, producción, entrega)
- `hub/Code.gs` backend (Google Apps Script) que reparte los números y guarda los presupuestos

Los datos de ventas y obras están en Firebase (`sis_ventas`, `sis_obras`).
