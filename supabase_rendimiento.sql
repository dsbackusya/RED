-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Rendimiento: índices por fecha para que Registros, Costos y Liquidaciones sigan siendo rápidos al crecer los datos.

create index if not exists envios_fecha_idx on envios (fecha);
create index if not exists envios_transporte_fecha_idx on envios (transporte, fecha);
create index if not exists envios_modo_pago_fecha_idx on envios (modo_pago, fecha);

-- El HRN ya tiene su índice (fecha_reporte, nodo); se actualizan las estadísticas de las tablas
analyze envios;
analyze hrn_detalle;
