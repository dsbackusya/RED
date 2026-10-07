-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Sin la pestaña Agencias: la zona pasa a guardarse en cada registro (se carga desde el Excel del despacho).
-- Los registros que ya existen reciben la zona que tenían en Agencias, para no perder el análisis por zona.
-- Las tablas de Agencias NO se borran: quedan como respaldo.

alter table envios add column if not exists zona text;

update envios e
set zona = coalesce(
  (select a.zona from agencias a
    where a.motivo = coalesce(e.motivo, 'DESPACHO')
      and regexp_replace(trim(translate(upper(a.destino), 'ÁÉÍÓÚÜÑ', 'AEIOUUN')), '\s+', ' ', 'g') = e.nodo
    limit 1),
  (select a.zona from agencias a
    where a.motivo = 'DESPACHO'
      and regexp_replace(trim(translate(upper(a.destino), 'ÁÉÍÓÚÜÑ', 'AEIOUUN')), '\s+', ' ', 'g') = e.nodo
    limit 1))
where e.zona is null;

notify pgrst, 'reload schema';
