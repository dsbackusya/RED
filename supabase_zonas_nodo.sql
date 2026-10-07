-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Espacio para el futuro apartado de Zonas: una fila por nodo con su zona (columnas NODO y ZONA).
-- Todavía no lo usa la app; queda listo para armar ese apartado cuando se decida.
-- Se llena con lo que ya existe: primero la zona de los registros (la más reciente de cada nodo) y luego el respaldo de Agencias.

create table if not exists zonas_nodo (
  nodo text primary key,
  zona text not null
);

alter table zonas_nodo enable row level security;
drop policy if exists "auth zonas_nodo" on zonas_nodo;
create policy "auth zonas_nodo" on zonas_nodo for all to anon, authenticated using (true) with check (true);

-- 1) zona más reciente de cada nodo en los registros
insert into zonas_nodo (nodo, zona)
select distinct on (nodo) nodo, zona
from envios
where zona is not null and trim(zona) <> ''
order by nodo, fecha desc, id
on conflict (nodo) do nothing;

-- 2) nodos del respaldo de Agencias que aún no están (HUANCAVELICA, JULIACA, PUNO, CAÑETE, TALARA…)
insert into zonas_nodo (nodo, zona)
select distinct on (n) n, zona
from (
  select regexp_replace(trim(translate(upper(destino), 'ÁÉÍÓÚÜÑ', 'AEIOUUN')), '\s+', ' ', 'g') as n, zona, motivo
  from agencias
  where zona is not null and trim(zona) <> ''
) a
order by n, (motivo = 'DESPACHO') desc
on conflict (nodo) do nothing;

notify pgrst, 'reload schema';
