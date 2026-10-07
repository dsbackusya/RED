-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Agencias separadas por motivo: una lista para DESPACHO y otra para RECOJO.
-- Lo que ya hay queda como DESPACHO. La lista de RECOJO empieza vacía (en la app hay un botón para copiar la de despacho).

alter table agencias add column if not exists motivo text not null default 'DESPACHO';
alter table agencias drop constraint if exists agencias_motivo_check;
alter table agencias add constraint agencias_motivo_check check (motivo in ('DESPACHO','RECOJO'));

-- la clave pasa de (destino) a (destino, motivo): un mismo destino puede estar en las dos listas
do $$
declare c text;
begin
  select conname into c from pg_constraint where conrelid = 'public.agencias'::regclass and contype = 'p';
  if c is not null and c <> 'agencias_destino_motivo_pkey' then
    execute format('alter table agencias drop constraint %I', c);
  end if;
  if not exists (select 1 from pg_constraint where conrelid = 'public.agencias'::regclass and conname = 'agencias_destino_motivo_pkey') then
    alter table agencias add constraint agencias_destino_motivo_pkey primary key (destino, motivo);
  end if;
end $$;

notify pgrst, 'reload schema';
