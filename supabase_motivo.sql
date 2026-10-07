-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Motivo del registro: DESPACHO o RECOJO. Lo ya cargado queda como DESPACHO.
-- Un nodo puede tener un despacho y un recojo el mismo día, por eso la clave única pasa a (fecha, nodo, motivo).

alter table envios add column if not exists motivo text not null default 'DESPACHO';

alter table envios drop constraint if exists envios_motivo_check;
alter table envios add constraint envios_motivo_check check (motivo in ('DESPACHO','RECOJO'));

-- quita la clave única anterior (fecha, nodo), sea cual sea su nombre
do $$
declare c record;
begin
  for c in
    select con.conname
    from pg_constraint con
    where con.conrelid = 'public.envios'::regclass and con.contype = 'u'
      and (select array_agg(att.attname::text order by att.attname::text)
           from unnest(con.conkey) k join pg_attribute att on att.attrelid = con.conrelid and att.attnum = k) = array['fecha','nodo']
  loop
    execute format('alter table envios drop constraint %I', c.conname);
  end loop;
end $$;

alter table envios drop constraint if exists envios_fecha_nodo_motivo_key;
alter table envios add constraint envios_fecha_nodo_motivo_key unique (fecha, nodo, motivo);

create index if not exists envios_motivo_fecha_idx on envios (motivo, fecha);
notify pgrst, 'reload schema';
