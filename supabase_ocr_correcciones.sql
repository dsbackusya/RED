-- Ejecutar en Supabase > SQL Editor, UNA SOLA VEZ.
-- Guarda las correcciones hechas a mano sobre lo que leyó el OCR de las facturas (solo cuando el valor final difiere del leído).
-- Sirve para revisar qué proveedores o formatos fallan y ajustar las reglas de lectura.

create table if not exists public.ocr_correcciones (
  id bigint generated always as identity primary key,
  creado timestamptz not null default now(),
  envio_id uuid,
  fecha date,
  nodo text,
  transporte text,
  campo text not null,          -- 'factura' | 'importe'
  leido text,                   -- lo que devolvió el OCR (vacío si no leyó nada)
  final text,                   -- lo que se guardó
  via_qr boolean default false,
  texto_ocr text                -- texto crudo que vio el OCR (recortado)
);

alter table public.ocr_correcciones enable row level security;
drop policy if exists "solo autenticados ocr_correcciones" on public.ocr_correcciones;
create policy "solo autenticados ocr_correcciones" on public.ocr_correcciones for all to authenticated using (true) with check (true);
revoke all on public.ocr_correcciones from anon;
