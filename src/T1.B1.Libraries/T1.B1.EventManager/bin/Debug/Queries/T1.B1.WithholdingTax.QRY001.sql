select
"Selected" as 'Seleccionar',
"DocEntry" as 'Llave Interna',
"DocNum" as 'Número de Documento',
"CardCode" as 'Socio de Negocio',
"Tipo" as 'Tipo',
sum(BaseAmnt) as 'Base'
from(
-- Facturas sin NC
select
DocEntry,DocNum,'Facturas' as Tipo
,BaseAmnt
from OINV
where DocDate >= convert(datetime,'[--StartDate--]',112) and DocDate <= convert(datetime,'[--EndDate--]',112)
and DocEntry not in (
select 
BaseEntry
from RIN1
where DocDate >= convert(datetime,'[--StartDate--]',112) and DocDate <= convert(datetime,'[--EndDate--]',112)
and BaseType = 13
)
union all
-- Facturas con NC en el mismo periodo fiscal
select distinct OINV.DocEntry, OINV.DOcNUm,'Facturas con Nota Credito periodo actual', (OINV.BaseAmnt - ORIN.BaseAmnt) from
OINV
inner join RIN1 on RIN1.BaseEntry = OINV.DocEntry and RIN1.BaseType ='13'
inner join ORIN on RIN1.DocEntry = ORIN.DocEntry
where OINV.DocDate >= convert(datetime,'[--StartDate--]',112) and OINV.DocDate <= convert(datetime,'[--EndDate--]',112)
and ORIN.DocDate <= convert(datetime,'[--EndDate--]',112)

union all
-- Facturas con NC en el mismo periodo fiscal
select distinct OINV.DocEntry, OINV.DOcNUm,'Facturas con Nota Credito periodo futuro', (OINV.BaseAmnt) from
OINV
inner join RIN1 on RIN1.BaseEntry = OINV.DocEntry and RIN1.BaseType ='13'
inner join ORIN on RIN1.DocEntry = ORIN.DocEntry
where OINV.DocDate >= convert(datetime,'[--StartDate--]',112) and OINV.DocDate <= convert(datetime,'[--EndDate--]',112)
and ORIN.DocDate >= convert(datetime,'[--EndDate--]',112)



union all

-- Notas credito emitidas en el mes que no son de la facturacion del mes
select distinct 
DocEntry, DocNum, 'Notas credito de faturas de periodos anteriores '
,BaseAmnt *-1
 from ORIN
where DocDate >= convert(datetime,'[--StartDate--]',112) and DocDate <= convert(datetime,'[--EndDate--]',112)
and DocEntry not in
(
select 
DocEntry
from RIN1
where DocDate >= convert(datetime,'[--StartDate--]',112) and DocDate <= convert(datetime,'[--EndDate--]',112)
and BaseType = 13
)

) as R
group by Tipo,DocEntry,DocNum
order by Tipo,DocNum