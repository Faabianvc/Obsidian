```sql
--Verificar que no exista el periodo

SELECT *

from crocs.pauta_etl

where periodo = 202608;

  

insert into crocs.pauta_etl (question_id, periodo, value, puntaje, puntaje_total, inicio, fin, created_at, deleted_at, value_text)

SELECT question_id, 202604 as periodo --Periodo a cargar

,value, puntaje, puntaje_total, inicio, fin, current_timestamp::timestamp as created_at, deleted_at, value_text

from crocs.pauta_etl

where periodo = 202603; --Periodo ANTERIOR
```


