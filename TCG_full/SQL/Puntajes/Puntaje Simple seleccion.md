

```sql
--INSERT INTO tcg_l1_v3.score_mapping (company_id, survey_id, question_id, score_mapping_type, answer_value, answer_score, question_total_score, valid_from, valid_until, is_active, updated_at, operator_type)
-- Todas las combinaciones:
--   Todas las preguntas que estén cargadas en csv_questions
--   Las respuestas 1, 2, 3 con su respectivo puntaje
SELECT 
 142 AS company_id --cambiar al company correspondiente
 , q.survey_id 
 , q.question_id
 , 0 AS score_mapping_type --valor fijo
 , n.answer_value
 , n.answer_score
 , n.question_total_score
 , '2020-01-01 00:00:00.000' AS valid_from
 , '2050-12-31 00:00:00.000' AS valid_until
 , 1 AS is_active
 , GETDATE() AS updated_at
 , 'equals' AS operator_type
FROM (
 SELECT distinct question_id, quiz_id AS survey_id 
 FROM nike.csv_questions --cambiar al schema/tabla que corresponde
 WHERE tipo = 'simple selection' AND has_score = 1 AND score_type = 'score_mapping' --Asegurarse que las preguntas estén configuradas de esta forma.
) q
CROSS JOIN (
 SELECT 1 AS answer_value, 1 AS answer_score, 1 AS question_total_score --respuesta valor 1 (sí)
 UNION ALL SELECT 2, 0, 1 --respuesta valor 2 (no)
 UNION ALL SELECT 3, 0, 0 --respuesta valor 3 (no aplica)
) n
WHERE TRUE 
 AND NOT EXISTS (
  SELECT 1 FROM tcg_l1_v3.score_mapping p
  WHERE p.question_id = q.question_id
   AND p.answer_value = n.answer_value
  )
;

/*
Sí 1
No 2
No Aplica 3
/* 
```

-Cambiar company_id
-Cambiar schema.
-en el csv_quesitons tiene que tener score_mapping y simple selection.
-==**Verificar que el orden sea Si, No, No aplica -> 1,2,3**==


```sql
 SELECT DISTINCT title, position
FROM question_choices qc 
WHERE qc.question_id IN (,,,)
```

-Poner los question_id