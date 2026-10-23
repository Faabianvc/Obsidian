-Tienen que estar cargadas las tareas nuevas en el csv_quizzes con la columna app 'v3'.

```sql
--v3
INSERT INTO NOMBRE_DEL_SCHEMA.csv_questions (question_id, quiz_id, original_title, tipo, app, orden)
SELECT q.id, q.survey_id, q.title, q."type", 'v3', q."position"
FROM tcg_l0_v3.questions q
JOIN NOMBRE_DEL_SCHEMA.csv_quizzes cqz ON q.survey_id = cqz.quiz_id
  AND cqz.app = 'v3'
WHERE q.deleted = 0
  AND q."type" NOT IN ('welcome page', 'content', 'end page')
  AND q.id NOT IN (SELECT DISTINCT question_id FROM NOMBRE_DEL_SCHEMA.csv_questions) --evita agregar 2 veces la misma pregunta.
```
