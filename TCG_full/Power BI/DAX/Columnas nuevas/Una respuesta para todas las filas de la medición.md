Medida utilizada para tomar una respuesta de pregunta de una medición, para que esa respuesta se asocie a todas las respuestas de la medición.

```D
entidad_sugerida =
CALCULATE(
    MAX(fct_respuestas[answer]),
    FILTER(
        fct_respuestas,
        fct_respuestas[survey_response_id] = EARLIER(fct_respuestas[survey_response_id])
            && fct_respuestas[question_category] = "Entidad-Sugerida"
    )
)
```
