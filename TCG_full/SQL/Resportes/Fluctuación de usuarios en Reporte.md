-Cambiar los id de los reportes para mas detalle.

``` sql
SELECT
  CAST(SUBSTRING_INDEX(es.route, '/powerbi-report/', -1) AS UNSIGNED) AS report_id,
  pr.report_name,
  u.full_name,
  u.id as user_id,
  u.email,
  MAX(es.created_at) as fecha_ingreso
FROM endpoint_stats es
JOIN users u ON es.user_id = u.id
JOIN powerbi_reports pr ON pr.id = CAST(SUBSTRING_INDEX(es.route, '/powerbi-report/', -1) AS UNSIGNED)
WHERE
  es.route LIKE '%/powerbi-report/%'
  AND pr.id IN
  (413,141,285,436,419,437,412,169,360,172,433,380,171,178,140,76,206,404,302,156,318,411,407,409,229,
 408,339,375,187,398,397,442,198,348,403,353,421,458,469,467,468)
AND ( MONTH(es.created_at) = '01' AND YEAR(es.created_at) = '2026' )
GROUP BY 1,2,3,4,5
ORDER BY
  6 DESC;
```
