# Incident Resolution Time Analysis Insights

## Top 8 Key Insights

1. **Hardware category has the longest average resolution time** at 1237 hours (nearly 52 days), which is significantly higher than other categories.
   Evidence SQL: SELECT category, AVG(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS avg_hours_to_resolve FROM data GROUP BY category ORDER BY avg_hours_to_resolve DESC LIMIT 1

2. **Hardware category shows a dramatic increase in resolution times starting in August 2023**, with average resolution times rising from ~100 days to over 2000 days.
   Evidence SQL: SELECT category, strftime('%Y-%m', CAST(opened_at AS TIMESTAMP)) as month, AVG(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS avg_hours_to_resolve FROM data WHERE category = 'Hardware' GROUP BY category, month ORDER BY month

3. **Hardware category has the highest maximum resolution time** at 2803 hours (about 117 days).
   Evidence SQL: SELECT category, MAX(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS max_hours_to_resolve FROM data GROUP BY category ORDER BY max_hours_to_resolve DESC LIMIT 1

4. **Network category has the second-longest average resolution time** at 188 hours (about 8 days).
   Evidence SQL: SELECT category, AVG(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS avg_hours_to_resolve FROM data GROUP BY category ORDER BY avg_hours_to_resolve DESC LIMIT 1 OFFSET 1

5. **Database and Software categories have similar average resolution times** at around 172 hours (about 7 days).
   Evidence SQL: SELECT category, AVG(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS avg_hours_to_resolve FROM data GROUP BY category ORDER BY avg_hours_to_resolve DESC LIMIT 2 OFFSET 2

6. **Inquiry / Help category has the shortest average resolution time** at 166 hours (about 7 days).
   Evidence SQL: SELECT category, AVG(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS avg_hours_to_resolve FROM data GROUP BY category ORDER BY avg_hours_to_resolve DESC LIMIT 1 OFFSET 4

7. **Hardware category accounts for the largest volume of incidents** with 177 records (29% of total).
   Evidence SQL: SELECT category, COUNT(*) as incident_count FROM data GROUP BY category ORDER BY incident_count DESC LIMIT 1

8. **There's a clear temporal trend showing increasing resolution times for Hardware category** from August 2023 onwards, indicating a systemic issue that needs investigation.
   Evidence SQL: SELECT category, strftime('%Y-%m', CAST(opened_at AS TIMESTAMP)) as month, AVG(CAST(closed_at AS TIMESTAMP) - CAST(opened_at AS TIMESTAMP)) * 24 AS avg_hours_to_resolve FROM data WHERE category = 'Hardware' GROUP BY category, month ORDER BY month

These insights indicate that the Hardware category is the primary concern for increasing resolution times, with a dramatic spike starting in August 2023 that warrants further investigation.