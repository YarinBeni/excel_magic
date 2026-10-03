# Key Insights from Goal Performance Analysis

## Insight 1: Overall Goal Achievement Rate
The overall goal achievement rate is 23.2%, with 116 out of 499 goals met.
Evidence SQL: 
```sql
SELECT COUNT(*) as total_goals, SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) as goals_met, (SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) * 100.0 / COUNT(*)) as achievement_rate FROM 'data.csv'
```

## Insight 2: Departmental Achievement Rates
IT department has the highest achievement rate at 47.58%, followed by HR (16.80%), Marketing (15.08%), and Finance (13.71%).
Evidence SQL: 
```sql
SELECT department, COUNT(*) as total_goals, SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) as goals_met, (SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) * 100.0 / COUNT(*)) as achievement_rate FROM 'data.csv' GROUP BY department ORDER BY achievement_rate DESC
```

## Insight 3: Priority Level Impact
Critical priority goals have the highest achievement rate at 40.00%, followed by High (27.59%), Medium (12.21%), and Low (8.60%).
Evidence SQL: 
```sql
SELECT priority, COUNT(*) as total_goals, SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) as goals_met, (SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) * 100.0 / COUNT(*)) as achievement_rate FROM 'data.csv' GROUP BY priority ORDER BY achievement_rate DESC
```

## Insight 4: High Achievement Goals Have Higher Completion Rates
Goals that were met had an average percent complete of 89.0% compared to 36.2% for goals that were not met.
Evidence SQL: 
```sql
SELECT AVG(percent_complete) as avg_percent_complete, AVG(target_percentage) as avg_target_percentage FROM 'data.csv' WHERE goal_met = 'True'
```

## Insight 5: Low Achievement Goals Have Lower Completion Rates
Goals that were not met had an average percent complete of 36.2% compared to 89.0% for goals that were met.
Evidence SQL: 
```sql
SELECT AVG(percent_complete) as avg_percent_complete, AVG(target_percentage) as avg_target_percentage FROM 'data.csv' WHERE goal_met = 'False'
```

## Insight 6: IT Department Has Highest Completion Rates
IT department has the highest average percent complete at 63.04% compared to other departments.
Evidence SQL: 
```sql
SELECT department, AVG(percent_complete) as avg_percent_complete, AVG(target_percentage) as avg_target_percentage FROM 'data.csv' GROUP BY department
```

## Insight 7: Critical Priority Goals Show Stronger Achievement
Critical priority goals show the strongest achievement rate at 40.00% compared to other priority levels.
Evidence SQL: 
```sql
SELECT priority, COUNT(*) as total_goals, SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) as goals_met, (SUM(CASE WHEN goal_met = 'True' THEN 1 ELSE 0 END) * 100.0 / COUNT(*)) as achievement_rate FROM 'data.csv' GROUP BY priority ORDER BY achievement_rate DESC
```

## Insight 8: Goal Distribution Across Departments and Priorities
The distribution of goals across departments and priorities shows that IT has the most Critical priority goals (62), while Finance has the most Low priority goals (31).
Evidence SQL: 
```sql
SELECT department, priority, COUNT(*) as goal_count FROM 'data.csv' GROUP BY department, priority
```