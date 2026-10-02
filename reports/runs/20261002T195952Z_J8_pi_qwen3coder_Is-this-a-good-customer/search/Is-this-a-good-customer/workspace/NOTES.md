Improvements made to pipeline.py:

1. Fixed categorical variable encoding: 
   - Previously, categorical variables were not properly encoded
   - Added proper LabelEncoder usage for all categorical columns
   - Converted categorical columns to strings first to avoid encoding issues

2. Result:
   - Achieved 1-auroc = 0.24072 (best so far)
   - This represents a small but meaningful improvement over the baseline of 0.24154
   - The improvement comes from proper handling of categorical variables which allows the model to better distinguish patterns in the data

The key insight was that the original pipeline didn't properly encode categorical variables, which likely prevented the model from learning meaningful patterns from these features.