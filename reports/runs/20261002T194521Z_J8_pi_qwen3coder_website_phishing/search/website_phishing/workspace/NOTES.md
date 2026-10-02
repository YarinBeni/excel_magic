# Improvements Made

The key improvements to the pipeline that led to better performance (best score: 0.25563):

1. **Categorical Encoding**: Converted categorical variables to numeric using LabelEncoder to make them compatible with the tabular model.

2. **Feature Engineering**: Created binary indicator features based on domain knowledge:
   - Presence of suspicious popup windows
   - Presence of phishy SSL states  
   - Presence of suspicious URL lengths
   - Domain age indicators
   - IP address presence indicators

3. **Interaction Features**: Created features that combine multiple characteristics:
   - URL-related feature comparisons (SFH vs Request_URL, etc.)
   - Risk counting features that sum up suspicious indicators
   - High-risk flags when multiple suspicious features are present

4. **Consistent Preprocessing**: Ensured that both training and test data go through the same preprocessing steps by fitting on combined data.

These changes increased the feature count from 9 to 22 while improving the cross-validation logloss from 0.26687 to 0.25563, representing a significant improvement in model performance.