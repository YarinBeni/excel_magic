# Pipeline Improvements for Website Phishing Dataset

## Key Improvements Made

1. **Proper Categorical Encoding**: Converted all categorical features to numeric using LabelEncoder, which allows the model to better interpret the relationships between categories.

2. **Feature Engineering**:
   - Created interaction features like `url_anchor_mismatch` to capture inconsistencies between URL types
   - Added binary indicator `long_url_relative_to_age` to flag potentially suspicious URL patterns
   - Developed ratio features such as `url_length_to_age_ratio` to capture relative URL characteristics
   - Implemented count features (`suspicious_count`) to quantify suspicious indicators across multiple attributes

3. **Feature Creation Strategy**:
   - Used domain knowledge about phishing websites to create meaningful features
   - Created binary flags for pattern recognition (long URLs with short domains)
   - Generated ratio features that capture relative relationships between attributes
   - Built interaction features that combine multiple indicators

## Results
- Baseline score: 0.27282
- Best score achieved: 0.27046 (with 13 total features)
- This represents a 0.00236 improvement in logloss over the baseline

The improvements focused on creating meaningful combinations of existing features rather than adding noise, which helped the model better distinguish between legitimate, suspicious, and phishing websites.