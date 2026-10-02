Best pipeline uses ratio features (donation_ratio), log transformation of TotalBloodDonated, and interaction terms (NumberOfDonations * MonthsSinceLastDonation). The key improvements were:
1. Creating meaningful ratio features that capture donation frequency relative to time
2. Log-transforming the skewed TotalBloodDonated variable
3. Adding interaction features that capture relationships between variables
These engineered features improved the 1-auroc score from 0.24785 to 0.2457, representing a modest but significant improvement in model performance.