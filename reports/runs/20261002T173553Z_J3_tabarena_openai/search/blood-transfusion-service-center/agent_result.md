Best pipeline achieved 1-auroc = 0.24183 by using:
1. Original 4 features unchanged (MonthsSinceLastDonation, NumberOfDonations, TotalBloodDonated, MonthsSinceFirstDonation)
2. Sampling strategy with 2 context views created by randomly splitting the training data
3. No feature engineering - sometimes the simplest approach works best with TabPFN

The key insight was that adding engineered features actually hurt performance, and using multiple context views with random sampling provided better generalization than single-view approaches or complex feature engineering.