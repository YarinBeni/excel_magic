Best pipeline achieved rmse=0.08700. The key improvement was:
1. Creating a Strouhal-like ratio feature: (frequency * chord) / velocity
2. Proper numerical stability with epsilon and clipping
3. Keeping the original features intact
4. No target transformation (which was causing issues in later attempts)

The physics-based feature engineering proved effective for this synthetic aeroacoustic dataset, where the target depends on the Strouhal ratio and related physical quantities.