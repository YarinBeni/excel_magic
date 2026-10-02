Warning: You are sending unauthenticated requests to the HF Hub. Please set a HF_TOKEN to enable higher rate limits and faster downloads.
# RelBench rel-hm probe

relbench 3.0.1
Fetching 18 files:   0%|          | 0/18 [00:00<?, ?it/s]Fetching 18 files:   6%|▌         | 1/18 [00:00<00:09,  1.83it/s]Fetching 18 files:  11%|█         | 2/18 [00:02<00:24,  1.54s/it]Fetching 18 files:  17%|█▋        | 3/18 [00:04<00:23,  1.57s/it]Fetching 18 files:  94%|█████████▍| 17/18 [00:04<00:00,  6.14it/s]Fetching 18 files: 100%|██████████| 18/18 [00:04<00:00,  3.99it/s]
dataset: RelBenchDataset('rel-hm')
task: ForecastRecommendationTask('user-item-purchase', dataset=RelBenchDataset('rel-hm')) 

- transactions: 15453651 rows, cols=['t_dat', 'customer_id', 'article_id', 'price', 'sales_channel_id'], pkey=None, fkeys={'customer_id': 'customer', 'article_id': 'article'}, time=t_dat
- article: 105542 rows, cols=['article_id', 'product_code', 'prod_name', 'product_type_no', 'product_type_name', 'product_group_name', 'graphical_appearance_no', 'graphical_appearance_name', 'colour_group_code', 'colour_group_name', 'perceived_colour_value_id', 'perceived_colour_value_name'], pkey=article_id, fkeys={}, time=None
- customer: 1371980 rows, cols=['customer_id', 'FN', 'Active', 'club_member_status', 'fashion_news_frequency', 'age', 'postal_code'], pkey=customer_id, fkeys={}, time=None
- task table train: 3878451 rows, cols=['timestamp', 'customer_id', 'article_id']
- task table val: 74575 rows, cols=['timestamp', 'customer_id', 'article_id']
- task table test: 67144 rows, cols=['timestamp', 'customer_id', 'article_id']

GlobalPopularity val: {'map': np.float64(0.003417955724543488)}
PastVisit val: {'map': np.float64(0.01896581118155737)}
