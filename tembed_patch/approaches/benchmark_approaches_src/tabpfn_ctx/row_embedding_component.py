import numpy as np
import pandas as pd
from benchmark_src.approach_interfaces.row_embedding_interface import RowEmbeddingInterface


class RowEmbeddingComponent(RowEmbeddingInterface):
    def __init__(self, approach_instance):
        self.approach_instance = approach_instance

    def setup_model_for_task(self, input_table: pd.DataFrame, dataset_information: dict):
        pass

    def create_row_embeddings_for_table(self, input_table: pd.DataFrame, train_size: int = None, train_labels: np.ndarray = None):
        return self.approach_instance.get_row_embeddings(input_table)
