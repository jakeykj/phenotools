from math import ceil
import numpy as np
from sklearn.preprocessing import normalize

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


class Phenotype(object):
    def __init__(self, phenotypes, dim_names, stats):
        self.phenotypes = phenotypes
        self.dim_names = dim_names
        self.stats = stats

    def __getitem__(self, idx):
        return self.phenotypes[idx]

    def __len__(self):
        return len(self.phenotypes)


def interpret_phenotypes(factors, item_idx2desc, dim_names, threshold=None):
    """Function to interpret the non-negative CP factors as phenotypes. The columns of the CP factors
    are first normalized by their l1-norm; the values less than the threshold are then filtered out.

    Args:

        factors (iterable of np.ndarray): CP factor matrices, one matrix for each dimension.
        item_idx2desc (iterable of dict): Concept mapping of the factor matrices. Each one must be
            a {idx: description} dict, and must has the same order with factors.
        dim_names: The name of each dimensions.
        threshold: The hard threshold to filter out items with very small values.

    Returns:
        Phenotype: The phenotype interpretation of the input factors.
    """
    assert len(factors) == len(item_idx2desc), 'Number of factors and the concept mappings must be the same.'
    assert all([factor.min() >= 0 for factor in factors]), 'All CP factors must be non-negative.'

    phenotypes = []
    n_dims = len(factors)
    n_factors = factors[0].shape[1]
    item_sortidx = [np.argsort(-U, axis=0) for U in factors]

    factors = [normalize(factor, axis=0, norm='l1') for factor in factors]

    if threshold is not None:
        for i, factor in enumerate(factors):
            factor[factor < threshold[i]] = 0

    # get stats of phenotype overlapping
    def overlapping_stat(factor):
        positive_flag = factor > 0
        return positive_flag.sum(axis=1)

    stats = {
        'overlap': [overlapping_stat(factor) for factor in factors]
    }

    for r in range(n_factors):
        phenotype_definition = []
        for j in range(n_dims):
            dim_j = []
            for idx in item_sortidx[j][:, r]:
                if factors[j][idx, r] > 0:
                    dim_j.append((idx, item_idx2desc[j][idx], factors[j][idx, r]))
                else:
                    break
            phenotype_definition.append(dim_j)
        phenotypes.append(phenotype_definition)
    return Phenotype(phenotypes, dim_names, stats)



                





