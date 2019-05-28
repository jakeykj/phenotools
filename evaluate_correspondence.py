import numpy as np
from sklearn.preprocessing import normalize


def interpret_corres_matrix(corrs, row_mapping, col_mapping, col_normalize=True):
    """Function to interpret the correspondence matrix. The correspondence are interpreted column-wisely. If
    the correspondence matrix should be interpreted row-wisely, it should be transposed before calling this
    function.

    Args:
        corrs (np.ndarray): The correspondence matrix to be interpreted.
        row_mapping (dict): The idx to name mapping of each row.
        col_mapping (dict): The idx to name mapping of each column.
        col_normalize (bool): Indicate whether to normalize the columns, default is True.

    Returns:
        The correspondence of the row concepts to each column concept.
    """

    if col_normalize:
        corrs = normalize(corrs, axis=0, norm='l1')
    sorts = np.argsort(-corrs, axis=0)
    col_names = np.array([row_mapping[i] for i in range(len(row_mapping))])
    cols = []
    for i in range(corrs.shape[1]):
        col = list(map(lambda x: x if x[1] > 0 else None, zip(col_names[sorts[:, i]], corrs[sorts[:, i], i])))
        cols.append((col_mapping[i], col))
    return cols


def extract_diag_med_corres(D, Us, Ud, Um, dxmap_idx2desc, rxmap_idx2desc):
    corrs = {}
    for dx_idx in range(Ud.shape[0]):
        pts = D[:, dx_idx].nonzero()[0]
        pt_factor = Us[pts, :].sum(axis=0)
        corrs_pt = Um @ np.diag(pt_factor) @ Ud[dx_idx, :].reshape(-1, 1)
        corrs_pt = corrs_pt / corrs_pt.sum()
        
        nnz_idx = corrs_pt.nonzero()[0]
        corrs_pt = [(rxmap_idx2desc[i], corrs_pt[i][0]) for i in nnz_idx]
        corrs_pt = sorted(corrs_pt, key=lambda x: x[1], reverse=True)
        
        corrs[dxmap_idx2desc[dx_idx]] = corrs_pt
    
    dx_count = D.sum(axis=0)
    dx_order = sorted(range(Ud.shape[0]), key=lambda x: dx_count[x], reverse=True)
    corrs = [(dxmap_idx2desc[i], corrs[dxmap_idx2desc[i]]) for i in dx_order]
    return corrs


