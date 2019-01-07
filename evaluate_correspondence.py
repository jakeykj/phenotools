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


