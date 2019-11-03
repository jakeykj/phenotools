phenotools
==========

Usage
-----
.. code:: python
    factors = [factor_dx, factor_rx]
    item_dicts = [('Diagnosis', dx_idx2desc), ('Medications', 'rx_idx2desc)]
    filepath = './phenotypes.xlsx'
    
    factor_matrices_to_xlsx(factors, item_dicts, filepath, threshold=0.1, highlight=True)

Parameters
~~~~~~~~~~
* factors: iterable
    A list of factor matrices, with each corresponding to one modality of the phenotypes.
* item_dicts: iterable of iterable
    A list of name and index-to-description mapping for each modality.
    Example: ``[('Diagnosis', dx_idx2desc), ('Medications', 'rx_idx2desc)]``
* filepath: str
    the path to save the excel file.
* threshold: iterable
    The threshold value for each modality. Each factor matrix is normalized column-wisely by 
    the :math:`\ell_1` norm, then values below threshold is set to zero. If `None` is given (by default), 
    no thresholding is performed after normalization.
* highlight: bool
    If `True`, highlight items that appear frequently in different phenotypes. [Default: `Faslse`]
