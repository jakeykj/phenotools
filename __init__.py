# from .evaluate_phenotypes import interpret_phenotypes
# from .evaluate_correspondence import interpret_corres_matrix
#
# from .xlsx_writer import phenotypes_to_xlsx, corrspondence_to_xlsx
from .utils import factor_matrices_to_xlsx, batch_factors_to_xlsx, extract_corrs_to_xlsx #, corrs_matrix_to_xlsx
from .evaluate_correspondence import extract_diag_med_corres
from .evaluate_phenotypes import sparsity_similarity