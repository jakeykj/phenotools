from openpyxl import Workbook

from .evaluate_phenotypes import interpret_phenotypes
from .evaluate_correspondence import extract_diag_med_corres
from .xlsx_writer import phenotypes_to_xlsx, phenotypes_to_excel_worksheet, corrspondence_to_xlsx


def factor_matrices_to_xlsx(factors, item_dicts, filepath, threshold=None, highlight=False):
    dim_names, item_idx2desc = zip(*item_dicts)
    phenotypes = interpret_phenotypes(factors, item_idx2desc, dim_names, threshold=threshold)
    phenotypes_to_xlsx(phenotypes, filepath, highlight=False)


def batch_factors_to_xlsx(factors_dict, item_dicts, filepath, threshold=None):
    wb = Workbook()
    wb.remove(wb.active)

    dim_names, item_idx2desc = zip(*item_dicts)

    for name, factors in factors_dict.items():
        ws = wb.create_sheet(name)
        phenotypes = interpret_phenotypes(factors, item_idx2desc, dim_names, threshold=threshold)
        phenotypes_to_excel_worksheet(phenotypes, ws)
    wb.save(filepath)


# def corrs_matrix_to_xlsx(corrs_matrix, row_mapping, col_mapping, filepath, col_normalize=True, ws_name=None):
#     corrs_cols = interpret_corres_matrix(corrs_matrix, row_mapping, col_mapping, col_normalize)
#     corrspondence_to_xlsx(corrs_cols, filepath, ws_name=ws_name)

def extract_corrs_to_xlsx(D, Us, Ud, Um, dxmap_idx2desc, rxmap_idx2desc, filepath, ws_name=None):
    corrs = extract_diag_med_corres(D, Us, Ud, Um, dxmap_idx2desc, rxmap_idx2desc)
    corrspondence_to_xlsx(corrs, filepath, ws_name=ws_name)