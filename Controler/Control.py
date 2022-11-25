#%%
import xlwings as xw
import numpy as np
import pandas as pd
#import seaborn as sns
import matplotlib.pyplot as plt
#from PVT_Correlations.model.Functions import Bo
#from PVT_Correlations.model.Functions import Pb
#from PVT_Correlations.model.Functions import Rs
#from PVT_Correlations.Functions import uo

# ------------------------------------
# Create names for sheets
SHEET_SUMMAR = "Datos"
SHEET_RESULTS = "Resultados"

# Name of columns for distribution definitions
VARIABLES = "Variables"
VALORES = "Valores"
PARAMETROS = "Parametros"
CORRELACION = "Correlación"

# Data
STOC_VALUES = "df_bo_calculator"
# Result cells # Call range cells from MS Excel
BO_STANDING = "Bo_STANDING"
BO_AL_MARHOUN = "Bo_Al_Marhoun"
RS_STANDING = "Rs_Standing"
RS_AL_MARHOUN = "Rs_Al_Marhoun"
PB_STANDING = "Pb_Standing"
PB_AL_MARHOUN = "Pb_Al_Marhoun"
UO_BEAL = "uo_Beal"
UO_GLASO = "uo_Glaso"
VALUES = "Valores"
CORRELACION_S = "Standing"
CORRELACION_AL = "AL_MARHOUN"
CORRELACION_B = "Beggs & Robinson"
CORRELACION_G = "Glaso"


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[SHEET_SUMMAR]

    df_TVD = sheet[STOC_VALUES].options(pd.DataFrame, index=False, expand="table").value
    input_col_names = df_TVD["Valores"].to_list()
    Rs_value, Yo_value, Yg_value, T_value, P_value, API_value, BASIO, BASIO2 = tuple(
        input_col_names
    )
    inpt_idx = [CORRELACION_S, CORRELACION_AL]
    resultsBo = {}
    resultsPb = {}
    resultsRs = {}
    for col in inpt_idx:
        print(col)
        resultsBo[col] = Bo(col, Rs_value, Yg_value, Yo_value, T_value)
        resultsPb[col] = Pb(col, Rs_value, Yg_value, T_value, API_value, Yo_value)
        resultsRs[col] = Rs(col, P_value, API_value, T_value, Yg_value, Yo_value)
    inpt_idx2 = [CORRELACION_B, CORRELACION_G]
    resultsuo = {}
    for col2 in inpt_idx2:
        resultsuo[col2] = uo(col2, API_value, T_value)

    PVT_summary_result = [
            resultsBo[CORRELACION_S],
            resultsBo[CORRELACION_AL],
            resultsPb[CORRELACION_S],
            resultsPb[CORRELACION_AL],
            resultsRs[CORRELACION_S],
            resultsRs[CORRELACION_AL],
            resultsuo[CORRELACION_B],
            resultsuo[CORRELACION_G],
    ]
    sheet[BO_STANDING].options(transpose=True).value = PVT_summary_result
    print(PVT_summary_result)


if __name__ == "__main__":
    xw.Book("Control.xlsm").set_mock_caller()
    main()