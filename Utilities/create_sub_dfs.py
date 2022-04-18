import pandas as pd
import numpy as np


def return_frames(df_Master, feature_qty, limit):
    """
    This function returns sub dataframes back of the master data frame passed to it.
    the function accepts the number of features that will be displayed and the size limit
    of each subframe (use 40). It returns a disctionary of subframes e.b. dfs[0], dfs [1], etc.
    """
    # Calculate the number of subframes
    number_of_sub_dfs = round(feature_qty/limit)
    # Create that number of data_sets
    dfs = np.array_split(df_Master, number_of_sub_dfs)
    return dfs
