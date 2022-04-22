import numpy as np


def return_frames(df_Master, feature_qty, block_limit, sort_column):
    """
    This function returns sub dataframes back of the master data frame passed to it.
    the function accepts the number of features that will be displayed and the size limit
    of each subframe (use 40). It returns a dictionary of sub-frames e.b. dfs[0], dfs [1], etc.
    you can determine how many subframe were created by len(dfs).
    """
    # Sort the master dataset 1st
    df_Master.sort_values(sort_column, inplace=True, ascending=True)
    # Calculate the number of subframes
    number_of_sub_dfs = round(feature_qty/block_limit)
    if number_of_sub_dfs == 0:
        number_of_sub_dfs = 1
    # Create that number of data_sets
    dfs = np.array_split(df_Master, number_of_sub_dfs)
    return dfs
