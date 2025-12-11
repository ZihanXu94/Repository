import pandas as pd
import numpy as np
import networkx as nx


def load_data(analysis_type='decline'):
    """
    Load trade data
    analysis_type: 'decline' analyzes import declining countries | 'rise' analyzes import increasing countries
    """
    # Load trade data (maintaining original country order)
    # df_t1 = pd.read_excel('wild_average_2010_2019.xlsx', index_col=0)
    # df_t2 = pd.read_excel('wild_average_2020_2022.xlsx', index_col=0)
    df_t1 = pd.read_excel('grain_average_2010_2019.xlsx', index_col=0)
    df_t2 = pd.read_excel('grain_average_2020_2022.xlsx', index_col=0)
    exporters = df_t1.columns.tolist()  # All exporting countries

    if analysis_type == 'decline':
        # Load import declining countries list and standardize
        # decline_df = pd.read_excel('wild_less_than_zero_countries_short.xlsx', header=None)
        decline_df = pd.read_excel('grain_less_than_zero_countries_full.xlsx', header=None)
        importers = [str(s).strip() for s in decline_df[0] if str(s).strip() in df_t1.index]
    else:  # analysis_type == 'rise'
        # Load import increasing countries list and standardize
        # rise_df = pd.read_excel('wild_greater_than_zero_countries_short.xlsx', header=None)
        rise_df = pd.read_excel('grain_greater_than_zero_countries_full.xlsx', header=None)
        importers = [str(s).strip() for s in rise_df[0] if str(s).strip() in df_t1.index]

    return df_t1, df_t2, exporters, importers


def calculate_export_impact(df_t1, df_t2, importers, exporters, analysis_type='decline'):
    """
    Calculate impact indicators for each exporting country on import changes
    analysis_type: 'decline' analyzes decline impact | 'rise' analyzes rise impact
    """
    # Calculate change matrix (export perspective)
    if analysis_type == 'decline':
        # Analyze decline: calculate gap (where t1 > t2)
        gap = df_t1.loc[importers, exporters] - df_t2.loc[importers, exporters]
        gap[gap < 0] = 0  # Keep only export decline portion
    else:  # analysis_type == 'rise'
        # Analyze rise: calculate increase (where t2 > t1)
        gap = df_t2.loc[importers, exporters] - df_t1.loc[importers, exporters]
        gap[gap < 0] = 0  # Keep only export increase portion

    # ========== Indicator Calculation ==========
    # 1. Direct Contribution (DC)
    total_gap_per_importer = gap.sum(axis=1)
    valid_importers = total_gap_per_importer[total_gap_per_importer > 0].index
    gap = gap.loc[valid_importers]
    dc = gap.div(total_gap_per_importer[valid_importers], axis=0).fillna(0)

    # 2. Global Weight (GW)
    log_gap = np.log1p(total_gap_per_importer[valid_importers])
    gw = (dc.T * log_gap).T.sum(axis=0)

    # 3. Gap Betweenness Centrality (GBC) - Modified key point
    G = nx.DiGraph()
    for imp in gap.index:
        for exp in gap.columns:
            if gap.loc[imp, exp] > 0:
                G.add_edge(imp, exp, weight=1 / (gap.loc[imp, exp] + 1e-9))  # Avoid division by zero

    # Calculate centrality for all nodes then filter for exporters
    all_gbc = nx.betweenness_centrality(G, weight='weight', normalized=True)
    gbc = {k: v for k, v in all_gbc.items() if k in exporters}

    # 4. PageRank Calculation (export network)
    trans_matrix = gap.T.div(gap.sum(axis=1) + 1e-9, axis=0)  # Add smoothing term
    G_pr = nx.from_pandas_adjacency(trans_matrix.T, create_using=nx.DiGraph)
    pr = nx.pagerank(G_pr, max_iter=500)

    return gw.to_dict(), gbc, pr


def entropy_weight(gw, gbc, pr, exporters):
    """Calculate comprehensive weights using entropy weight method"""
    df = pd.DataFrame(index=exporters)
    df['GW'] = pd.Series(gw).reindex(exporters).fillna(0)
    df['GBC'] = pd.Series(gbc).reindex(exporters).fillna(0)
    df['PR'] = pd.Series(pr).reindex(exporters).fillna(0)

    # Normalization
    def normalize(s):
        return (s - s.min()) / (s.max() - s.min() + 1e-9)

    df_norm = df.apply(normalize)

    # Entropy weight calculation
    p = df_norm.div(df_norm.sum(axis=0) + 1e-9, axis=0)
    e = - (p * np.log(p + 1e-9)).sum(axis=0) / np.log(len(df))
    weights = (1 - e) / (1 - e).sum()

    return df, weights


def main(analysis_type='decline'):
    """
    Main function
    analysis_type: 'decline' analyzes import decline | 'rise' analyzes import rise
    """
    # Data loading
    df_t1, df_t2, exporters, importers = load_data(analysis_type)

    # Export country impact indicator calculation
    gw, gbc, pr = calculate_export_impact(df_t1, df_t2, importers, exporters, analysis_type)

    # Entropy weight calculation
    process_df, weights = entropy_weight(gw, gbc, pr, exporters)

    # Comprehensive score
    process_df['Score'] = (process_df['GW'] * weights['GW'] +
                           process_df['GBC'] * weights['GBC'] +
                           process_df['PR'] * weights['PR'])

    # Output file names adjusted based on analysis type
    if analysis_type == 'decline':
        score_file = 'exporters_score_causing_import_decline.xlsx'
        process_file = 'exporters_score_process_causing_import_decline.xlsx'
    else:
        score_file = 'exporters_score_causing_import_rise.xlsx'
        process_file = 'exporters_score_process_causing_import_rise.xlsx'

    # Output results
    process_df['Score'].sort_values(ascending=False).to_excel(
        score_file, index_label='ExportCountry')

    # Process data output (includes three weights)
    weight_df = pd.DataFrame([weights], columns=['GW', 'GBC', 'PR'], index=['Weight'])
    with pd.ExcelWriter(process_file) as writer:
        process_df.to_excel(writer, sheet_name='RawIndicators')
        weight_df.to_excel(writer, sheet_name='Weights')


if __name__ == '__main__':
    # Usage instructions:
    # 1. Analyze import declining countries
    main(analysis_type='decline')

    # 2. Analyze import increasing countries
    # main(analysis_type='rise')
