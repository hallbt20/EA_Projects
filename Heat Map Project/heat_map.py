import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.preprocessing import robust_scale


def gen_heat_map_raw_numbers(df):
    # Define the columns for heat maps
    columns = ['Total Potential Growth', 'Sales ($)', 'Marketing Budget Investment', 'Marketing Time Investment']

    # Create subplots
    fig, axes = plt.subplots(1, len(columns), figsize=(16, 5))

    for idx, col in enumerate(columns):
        if col != 'Marketing Time Investment':
            formatted_col = df[col].apply(lambda x: f'${x:,.0f}')
            sns.heatmap(
                df[[col]],
                ax=axes[idx],
                annot=formatted_col.values.reshape(-1, 1),
                fmt='',
                cmap="Greens",
                cbar=False,
                yticklabels=not idx
            )
        else:
            sns.heatmap(
                df[[col]],
                ax=axes[idx],
                annot=True,
                fmt='.0f',
                cmap="Greens",
                cbar=False,
                yticklabels=False
            )

    for i in range(4):
        axes[i].xaxis.tick_top()
        axes[i].xaxis.set_label_position('top')
        axes[i].tick_params(length=0)

        if i:
            axes[i].set_ylabel('')

    # Adjust layout
    plt.tight_layout()
    plt.subplots_adjust(wspace=0, hspace=0)

    plt.show()

    return


def gen_heat_map_ratio_numbers(df):
    # Define the columns for heat maps
    columns = [
        'Total Potential Growth (Ratio)',
        'Sales ($) (Ratio)',
        'Marketing Budget Investment (Ratio)',
        'Marketing Time Investment'
    ]

    # Create subplots
    fig, axes = plt.subplots(1, len(columns), figsize=(20, 5))

    for idx, col in enumerate(columns):
        if col != 'Marketing Time Investment':
            sns.heatmap(
                df[[col]],
                ax=axes[idx],
                annot=True,
                fmt='.1%',
                cmap="Greens",
                cbar=False,
                yticklabels=not idx
            )
        else:
            sns.heatmap(
                df[[col]],
                ax=axes[idx],
                annot=True,
                fmt='.0f',
                cmap="Greens",
                cbar=False,
                yticklabels=False
            )

    for i in range(4):
        axes[i].xaxis.tick_top()
        axes[i].xaxis.set_label_position('top')
        axes[i].tick_params(length=0)

        if i:
            axes[i].set_ylabel('')

    # Adjust layout
    plt.tight_layout()
    plt.subplots_adjust(wspace=0, hspace=0)

    plt.show()


def gen_heat_map_eff_metric(df):
    # Define the columns for heat maps
    columns = [
        'Growth per Budget Dollar',
        'Sales per Budget Dollar',
        'Growth per Time',
        'Sales per Time'
    ]

    # Scatter plot for Growth per Budget Dollar vs Marketing Time Investment
    plt.figure(figsize=(12, 6))

    plt.subplot(1, 2, 1)
    plt.scatter(df['Marketing Time Investment'], df['Growth per Budget Dollar'], color='b')
    plt.xlabel('Marketing Time Investment')
    plt.ylabel('Growth per Budget Dollar')
    plt.title('Growth per Budget Dollar vs Marketing Time Investment')

    # Scatter plot for Sales per Budget Dollar vs Marketing Time Investment
    plt.subplot(1, 2, 2)
    plt.scatter(df['Marketing Time Investment'], df['Sales per Budget Dollar'], color='r')
    plt.xlabel('Marketing Time Investment')
    plt.ylabel('Sales per Budget Dollar')
    plt.title('Sales per Budget Dollar vs Marketing Time Investment')

    plt.tight_layout()
    plt.show()


def scatterplots(df):
    colors = plt.cm.tab20.colors  # Use a colormap for different practices
    practices = df.index.unique()
    color_map = {practice: colors[i % len(colors)] for i, practice in enumerate(practices)}

    # Use fig, axes = plt.subplots to create the scatter plots with a single compact legend to the right
    fig, axes = plt.subplots(2, 3, figsize=(20, 10))

    # Scatter plot 1: Total Potential Growth vs. Marketing Budget Investment
    for practice in practices:
        subset = df.loc[practice]
        axes[0, 0].scatter(subset['Marketing Budget Investment'], subset['Total Potential Growth'],
                           color=color_map[practice], label=practice)
    axes[0, 0].set_title('Total Potential Growth vs. Marketing Budget Investment')
    axes[0, 0].set_xlabel('Marketing Budget Investment')
    axes[0, 0].set_ylabel('Total Potential Growth')

    # Scatter plot 2: Total Potential Growth vs. Marketing Time Investment
    for practice in practices:
        subset = df.loc[practice]
        axes[0, 1].scatter(subset['Marketing Time Investment'], subset['Total Potential Growth'],
                           color=color_map[practice], label=practice)
    axes[0, 1].set_title('Total Potential Growth vs. Marketing Time Investment')
    axes[0, 1].set_xlabel('Marketing Time Investment')
    axes[0, 1].set_ylabel('Total Potential Growth')
    axes[0, 1].xaxis.set_major_locator(plt.MaxNLocator(integer=True))

    # Scatter plot 3: Sales ($) vs. Marketing Budget Investment
    for practice in practices:
        subset = df.loc[practice]
        axes[1, 0].scatter(subset['Marketing Budget Investment'], subset['Sales ($)'], color=color_map[practice],
                           label=practice)
    axes[1, 0].set_title('Sales ($) vs. Marketing Budget Investment')
    axes[1, 0].set_xlabel('Marketing Budget Investment')
    axes[1, 0].set_ylabel('Sales ($)')

    # Scatter plot 4: Sales ($) vs. Marketing Time Investment
    for practice in practices:
        subset = df.loc[practice]
        axes[1, 1].scatter(subset['Marketing Time Investment'], subset['Sales ($)'], color=color_map[practice],
                           label=practice)
    axes[1, 1].set_title('Sales ($) vs. Marketing Time Investment')
    axes[1, 1].set_xlabel('Marketing Time Investment')
    axes[1, 1].set_ylabel('Sales ($)')
    axes[1, 1].xaxis.set_major_locator(plt.MaxNLocator(integer=True))

    axes[0, 2].axis('off')
    axes[1, 2].axis('off')

    # Add a single compact legend to the right
    handles, labels = axes[0, 0].get_legend_handles_labels()
    plt.figlegend(handles, labels, loc=(.7, .33))

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    # Read in Excel file and remove totals row at bottom
    df = pd.read_excel('Heat Map Files/Heat Map Data Inputs_KL.xlsx')
    df = df[:-1]

    # Set 'Practice' as index
    df.set_index('Practice', inplace=True)

    # Generate heat maps for relevant columns
    #gen_heat_map_raw_numbers(df)

    # All relevant columns
    cols = ['Total Potential Growth', 'Sales ($)', 'Marketing Budget Investment', 'Marketing Time Investment']

    # Normalize all relevant columns (i.e., find ratios)
    df['Total Potential Growth (Ratio)'] = df['Total Potential Growth'] / df['Total Potential Growth'].sum()
    df['Sales ($) (Ratio)'] = df['Sales ($)'] / df['Sales ($)'].sum()
    df['Marketing Budget Investment (Ratio)'] = df['Marketing Budget Investment'] / df[
        'Marketing Budget Investment'].sum()
    df['Marketing Time Investment (Ratio)'] = df['Marketing Time Investment'] / df['Marketing Time Investment'].sum()

    # Generate heat maps for ratios of relevant columns
    #gen_heat_map_ratio_numbers(df)

    # Calculations for efficiency metrics
    df['Growth per Budget Dollar'] = df['Total Potential Growth'] / df['Marketing Budget Investment']
    df['Sales per Budget Dollar'] = df['Sales ($)'] / df['Marketing Budget Investment']
    df['Growth per Time'] = df['Total Potential Growth'] / df['Marketing Time Investment']
    df['Sales per Time'] = df['Sales ($)'] / df['Marketing Time Investment']

    # Generate scatter plots for efficiency metrics
    #gen_heat_map_eff_metric(df)

    scatterplots(df)
