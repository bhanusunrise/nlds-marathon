import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import json  # Import the json module
import plotly.express as px
from streamlit_autorefresh import st_autorefresh

# Loading Data


# time to live - you can change this as it is required to be refereshed
@st.cache_data(ttl=5)
def load_data(sheet_url):
    try:
        data = pd.read_csv(sheet_url)
        return data
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

# Function to create a bar chart based on the specified metric


def create_bar_chart_seperate(df, entity, metric, title):
    filtered_df = df[df['Entity'] == entity]
    fig = px.bar(filtered_df, x='Function', y=metric, title=title, labels={
                 'Function': 'Function', 'Entity': 'Entity', metric: metric}, color='Function')
    return fig

# Function to create a bar chart based on the total points of each entity


def create_bar_chart(entity_sum):
    # Convert entity sum dictionary to DataFrame
    df_entity_sum = pd.DataFrame.from_dict(entity_sum, orient='index')

    # Reset index to make entity a column instead of index
    df_entity_sum.reset_index(inplace=True)
    df_entity_sum.rename(columns={'index': 'Entity'}, inplace=True)

    # Create a bar chart using Plotly Express
    fig = px.bar(df_entity_sum, x='Entity', y='Total', title='Total Score', labels={
                 'Entity': 'Entity', 'Total': 'Total Points'}, color='Entity')

    # Hide the legend
    fig.update_layout(showlegend=False)

    return fig

# Function to calculate the total 'Applied' related to each entity


def calculate_total_applied(df):
    entity_applied_total = {}
    for index, row in df.iterrows():
        entity = row['Entity']
        applied = row['Applied']
        if entity not in entity_applied_total:
            entity_applied_total[entity] = applied
        else:
            entity_applied_total[entity] += applied
    return entity_applied_total

# Function to calculate the total 'Approved' related to each entity


def calculate_total_approved(df):
    entity_approved_total = {}
    for index, row in df.iterrows():
        entity = row['Entity']
        approved = row['Approved']
        if entity not in entity_approved_total:
            entity_approved_total[entity] = approved
        else:
            entity_approved_total[entity] += approved
    return entity_approved_total

# Function to calculate the total points of each entity


def calulate_total_points(df):
    entity_sum = {}
    for index, row in df.iterrows():
        entity = row['Entity']
        total = row['Total']
        if entity not in entity_sum:
            entity_sum[entity] = total
        else:
            entity_sum[entity] += total
    return entity_sum

# Function to calculate the count of 'Applied' related to each entity based on the selected function


def count_applied_by_entity(df, selected_function):
    filtered_df = df[df['Function'] == selected_function]
    applied_counts = filtered_df.groupby(
        'Entity')['Applied'].sum().reset_index()
    applied_counts.rename(columns={'Applied': 'Count_Applied'}, inplace=True)
    return applied_counts

# Function to calculate the count of 'Approved' related to each entity based on the selected function


def count_approved_by_entity(df, selected_function):
    filtered_df = df[df['Function'] == selected_function]
    approved_counts = filtered_df.groupby(
        'Entity')['Approved'].sum().reset_index()
    approved_counts.rename(
        columns={'Approved': 'Count_Approved'}, inplace=True)
    return approved_counts


icon_path = 'https://aiesec.lk/data/dist/images/favicon.png'


def calculate_approval_ranks(df):
    # Sort the DataFrame by 'Total_Approved' column in descending order
    df_sorted = df.sort_values(by='Total_Approved', ascending=False)
    # Add a new column 'Rank' to store the ranks
    df_sorted['Rank'] = range(1, len(df_sorted) + 1)

    return df_sorted


def calculate_ranks_on_score(df):
    # Sort the DataFrame by 'Total' column in descending order
    df_sorted = df.sort_values(by='Total', ascending=False)
    # Add a new column 'Rank' to store the ranks
    df_sorted['Rank'] = range(1, len(df_sorted) + 1)

    return df_sorted


def display_leaderboard(df, total_approvals, total_applications):

    # Rename the column 'Total' to 'Total Points'
    df.rename(columns={'Total': 'Total Points'}, inplace=True)
    df.rename(columns={'Total_Approved': 'Total Approvals'}, inplace=True)
    df.rename(columns={'Total_Applied': 'Total Applications'}, inplace=True)

    # Define a layout with two columns
    col1, col2, col3 = st.columns([3, 1, 1])

    # Display the total approvals in the first column
    with col1:
        st.subheader('Leaderboard')

    # Display the leaderboard in the second column
    with col2:
        # st.metric(label="Total AP Approvals", value=total_approvals)
        # st.button(f"Total AP Approvals : **{total_approvals}**", key="no_action_button")
        st.metric(label="Total Approvals", value=total_approvals)

    with col3:
        st.metric(label="Total Applications", value=total_applications)

    st.dataframe(df.set_index('Rank'), use_container_width=True, height=250)


def display_approval_ranks(df):
    # Calculate ranks based on approvals
    df_with_ranks = calculate_approval_ranks(df)

    # Drop the index column
    df_without_index = df_with_ranks[['Rank', 'Entity', 'Total_Approved']]
    # Rename the column 'Total_Approved' to 'Total Approvals'
    df_with_ranks.rename(
        columns={'Total_Approved': 'Total Approvals'}, inplace=True)

    # Apply gold, silver, and bronze medals to the 'Entity' column
    df_with_ranks['Entity'] = df_with_ranks.apply(lambda row:
                                                  f"🥇 {row['Entity']}" if row['Rank'] == 1 else
                                                  f"🥈 {row['Entity']}" if row['Rank'] == 2 else
                                                  f"🥉 {row['Entity']}" if row['Rank'] == 3 else
                                                  row['Entity'], axis=1)

    # Calculate the total of the 'Total Approvals' column
    tot_ap_approvals = df_with_ranks['Total Approvals'].sum()

    # display the leaderboard section
    display_leaderboard(df_with_ranks, tot_ap_approvals)


def display_score_ranks(df):
    # Calculate ranks based on scores
    df_with_ranks = calculate_ranks_on_score(df)

    # Drop the index column
    df_without_index = df_with_ranks[['Rank', 'Entity', 'Total']]

    # Apply gold, silver, and bronze medals to the 'Entity' column
    df_with_ranks['Entity'] = df_with_ranks.apply(lambda row:
                                                  f"🥇 {row['Entity']}" if row['Rank'] == 1 else
                                                  f"🥈 {row['Entity']}" if row['Rank'] == 2 else
                                                  f"🥉 {row['Entity']}" if row['Rank'] == 3 else
                                                  row['Entity'], axis=1)

    # Calculate the total of the 'Total Points' column
    # tot_points = df_with_ranks['Total Points'].sum()

    # display the leaderboard section
    return df_with_ranks
    # display_leaderboard(df_with_ranks)


def applied_bar_chart_and_data(data):
    # Calculate total 'Applied' related to each entity
    entity_applied_total = calculate_total_applied(data)

    # Convert dictionary to DataFrame
    df_entity_applied_total = pd.DataFrame.from_dict(
        entity_applied_total, orient='index', columns=['Total_Applied'])
    df_entity_applied_total.reset_index(inplace=True)
    df_entity_applied_total.rename(columns={'index': 'Entity'}, inplace=True)

    # Create a colored bar chart using Plotly Express
    fig_applied = px.bar(df_entity_applied_total, x='Entity', y='Total_Applied', title='Total Applications by Entity', labels={
                         'Entity': 'Entity', 'Total_Applied': 'Applications'}, color='Entity')

    # Hide the legend
    fig_applied.update_layout(showlegend=False)

    return fig_applied, df_entity_applied_total


def approved_bar_chart_and_data(data):
    # Calculate total 'Approved' related to each entity
    entity_approved_total = calculate_total_approved(data)

    # Convert dictionary to DataFrame
    df_entity_approved_total = pd.DataFrame.from_dict(
        entity_approved_total, orient='index', columns=['Total_Approved'])
    df_entity_approved_total.reset_index(inplace=True)
    df_entity_approved_total.rename(columns={'index': 'Entity'}, inplace=True)
    # Create a colored bar chart using Plotly Express
    fig_approved = px.bar(df_entity_approved_total, x='Entity', y='Total_Approved', title='Total Approvals by Entity', labels={
                          'Entity': 'Entity', 'Total_Approved': 'Approvals'}, color='Entity')
    # Hide the legend
    fig_approved.update_layout(showlegend=False)

    return fig_approved, df_entity_approved_total


def total_points_and_ranks(data):
    entity_points_total = calulate_total_points(data)
    df_entity_points_total = pd.DataFrame.from_dict(
        entity_points_total, orient='index', columns=['Total'])
    df_entity_points_total.reset_index(inplace=True)
    df_entity_points_total.rename(columns={'index': 'Entity'}, inplace=True)

    # Use the function to display the ranks table
    # display_approval_ranks(df_entity_approved_total)
    df_ranks = display_score_ranks(df_entity_points_total)

    return df_ranks

# Main Streamlit app


def main():
    st.set_page_config(
        layout="centered",
        # You can change the page title here
        page_title="Transcend Hackathon - Dashboard",
        page_icon=icon_path,
    )

    # The Dashboard Title (You can change here)
    st.title("Transcend Hackathon - Dashboard")

    # Set interval to 5 minutes
    st_autorefresh(interval=5 * 60 * 1000, key="data_refresh")
    # URL to your Google Sheets data
    # Datasource url / Google Sheets CSV
    sheet_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTXOP09TlmTmfTCx5x7Dwgm8s80W4z7m9plWqbZ7Lfodxox-26BoTNDq-tozEQylR7jKa3UbtIjU1I1/pub?gid=1562137798&single=true&output=csv"

    # Load data using the cached function
    data = load_data(sheet_url)

    if data is not None:

        # Check if the 'Entity' column exists in the DataFrame
        if 'Entity' in data.columns:

            # calculation of leaderboard items
            fig_applied, df_entity_applied_total = applied_bar_chart_and_data(
                data)
            fig_approved, df_entity_approved_total = approved_bar_chart_and_data(
                data)
            df_ranks = total_points_and_ranks(data)

            # df_combined = pd.concat(df_ranks, df_entity_applied_total, df_entity_approved_total, on='Entity')
            # df_combined = df_ranks.merge(df_entity_applied_total, on='Entity').merge(df_entity_approved_total, on='Entity')
            df_combined = df_entity_applied_total.merge(
                df_entity_approved_total, on='Entity').merge(df_ranks, on='Entity')

            display_leaderboard(df_combined, df_combined['Total_Approved'].sum(
            ), df_combined['Total_Applied'].sum())

            st.plotly_chart(fig_approved, use_container_width=True)
            st.plotly_chart(fig_applied, use_container_width=True)

            ###############################################################################

            st.subheader('Functional Analysis')
            # Create a select box to choose the 'Function'
            selected_function = st.selectbox(
                'Select Function', data['Function'].unique())

            # Get the count of 'Applied' related to each entity based on the selected function
            applied_counts = count_applied_by_entity(data, selected_function)

            # Create a bar chart using Plotly Express
            fig_1 = px.bar(applied_counts, x='Entity', y='Count_Applied', title=f'Applications by Entity for {
                           selected_function} Function', labels={'Entity': 'Entity', 'Count_Applied': 'Applications'}, color='Entity')
            fig_1.update_layout(showlegend=False)

            # Get the count of 'Approved' related to each entity based on the selected function
            approved_counts = count_approved_by_entity(data, selected_function)

            # Create a bar chart using Plotly Express
            fig_2 = px.bar(approved_counts, x='Entity', y='Count_Approved', title=f'Approvals by Entity for {
                           selected_function} Function', labels={'Entity': 'Entity', 'Count_Approved': 'Approvals'}, color='Entity')
            fig_2.update_layout(showlegend=False)

            st.plotly_chart(fig_2, use_container_width=True)
            st.plotly_chart(fig_1, use_container_width=True)

            st.write("<br><br>", unsafe_allow_html=True)
            # Footer - It would be great if you could give us a recognition for the team.
            st.write("<p style='text-align: center;'>Made with ❤️ by &lt;/Dev.Team&gt; of <strong>AIESEC in Sri Lanka</strong></p>", unsafe_allow_html=True)

        else:
            st.error("The 'Entity' column does not exist in the loaded data.")
    else:
        st.error("Failed to load data.")


if __name__ == "__main__":
    main()
