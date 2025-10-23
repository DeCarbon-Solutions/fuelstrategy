import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO, StringIO
import openpyxl
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# --- Page Configuration ---
st.set_page_config(layout="wide", page_title="Fuel Supplier Strategy Optimization Results")

# --- MASTER FUEL LIST ---
MASTER_FUEL_LIST = sorted(list(set([
    'VLSFO', 'Bio-methanol', 'Biodiesel', 'B24', 'Blue Ammonia', 'B30', 'HVO',
    'E-Methanol', 'E-Diesel', 'Bio-ethanol', 'B50', 'Bio-methane', 'E-methane', 'E-Ammonia',
    'Green Hydrogen', 'Blue Hydrogen'
])))


# --- Helper Functions ---
@st.cache_data
def create_excel_template(year=2030):
    output = BytesIO()
    if year == 2030:
        fuel_types_template = ['VLSFO', 'Bio-methanol', 'Biodiesel', 'B24', 'Blue Ammonia', 'B30', 'HVO']
    elif year == 2040:
        fuel_types_template = ['VLSFO', 'Bio-methanol', 'Biodiesel', 'B24', 'Blue Ammonia', 'B30', 'HVO', 'E-Methanol',
                               'E-Diesel', 'Bio-ethanol', 'B50', 'Bio-methane', 'E-methane', 'E-Ammonia']
    else:  # 2050
        fuel_types_template = ['VLSFO', 'Bio-methanol', 'Biodiesel', 'Bio-methane', 'E-methane', 'E-Ammonia',
                               'Blue Ammonia', 'Green Hydrogen', 'Blue Hydrogen', 'E-Methanol', 'E-Diesel',
                               'Bio-ethanol', 'B30', 'B50']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame({'Metric': ['Total Sourcing Cost (Million USD)', 'Carbon Levy/Credit Cost (Million USD)',
                                 'TOTAL ANNUAL COST (Million USD)', 'Resulting Fleet WtW GFI (gCO2eq/MJ)',
                                 'Applicable GFI Threshold (gCO2eq/MJ)', 'Carbon Levy/Credit Rate ($/tonne)'],
                      'Value': [0] * 6}).to_excel(writer, sheet_name='Overall_Summary', index=False)
        pd.DataFrame({'FuelType': fuel_types_template, 'Total_GJ': [0] * len(fuel_types_template),
                      'Percentage_Mix': [0] * len(fuel_types_template)}).to_excel(writer, sheet_name='Fuel_Mix_Summary',
                                                                                  index=False)

        prod_data = {'Pathway': [f'x{i + 1}_' for i in range(7)]}
        for fuel in fuel_types_template:
            prod_data[fuel] = [0] * 7
        pd.DataFrame(prod_data).to_excel(writer, sheet_name='Production_Plan_GJ', index=False)

        proc_data = {'Source': ['External Procurement']}
        for fuel in fuel_types_template:
            proc_data[fuel] = [0]
        pd.DataFrame(proc_data).to_excel(writer, sheet_name='Procurement_Plan_GJ', index=False)

        pd.DataFrame({'Refinery': ['REPLAN', 'REFAP', 'RNEST', 'REDUC', 'RPBC'], 'Usage_GJ': [0] * 5,
                      'Utilization_Percent': [0] * 5}).to_excel(writer, sheet_name='Refinery_Utilization', index=False)

    return output.getvalue()


def parse_single_results_excel(uploaded_file, master_fuel_list):
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        data = {
            'summary': all_sheets['Overall_Summary'], 'mix': all_sheets['Fuel_Mix_Summary'],
            'production': all_sheets['Production_Plan_GJ'], 'procurement': all_sheets['Procurement_Plan_GJ'],
            'utilization': all_sheets['Refinery_Utilization']
        }
        for df_name in ['production', 'procurement']:
            df = data[df_name]
            index_col_name = df.columns[0]
            df = df.set_index(index_col_name)
            data[df_name] = df.reindex(columns=master_fuel_list, fill_value=0).reset_index()
        df_mix = data['mix'].set_index('FuelType')
        data['mix'] = df_mix.reindex(master_fuel_list, fill_value=0).reset_index().rename(columns={'index': 'FuelType'})
        return data
    except KeyError as e:
        st.error(f"Error reading '{uploaded_file.name}': A required sheet is missing. Missing sheet: {e}")
        st.warning(
            "Please ensure the file contains all required sheets: 'Overall_Summary', 'Fuel_Mix_Summary', 'Production_Plan_GJ', 'Procurement_Plan_GJ', 'Refinery_Utilization'.")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred while reading '{uploaded_file.name}': {e}")
        return None


@st.cache_data
def create_gfi_compliance_chart(gfi_results_dict=None):
    gfi_data_string = "Year\tGFI Base\tGFI DC\tGFI_Credit\n2028\t89.568\t77.439\t19\n2029\t87.702\t75.573\t19\n2030\t85.836\t73.707\t19\n2031\t81.7308\t69.6018\t19\n2032\t77.6256\t65.4966\t19\n2033\t73.5204\t61.3914\t19\n2034\t69.4152\t57.2862\t19\n2035\t65.31\t53.181\t19\n2036\t58.779\t46.65\t19\n2037\t52.248\t40.119\t19\n2038\t45.717\t32.655\t19\n2039\t39.186\t27.057\t19\n2040\t32.655\t20.526\t14\n2041\t31.0689\t18.9399\t14\n2042\t29.4828\t17.3538\t14\n2043\t27.8967\t15.7677\t14\n2044\t26.3106\t14.1816\t3\n2045\t24.7245\t12.5955\t3\n2046\t23.1384\t11.0094\t3\n2047\t21.5523\t9.4233\t3\n2048\t19.9662\t7.8372\t3\n2049\t18.3801\t6.2511\t3\n2050\t16.794\t4.665\t3"
    df_gfi = pd.read_csv(StringIO(gfi_data_string), sep='\t')
    fig, ax = plt.subplots(figsize=(12, 7))
    ax.fill_between(df_gfi['Year'], df_gfi['GFI Base'], 95, color='red', alpha=0.3, label='Zone 1: High Penalty')
    ax.fill_between(df_gfi['Year'], df_gfi['GFI DC'], df_gfi['GFI Base'], color='orange', alpha=0.4,
                    label='Zone 2: Low Penalty')
    ax.fill_between(df_gfi['Year'], df_gfi['GFI_Credit'], df_gfi['GFI DC'], color='lightgreen', alpha=0.5,
                    label='Zone 3: Compliant')
    ax.fill_between(df_gfi['Year'], 0, df_gfi['GFI_Credit'], color='darkgreen', alpha=0.6,
                    label='Zone 4: Credit Earning')
    if gfi_results_dict:
        for year, gfi_value in gfi_results_dict.items():
            ax.plot(year, gfi_value, marker='*', markersize=18, color='gold', markeredgecolor='black',
                    label=f'Fleet GFI ({year}): {gfi_value:.2f}', zorder=10)
    ax.set_xlabel("Year");
    ax.set_ylabel("GFI Value (gCO2eq/MJ)");
    ax.set_ylim(0, 95)
    ax.set_xticks(df_gfi['Year']);
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.legend(loc='upper right');
    ax.grid(True, linestyle=':', alpha=0.6);
    plt.tight_layout()
    return fig


# --- App Layout ---
st.title("⛽ Fuel Supplier Optimization Results Visualizer")
st.header("1. Upload Scenario Results")
if 'parsed_data' not in st.session_state: st.session_state.parsed_data = {}

years_to_analyze = st.multiselect('Select Years to Analyze:', [2030, 2040, 2050], default=[2030])

if years_to_analyze:
    uploaded_files = st.file_uploader(
        f"Upload one Excel file for each selected year. **Name files with the year** (e.g., 'results_2030.xlsx').",
        type=['xlsx'], accept_multiple_files=True, key='file_uploader')
    if uploaded_files:
        current_filenames = {f.name for f in uploaded_files}
        if current_filenames != st.session_state.get('last_uploaded_filenames', set()):
            st.session_state.parsed_data = {}
            for file in uploaded_files:
                year_in_filename = [year for year in years_to_analyze if str(year) in file.name]
                if year_in_filename:
                    parsed = parse_single_results_excel(file, MASTER_FUEL_LIST)
                    if parsed: st.session_state.parsed_data[year_in_filename[0]] = parsed
                else:
                    st.warning(
                        f"Could not determine year for file '{file.name}'. Please include the year (2030, 2040, or 2050) in the filename.")
            st.session_state.last_uploaded_filenames = current_filenames
            if st.session_state.parsed_data: st.success(
                f"Successfully parsed results for: **{', '.join(map(str, sorted(st.session_state.parsed_data.keys())))}**")
else:
    st.info("Please select one or more years to begin.")

st.markdown("---")
st.subheader("Download Templates")
col1, col2, col3 = st.columns(3)
with col1:
    st.download_button(label="📥 Download Template for 2030", data=create_excel_template(2030),
                       file_name="matlab_results_template_2030.xlsx")
with col2:
    st.download_button(label="📥 Download Template for 2040", data=create_excel_template(2040),
                       file_name="matlab_results_template_2040.xlsx")
with col3:
    st.download_button(label="📥 Download Template for 2050", data=create_excel_template(2050),
                       file_name="matlab_results_template_2050.xlsx")

# --- Section 2: Results Visualization ---
if st.session_state.parsed_data:
    st.divider();
    st.header("📊 Comparative Results Dashboard")
    gfi_results_to_plot = {year: data['summary'].set_index('Metric').loc['Resulting Fleet WtW GFI (gCO2eq/MJ)', 'Value']
                           for year, data in st.session_state.parsed_data.items()}
    st.pyplot(create_gfi_compliance_chart(gfi_results_to_plot));
    st.markdown("---")

    sorted_years = sorted(st.session_state.parsed_data.keys())
    cols = st.columns(len(sorted_years))

    for i, year in enumerate(sorted_years):
        with cols[i]:
            data = st.session_state.parsed_data[year];
            st.subheader(f"Results for {year}")

            summary_df = data['summary'].set_index('Metric')
            st.metric("TOTAL ANNUAL COST", f"${summary_df.loc['TOTAL ANNUAL COST (Million USD)', 'Value']:,.2f}M")
            st.metric("Resulting Fleet WtW GFI",
                      f"{summary_df.loc['Resulting Fleet WtW GFI (gCO2eq/MJ)', 'Value']:.2f} g/MJ")
            st.metric("Carbon Levy Cost", f"${summary_df.loc['Carbon Levy/Credit Cost (Million USD)', 'Value']:,.2f}M")
            with st.expander("View Overall Summary Data"):
                st.dataframe(summary_df)
            st.markdown("---")

            mix_df = data['mix'];
            mix_to_plot = mix_df[mix_df['Percentage_Mix'] > 0.01]
            fig_mix = px.pie(mix_to_plot, names='FuelType', values='Percentage_Mix', hole=0.4, title=f"Fuel Mix (%)")
            fig_mix.update_layout(showlegend=False, margin=dict(t=30, b=0), height=300)
            fig_mix.update_traces(textposition='inside', textinfo='percent+label');
            st.plotly_chart(fig_mix, use_container_width=True)
            with st.expander("View Fuel Mix Data Table"):
                st.dataframe(mix_df.style.format({'Total_GJ': '{:,.0f}', 'Percentage_Mix': '{:.2f}%'}))
            st.markdown("---")

            prod_df = data['production'].set_index(data['production'].columns[0])
            prod_df_melted = prod_df.melt(ignore_index=False, var_name='FuelType', value_name='GJ').reset_index()
            prod_to_plot = prod_df_melted[prod_df_melted['GJ'] > 1];
            fig_prod = px.bar(prod_to_plot, x=prod_df.index.name, y='GJ', color='FuelType',
                              title=f"Production Plan (GJ)")
            fig_prod.update_layout(height=400, xaxis_title=None);
            st.plotly_chart(fig_prod, use_container_width=True)
            with st.expander("View Production Plan Data Table"):
                st.dataframe(prod_df.loc[:, (prod_df.sum(axis=0) > 0)].style.format('{:,.0f}'))
            st.markdown("---")

            # --- FIX: ADDED PROCUREMENT PLAN VISUALIZATION ---
            proc_df = data['procurement'].set_index(data['procurement'].columns[0])
            proc_series = proc_df.iloc[0]  # Get the first (and only) row
            proc_to_plot = proc_series[proc_series > 1]  # Filter out zero/small values

            if not proc_to_plot.empty:
                st.subheader("Procurement Plan (GJ)")
                fig_proc = px.bar(proc_to_plot, x=proc_to_plot.index, y=proc_to_plot.values,
                                  title="External Fuel Procurement",
                                  labels={'x': 'Fuel Type', 'y': 'Volume (GJ)'}, text_auto=True)
                fig_proc.update_layout(height=350, yaxis_title="Procurement (GJ)")
                st.plotly_chart(fig_proc, use_container_width=True)
                with st.expander("View Procurement Plan Data Table"):
                    st.dataframe(proc_df.loc[:, (proc_df.sum(axis=0) > 0)].style.format('{:,.0f}'))
            else:
                st.info("No external procurement required for this scenario.")
            st.markdown("---")

            util_df = data['utilization']
            st.subheader("Refinery Utilization")
            fig_util_gj = px.bar(util_df, x='Refinery', y='Usage_GJ', title=f"Refinery Usage (GJ)", color='Refinery')
            fig_util_gj.update_layout(showlegend=False, height=350, yaxis_title="Production (GJ)");
            st.plotly_chart(fig_util_gj, use_container_width=True)
            with st.expander("View Refinery Utilization Data Table"):
                st.dataframe(util_df.style.format({'Usage_GJ': '{:,.0f}', 'Utilization_Percent': '{:.1f}%'}))

elif years_to_analyze:
    st.info("☝️ Please upload the corresponding Excel file(s) above to see the results dashboard.")