import streamlit as st
import pandas as pd
import datetime
import glob
import plotly.express as px
import numpy as np
from io import BytesIO
import base64
import statsmodels.api as sm
from statsmodels.tsa.seasonal import seasonal_decompose



@st.cache()
def col_process(df):
    #Preprocess data
    # drop rows with code is null - e.g. A Team, Commissionary, etc
    df.dropna(0, how='any', inplace=True)
    # remove 'grand total' rows
    df1 = df[(df['Date'] != 'Grand Total') & (df['Store Name'] != 'Total') & (df['Date'] != 'Total All Store')]
    return df1

@st.cache()
def col_total_percent(df): 
    actual_sales = df['Actual sales'].sum()
    total_col = df['Total COL $ (included Holiday and paid leave days)'].sum()
    total_col_percent = 100 * total_col / actual_sales
    return total_col_percent, actual_sales

@st.cache()
def read_col_files():
    all_files = glob.glob("./Data/*.xls")
    col = pd.DataFrame()
    trans = pd.DataFrame()
    for file in all_files:
        print('read: ', file)
        df = pd.read_excel(file, sheet_name='HR Data ', parse_dates=['Date'])
        df_trans = pd.read_excel(file, sheet_name='Sales Data', parse_dates=['date'])
        df = col_process(df)
        col = pd.concat([col, df])
        trans = pd.concat([trans,df_trans])
        # drop duplicates
        col = col.drop_duplicates(subset=['Actual sales','Total COL $ (included Holiday and paid leave days)'])
        trans = trans.drop_duplicates(subset=['date','shopcode','Sale','Trans'])
        # col['Date'] = pd.to_datetime(col['Date'])
        col = col.fillna(0)
    return col, trans

@st.cache
def store_code():
    df = pd.read_excel('./store_info/Tracking store by year.xls', nrows=100,usecols=['Store Code','Title', 'AC', 'Region','Province','Concept','Opening Date'],parse_dates=['Opening Date'])
    df.rename(columns={'Title':'Store'}, inplace=True)
    df['full_name'] = df['Store Code'] + '-' + df['Store']
    return df

@st.cache(allow_output_mutation=True)
def data_filter(df, trans, store_code):
    data_filter = df[df['Code'].isin(store_code)]
    trans_filter = trans[trans['shopcode'].isin(store_code)]
    return data_filter, trans_filter

st.cache()
def filtered_data_merged(df, trans, store):
    filtered_data_merged = df.merge(store,how='left',left_on='Code',right_on='Store Code')
    filtered_data_merged = filtered_data_merged.drop(['Store Code','Store','Opening Date','full_name'], axis=1)
    filtered_data_merged['Date'] = pd.to_datetime(filtered_data_merged['Date'])

    filtered_trans_merged = trans.merge(store,how='left',left_on='shopcode',right_on='Store Code')
    filtered_trans_merged = filtered_trans_merged.drop(['Store Code','Store','Opening Date','full_name'], axis=1)
    filtered_trans_merged['date'] = pd.to_datetime(filtered_trans_merged['date'])

    filtered_data_merged = filtered_data_merged.set_index(['Date','Code'])
    trans['ordertype desc'] = trans['ordertype desc'].str.strip()

    # extract sales/trans by channel and merge to filtered_data_merged dataframe
    for type in trans['ordertype desc'].unique():
        disc = trans[trans['ordertype desc'] == type].set_index(['date','shopcode'])[['Sale','Trans']]
        disc.index.rename(['Date','Code'], inplace=True)

        sales_column = type + ' sales'
        trans_column = type + ' trans'
        disc.columns = [[sales_column, trans_column]]

        filtered_data_merged = filtered_data_merged.merge(disc, left_index=True, right_index=True, suffixes=('_left','_right'))
    filtered_data_merged.rename(columns=''.join,inplace=True)

    filtered_data_merged.reset_index(inplace=True)
    # rename columns COL%, COL$ and COL.1% to forecast COL%, forecast COL$ and actual COL%
    filtered_data_merged = filtered_data_merged.rename(columns={'COL $':'Forecast COL $','COL %':'Forecast COL %','COL %.1':'Actual COL %'})

    # Assign column types
    filtered_data_merged = filtered_data_merged.astype({
            'Code': 'string',
            'Store Name': 'string',
            'Forecast Sales': 'int64',
            'Forecast Hours': 'int64',
            'Forecast COL $': 'float64',
            'Forecast COL %': 'float64',
            'Actual sales': 'float64',
            'Total actual hours (included Holiday and paid leave days)':'int64',
            'Actual hours of MNGT': 'int64',
            'Actual hours of TMs Full time': 'int64',
            'Actual hours of TMs Part time': 'int64',
            'Total actual hours (excluded Holiday and paid leave days)': 'int64',
            'Hours of holiday/paid leave days': 'int64',
            'Total COL $ (included Holiday and paid leave days)':'float64',
            'COL $  of TM Full time':'float64',
            'COL $  of TM Part time':'float64',
            'COL $ Management':'float64',
            'Total COL (excluded Holiday and paid leave days)':'float64',
            'COL of holidays/paid leave days':'float64',
            'Actual COL %':'float64',
            'COL Val %':'float64',
            'Working hour in work shift':'int64',
            'Over time in normal day':'int64',
            'Over time in weekend':'int64',
            'Over time in holiday': 'int64',
            'working hour in normal day (Night)': 'int64',
            'Over time in normal day (Night)':'int64',
            'Over time in weekend (Night)':'int64',
            'Over time in holiday (Night)':'int64',
            '13th salary':'float64',
            'BSC bonus':'float64',
            'PA bonus':'float64',
            'Meal Allowance':'float64',
            'Insurance contribution Amount per day':'float64'
        })

    # Add extra columns Total Transaction, SPMH and TPMH
    filtered_data_merged['Total Transaction'] = filtered_data_merged['Pickup trans']+filtered_data_merged['Dinein trans']+filtered_data_merged['Delivery trans']
    if filtered_data_merged['Forecast Hours'].empty:
        filtered_data_merged['Forecast SPMH'] = 0
    else: 
        filtered_data_merged['Forecast SPMH'] = filtered_data_merged['Forecast Sales'] / filtered_data_merged['Forecast Hours']
    if filtered_data_merged['Total actual hours (included Holiday and paid leave days)'].empty:
        filtered_data_merged['Actual SPMH'] = 0
    else:
        filtered_data_merged['Actual SPMH'] = filtered_data_merged['Actual sales'] / filtered_data_merged['Total actual hours (included Holiday and paid leave days)']
        filtered_data_merged['Actual TPMH'] = filtered_data_merged['Total Transaction'] / filtered_data_merged['Total actual hours (included Holiday and paid leave days)']

    # Add MAPE
    def MAPE(actual, forecast):
        mape = 100* (abs(actual - forecast) / actual)
        mape.replace(np.inf,0, inplace=True)
        mape = mape.fillna(0)
        return mape

    filtered_data_merged['Sales MAPE'] = MAPE(filtered_data_merged['Actual sales'], filtered_data_merged['Forecast Sales'])

    filtered_data_merged['Labour MAPE'] = MAPE(filtered_data_merged['Total actual hours (included Holiday and paid leave days)'],filtered_data_merged['Forecast Hours'])


    # Add COL% Variance
    filtered_data_merged['COL% Variance'] = (filtered_data_merged['Actual COL %'] - filtered_data_merged['Forecast COL %']) / filtered_data_merged['Actual COL %']
    filtered_data_merged['COL% MAPE'] = MAPE(filtered_data_merged['Actual COL %'], filtered_data_merged['Forecast COL %'])
    filtered_data_merged['TA'] = filtered_data_merged['Actual sales'].div(filtered_data_merged['Total Transaction'])
    filtered_data_merged = filtered_data_merged.fillna(0)
    return filtered_data_merged

def to_excel(df):

    # Set up Excel file writer
    output = BytesIO()
    writer = pd.ExcelWriter(output)

    # Write each dataframe to a different worksheet.
    df.to_excel(writer, sheet_name='Summary')

    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df):
    val = to_excel(df)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="extract.xlsx">Download file</a>' # decode b'abc' => abc

@st.cache()
def load_baseline():
    # import baseline COL data and merge with current COL Dataframe
    bl_df = pd.read_excel('./Data/Baseline/COL Baseline - Jun 20.xls', sheet_name='Base line JUN 2019', header=0)
    return bl_df

def regression_table(fig, breakdown):
    regression_df = px.get_trendline_results(fig)
    regression_df = regression_df.set_index(breakdown)
    m=[]
    y=[]
    r2=[]
    for store in regression_df.index:
        _m = regression_df.loc[store].px_fit_results.params[1]
        m.append(_m)
        _y = regression_df.loc[store].px_fit_results.params[0]
        y.append(_y)
        _r2 = regression_df.loc[store].px_fit_results.rsquared
        r2.append(_r2)
    regression_df['Gradient'] = m
    regression_df['y-intercept'] = y
    regression_df['R2 Score'] = r2
    # st.dataframe(regression_df[['Gradient','y-intercept','R2 Score']].sort_values(breakdown))
    regression_df.dropna(inplace=True)
    return regression_df

def regression_plot(filtered_data_merged, sidebar_name, x_axis, y_axis):
    # Plot chart with all the settings

    if len(x_axis) > 0 and len(y_axis) > 0:
        with st.sidebar.beta_expander(sidebar_name, expanded=True):
            # Data grouping and splitplot
            breakdown_option = [None, 'AC','Region','Province','Concept','Store Name','Code']
            breakdown = st.selectbox('Data grouping', breakdown_option)
            split_plot = st.selectbox('Split plot by', breakdown_option)

            # Check box for add trendline
            if st.checkbox('Add trend line', value=True):
                trendline = "ols"
            else:
                trendline = None

            # Remove 0 values in the dataset to plot more accurate regression line
            if st.checkbox('Remove 0 values'):
                filtered_data_merged = filtered_data_merged[(filtered_data_merged[x_axis] > 0) & (filtered_data_merged[y_axis] > 0)]
            else:
                # Define 0 value
                zero_value_x_axis = filtered_data_merged[filtered_data_merged[x_axis] == 0][x_axis]
                zero_value_y_axis = filtered_data_merged[filtered_data_merged[y_axis] == 0][y_axis]

                st.subheader('No. of 0 values at x-axis: ')
                st.info(zero_value_x_axis.count())
                st.subheader('No. of 0 values at y-axis:')
                st.info(zero_value_y_axis.count())

            # Add boxplot for x-axis and y-axis
            if st.checkbox('Add boxplots for axes'):
                marginal = "box"
            else:
                marginal = None
                
        fig = px.scatter(filtered_data_merged,x=x_axis,y=y_axis,color=breakdown,opacity=0.5,facet_col=split_plot, facet_col_wrap=2,trendline=trendline,marginal_x=marginal, marginal_y=marginal, height=600).update_layout(autosize=True)
        fig.update_yaxes(rangemode="tozero")
        fig.update_xaxes(rangemode='tozero')
        st.plotly_chart(fig, use_container_width=True)

        if breakdown == None:
            if trendline:
                # Add trendline with variables
                regression_df = px.get_trendline_results(fig)
                st.subheader('Properties of regression line')
                st.write('Gradient:')
                st.info(round(regression_df.iloc[0].px_fit_results.params[1],2))
                st.write('y-intercept:')
                st.info(round(regression_df.iloc[0].px_fit_results.params[0],2))
                st.write('R2 Score')
                st.info(round(regression_df.iloc[0].px_fit_results.rsquared,2))
        else:
            if trendline:
                st.subheader('Properties of regression line')
                # Add trendline by stores with table
                regression_df = regression_table(fig, breakdown)
                st.dataframe(regression_df)
                st.subheader('Gradient-R2 Score Scatterplot')
                y_r2_plot = px.scatter(regression_df,x='Gradient',y='y-intercept', size='R2 Score',color=regression_df.index)
                y_r2_plot.add_hline(y=regression_df['y-intercept'].mean())
                y_r2_plot.add_vline(x=regression_df['Gradient'].mean())
                st.plotly_chart(y_r2_plot, use_container_width=True)

                # Download Data Section
                st.sidebar.subheader('Download Regression Data')
                st.sidebar.write('Click the link below to download the data for your own use:')
                st.sidebar.markdown(get_table_download_link(regression_df), unsafe_allow_html=True)

def spmh_time_series(data, resample):
    ts_df = filtered_data_merged.set_index('Date')
    spmh_df = ts_df.resample('D').mean()
    spmh_df = spmh_df[['Actual SPMH']]
    spmh_df['SPMH Moving Average'] = spmh_df['Actual SPMH'].ewm(span=7,adjust=False).mean()
    spmh_ts_plot = px.line(spmh_df)

    return spmh_df, spmh_ts_plot

# Expander - Data Filter Setting
with st.sidebar.beta_expander("Data Filter", expanded=True):

    store_code_func = store_code()
    # Store selection options
    store_select_option = st.radio('Select Stores:', ('by Individual/Multiple Stores','by AC','by Concept','by Region','All Stores'), index=4)
    if store_select_option == 'All Stores':
        store_code = store_code_func['Store Code']
        store_name = store_code_func['Store']
    elif store_select_option == 'by Individual/Multiple Stores':
        store_selected = st.multiselect('Select store code to be forecasted:', store_code_func['full_name'])
        store_code = store_code_func[store_code_func['full_name'].isin(store_selected)]['Store Code']
        store_name = store_code_func[store_code_func['full_name'].isin(store_selected)]['Store']
    elif store_select_option == 'by AC':
        ac_selected = st.multiselect('Select AC area to be forecasted:', store_code_func['AC'].unique())
        store_code = store_code_func[store_code_func['AC'].isin(ac_selected)]['Store Code']
        store_name = store_code_func[store_code_func['AC'].isin(ac_selected)]['Store']
    elif store_select_option == 'by Region':
        region_selected = st.selectbox('Select the region:', store_code_func['Region'].unique())
        store_code = store_code_func[store_code_func['Region']== region_selected]['Store Code']
        store_name = store_code_func[store_code_func['Region']== region_selected]['Store']
    elif store_select_option == 'by Concept':
        format_selected = st.selectbox('Select the store format:', store_code_func['Concept'].unique())
        store_code = store_code_func[store_code_func['Concept'] == format_selected]['Store Code']
        store_name = store_code_func[store_code_func['Concept'] == format_selected]['Store']

    data, trans = read_col_files()

    filtered_data, filtered_trans = data_filter(data, trans, store_code)
    filtered_data_merged= filtered_data_merged(filtered_data, filtered_trans, store_code_func)

    # Sidebar sales display
    st.write('Actual sales:')
    sales_data = filtered_data_merged['Actual sales'].groupby(filtered_data_merged.Date).sum()
    sales_data_plot = px.line(sales_data, x=sales_data.index, y='Actual sales')
    sales_data_plot.update_yaxes(matches=None, showticklabels=False, visible=False)
    sales_data_plot.update_xaxes(matches=None, showticklabels=False, visible=False)
    sales_data_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0}, width=304, height=100)
    st.plotly_chart(sales_data_plot, use_column_width=True)

    # Add date slider
    start_date = pd.to_datetime(filtered_data_merged['Date']).min().to_pydatetime()
    check_date = pd.to_datetime(filtered_data_merged['Date']).max().to_pydatetime() - datetime.timedelta(days = 30)
    end_date = pd.to_datetime(filtered_data_merged['Date']).max().to_pydatetime()
    date_range = st.slider(label='Select date range', min_value = start_date, max_value = end_date, value=(check_date, end_date))
    filtered_data_merged = filtered_data_merged.loc[(filtered_data_merged['Date']>=date_range[0]) & (filtered_data_merged['Date'] <= date_range[1])]
    #filtered_trans_merged = filtered_trans_merged.loc[(filtered_trans_merged['date']>=date_range[0]) & (filtered_trans_merged['date']<=date_range[1])]

# Multi page selector
page = st.selectbox('Page Navigation:',('COL KPI Dashboard','Exploratory Analysis','COL Time Series Analysis'))
if page == 'Exploratory Analysis':
    x_axis = st.sidebar.selectbox('x-axis',filtered_data_merged.columns)
    y_axis = st.sidebar.selectbox('y-axis',filtered_data_merged.columns)
    regression_plot(filtered_data_merged, 'Regression Plot', x_axis, y_axis)
    with st.beta_expander('Correlation Matrix',expanded=True):
        st.write(filtered_data_merged.corr(method='pearson'))
elif page == 'COL KPI Dashboard':
    ### Sales Forecast Accuracy (Actual Sales vs Forecast Sales) trend, vs previous period

    # Display stores with missing forecast
    col1, col2 = st.beta_columns(2)
    with col1:
        st.subheader('Missing Sales Forecast:')
        missing_sales_forecast = filtered_data_merged[filtered_data_merged['Forecast Sales']==0]
        st.dataframe(missing_sales_forecast.groupby('Store Name')['Forecast Sales'].count().sort_values(ascending=False))
    with col2:
        st.subheader('Missing Labour Forecast:')
        missing_labour_forecast = filtered_data_merged[filtered_data_merged['Forecast Hours']==0]
        st.dataframe(missing_labour_forecast.groupby('Store Name')['Forecast Hours'].count().sort_values(ascending=False))

    # Mean Absolute error of Sale and Labour Forecast
    st.subheader('Mean Absolute Percentage Error (MAPE):')
    MAPE_dist = pd.pivot_table(filtered_data_merged,values=['Sales MAPE','Labour MAPE','COL% MAPE'],index='Store Name', aggfunc={'Sales MAPE':[np.mean],'Labour MAPE':[np.mean],'COL% MAPE':[np.mean]})
    mape_group = filtered_data_merged.groupby(['Store Name'])[['COL% MAPE','Sales MAPE','Labour MAPE']].mean()
    
    with st.beta_expander('Display Data', expanded=True):
        st.dataframe(MAPE_dist)   
    # Sidebar parameter setting
    with st.sidebar.beta_expander('MAPE Parameters:', expanded=True):
        st.write('COL% MAPE: ', "{:.2f}".format(mape_group['COL% MAPE'].mean()))
        st.write('Sales MAPE: ', "{:.2f}".format(mape_group['Sales MAPE'].mean()))
        st.write('Labour MAPE: ', "{:.2f}".format(mape_group['Labour MAPE'].mean()))
        col_mape_target = st.number_input('COL% MAPE Target', value=30, min_value=0, max_value=35)
        sales_mape_target = st.number_input('Sales MAPE Target', value=23, min_value=0, max_value=35)
        labour_mape_target = st.number_input('Labour MAPE Target', value=15, min_value=0, max_value=35)

    with st.beta_expander('Display Chart', expanded=True):
        if len(store_code) > 10:
            mape_sales_col_plot = px.scatter(mape_group, x='Sales MAPE',y='Labour MAPE',color=mape_group.index, size='COL% MAPE',facet_col=None)
            mape_sales_col_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
            mape_sales_col_plot.add_hline(y=labour_mape_target)
            mape_sales_col_plot.add_vline(x=sales_mape_target)
            st.plotly_chart(mape_sales_col_plot)
        else:
            mape_plot = px.bar(mape_group, barmode='group')
            mape_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
            st.plotly_chart(mape_plot)
    
    with st.beta_expander('Action Recommendations', expanded=True):
        st.subheader('Immediate Action Required')
        important_urgent = mape_group[(mape_group['Sales MAPE'] > sales_mape_target) & (mape_group['Labour MAPE'] > labour_mape_target) & (mape_group['COL% MAPE'] > col_mape_target)].sort_values(by=['COL% MAPE'],ascending=False)
        st.write('Total', len(important_urgent),' stores that MAPE for COL%, Sales Forecasting and Labour planning is beyond target. They require immediate attention to improve sales forecasting and labour planning accuracy that will drive COL% consistency.')
        st.write(important_urgent)

        st.subheader('Improve Sales Forecasting')
        sales_improvement = mape_group[(mape_group['Sales MAPE'] > sales_mape_target) & (mape_group['Labour MAPE'] < labour_mape_target)].sort_values(by=['COL% MAPE'], ascending=False)
        st.write('Total', len(sales_improvement),' stores that are required to improve their sales forecasting accuracy, or improving labour arrangement if sales does not go as expected in order to improve COL% MAPE.')
        st.write(sales_improvement)

        st.subheader('Improve Labour Planning')
        labour_improvement = mape_group[(mape_group['Sales MAPE'] < sales_mape_target) & (mape_group['Labour MAPE'] > labour_mape_target)].sort_values(by=['COL% MAPE'], ascending=False)
        st.write('Total', len(labour_improvement),' stores that are required to improve their labour planning accuracy')
        st.write(labour_improvement)

        st.subheader('Deserve some recognitions')
        good_job = mape_group[(mape_group['COL% MAPE'] < col_mape_target) & (mape_group['Sales MAPE'] < sales_mape_target) & (mape_group['Labour MAPE'] < labour_mape_target)].sort_values(by=['COL% MAPE'], ascending=True)
        st.write('Total', len(good_job), ' have achieved all MAPE within target. Well Done.')
        st.write(good_job)

    
    st.subheader('MAPE Time Series:')

    mape_time_series = filtered_data_merged.groupby(['Date'])[['Sales MAPE','Labour MAPE','COL% MAPE']].mean()
    mape_time_series['COL% MAPE Moving Average'] = mape_time_series['COL% MAPE'].ewm(span=30,adjust=False).mean()
    mape_time_series['Sales MAPE Moving Average'] = mape_time_series['Sales MAPE'].ewm(span=30,adjust=False).mean()
    mape_time_series['Labour MAPE Moving Average'] = mape_time_series['Labour MAPE'].ewm(span=30,adjust=False).mean()
    
    with st.beta_expander('Display Data', expanded=True):
        st.write(mape_time_series)

    with st.beta_expander('Display Chart', expanded=True):
        mape_time_plot = px.line(mape_time_series)
        mape_time_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
        st.plotly_chart(mape_time_plot)

    # Productivity trend
    st.header('Productivity')
    st.subheader('Days below Baseline Sales')
    baseline_df = load_baseline()
    baseline_df = baseline_df[['Store code', 'Baseline Labour Hour','Baseline Management Hour','Baseline Total Hour','Baseline Sales','Baseline SPMH']]

    # Merge baseline with the dataframe
    b_df = filtered_data_merged.merge(baseline_df,how='left', left_on='Code', right_on='Store code',copy=False)
    b_df =  b_df.drop(['Store code'], axis=1)

    # Filter actual sales below baseline sales
    below_baseline_sales = b_df[b_df['Actual sales'] <= b_df['Baseline Sales']]
    sales_baseline_percent = round(len(below_baseline_sales)/len(b_df) * 100)
    
    # Filter manhour exceed baseline manhour desipite sales below baseline
    excess_manhour_vs_baseline = below_baseline_sales[below_baseline_sales['Baseline Total Hour']<below_baseline_sales['Total actual hours (excluded Holiday and paid leave days)']]
    excess_manhour_vs_baseline['Manhour Variance'] = excess_manhour_vs_baseline['Total actual hours (excluded Holiday and paid leave days)'] - excess_manhour_vs_baseline['Baseline Total Hour']
    excess_manhour_vs_baseline['Manhour Variance %'] = round(excess_manhour_vs_baseline['Manhour Variance'].div(excess_manhour_vs_baseline['Baseline Total Hour']).mul(100),2)
    excess_manhour_vs_baseline = excess_manhour_vs_baseline.set_index('Store Name').sort_values(['Manhour Variance'], ascending=False)
    
    manhour_saving = excess_manhour_vs_baseline['Manhour Variance'].sum()
    total_man_hour = filtered_data_merged['Total actual hours (excluded Holiday and paid leave days)'].sum()
    
    percent_manhour_saving = round(100* manhour_saving / total_man_hour, 2)
    
    # Description
    st.write('There are total', len(b_df),'work days in the filter, but', len(below_baseline_sales),'work days were below baseline, equivalent to',sales_baseline_percent, '%.')
    with st.beta_expander('Display Data when Sales below baseline by Store & Date', expanded=False):
        st.write(below_baseline_sales[['Date','Store Name','Actual sales','Baseline Sales','Actual SPMH','Baseline SPMH']])
    
    st.subheader('Manhour Variance vs Baseline Hours by Store')
    st.write(len(excess_manhour_vs_baseline),'out of',len(below_baseline_sales),'below baseline sales works days did not follow the manhour baseline requirements. Below is the distribution of MH Variance by store.')
    st.write(manhour_saving,'manhour can be reduced, equivalent to', percent_manhour_saving,'% of total manhour in the chosen period')
    excess_manhour_vs_baseline['dis_Date'] = excess_manhour_vs_baseline['Date'].apply(lambda x: x.strftime("%d %b, %Y"))
    excess_manhour_vs_baseline['Weekday'] = excess_manhour_vs_baseline['Date'].apply(lambda x: x.strftime("%A"))
    mh_var_plot = px.strip(excess_manhour_vs_baseline, x=excess_manhour_vs_baseline.index,y='Manhour Variance', hover_data=['dis_Date', 'Weekday'])
    st.plotly_chart(mh_var_plot)
    st.write(excess_manhour_vs_baseline[['Baseline Total Hour','Total actual hours (excluded Holiday and paid leave days)','Manhour Variance','Manhour Variance %']].groupby(excess_manhour_vs_baseline.index).agg({
        'Baseline Total Hour': np.sum,
        'Total actual hours (excluded Holiday and paid leave days)': np.sum,
        'Manhour Variance': np.sum,
        'Manhour Variance %': np.mean
    }).sort_values(by=['Manhour Variance'],ascending=False))
    
    st.subheader('Productivity (SPMH) vs Actual Sales:')
    # regression_plot(filtered_data_merged, 'SPMH vs Actual Sales Parameters:', 'Actual sales','Actual SPMH')

    spmh_sales_reg_df = filtered_data_merged[['Date','Code','Store Name','Actual sales','Actual SPMH']].sort_values(['Store Name','Date'])
    Actual_Sales = spmh_sales_reg_df['Actual sales'].fillna(0)
    Actual_SPMH = spmh_sales_reg_df['Actual SPMH'].fillna(0)
    spmh_sales_reg_df['dis_Date'] = spmh_sales_reg_df['Date'].apply(lambda x: x.strftime("%d %b, %Y"))
    spmh_sales_reg_df['Weekday'] = spmh_sales_reg_df['Date'].apply(lambda x: x.strftime("%A"))

    spmh_store_plt = px.scatter(spmh_sales_reg_df, x=Actual_Sales,y=Actual_SPMH,color='Store Name',trendline='ols', hover_data=['dis_Date','Weekday'])
    st.plotly_chart(spmh_store_plt)
    
    with st.beta_expander('Display Table', expanded=True):
        regression_table = regression_table(spmh_store_plt, 'Store Name')
        regression_table = regression_table[['Gradient','y-intercept','R2 Score']]
        st.dataframe(regression_table)
    with st.beta_expander('Display Chart', expanded=True):
        st.subheader('SPMH Scatterplot (Gradient vs y-intercept)')
        regression_plot = px.scatter(regression_table,x='Gradient',y='y-intercept', size='R2 Score',color=regression_table.index)
        regression_plot.add_hline(y=regression_table['y-intercept'].mean())
        regression_plot.add_vline(x=regression_table['Gradient'].mean())
        st.plotly_chart(regression_plot, use_container_width=True)

    st.subheader('SPMH Time Series and Trend')

    spmh_df, spmh_ts_plot = spmh_time_series(filtered_data_merged, 'D')
    spmh_ts_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
    st.plotly_chart(spmh_ts_plot, use_container_width=True )
    
    st.subheader('SPMH Weekdays Boxplot')
    spmh_df['weekdays'] = spmh_df.index.day_name()
    weekday_spmh_plot = px.box(spmh_df, x='weekdays', y='Actual SPMH')
    weekday_spmh_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
    st.plotly_chart(weekday_spmh_plot, use_container_width=True)

    with st.beta_expander('Action Recommendations', expanded=False):
        st.write('These stores are low in productivity and has a steep difference in SPMH in high and low sales situation:')
        spmh_worst_stores = regression_table[(regression_table['Gradient']>regression_table['Gradient'].mean()) & (regression_table['y-intercept']<regression_table['y-intercept'].mean())]
        st.write(spmh_worst_stores)
    #st.write(store_code_func)
    #st.write(filtered_data_merged.columns)
    #meal_df = filtered_data_merged[['Date','Store Name','Meal Allowance']]
    #meal_df['Quantity'] = meal_df['Meal Allowance'] / 25000
    #st.write(meal_df)
    #meal_df = meal_df.set_index('Date')
    #meal_con = meal_df.resample('M').sum()
    #st.write(meal_con)

    #ot_df = filtered_data_merged[['Date','Store Name','Over time in holiday']].set_index('Date')
    #ot_con = ot_df.resample('M').sum()
    #st.write(ot_con)
    #st.write('OT in 1 Jan 2020: ',ot_df.loc['Jan 01, 2020','Over time in holiday'].sum())
elif page == 'COL Time Series Analysis':
    with st.beta_expander('Time Series Analysis', expanded=True):
        spmh_df, spmh_ts_plot = spmh_time_series(filtered_data_merged, 'D')
        st.write(spmh_df)
        st.plotly_chart(spmh_ts_plot, use_container_width=True )

    with st.beta_expander('System aggregated Information', expanded=True):
        # Select box to display data rsampled by Hour, Day, Week, and Month
        resample_data =[['Day','D'],['Week','W'],['Month','M']]
        resample_df = pd.DataFrame(resample_data, columns=['name','id'])
        resample_values = resample_df['name'].tolist()
        resample_id = resample_df['id'].tolist()
        dic = dict(zip(resample_id,resample_values))
        data_resample_option = st.selectbox('Data resample by:',resample_id,format_func=lambda x:dic[x])
        st.write(filtered_data_merged.columns)
        # Apply sampled option to dataframe
        system_total_df = filtered_data_merged.resample(data_resample_option, on='Date').agg({
            'Forecast Sales': np.sum,
            'Forecast Hours': np.sum,
            'Forecast COL $': np.sum,
            'Forecast COL %': np.mean,
            'Actual sales': np.sum,
            'Total actual hours (included Holiday and paid leave days)':np.sum,
            'Actual hours of MNGT': np.sum,
            'Actual hours of TMs Full time': np.sum,
            'Actual hours of TMs Part time': np.sum,
            'Total actual hours (excluded Holiday and paid leave days)': np.sum,
            'Hours of holiday/paid leave days': np.sum,
            'Total COL $ (included Holiday and paid leave days)':np.sum,
            'COL $  of TM Full time': np.sum,
            'COL $  of TM Part time': np.sum,
            'COL $ Management': np.sum,
            'Total COL (excluded Holiday and paid leave days)': np.sum,
            'COL of holidays/paid leave days': np.sum,
            'Actual COL %': np.mean,
            'COL Val %': np.mean,
            'Working hour in work shift': np.sum,
            'Over time in normal day':np.sum,
            'Over time in weekend':np.sum,
            'Over time in holiday': np.sum,
            'working hour in normal day (Night)': np.sum,
            'Over time in normal day (Night)':np.sum,
            'Over time in weekend (Night)':np.sum,
            'Over time in holiday (Night)':np.sum,
            '13th salary':np.sum,
            'BSC bonus':np.sum,
            'PA bonus':np.sum,
            'Meal Allowance':np.sum,
            'Insurance contribution Amount per day':np.sum,
            'Actual TPMH': np.mean,
            'Actual SPMH': np.mean,
            'Total Transaction':np.sum,
            'TA': np.mean
        })

        system_total_df['weekdays'] = system_total_df.index.day_name()
        system_total_df['Actual SPMH'] = system_total_df['Actual sales'].div(system_total_df['Total actual hours (included Holiday and paid leave days)'])
        system_total_df['Actual COL %'] = 100*system_total_df['Total COL $ (included Holiday and paid leave days)'].div(system_total_df['Actual sales'])
        st.dataframe(system_total_df)
        spmh_sales_plot = px.scatter(system_total_df,x='Actual sales',y='Actual SPMH', trendline='ols')
        st.plotly_chart(spmh_sales_plot)
        system_total_df['Labour rate'] = system_total_df['Actual COL %'].div(100).mul(system_total_df['Actual SPMH'])
        
        col1, col2 = st.beta_columns(2)
        with col1:
            st.subheader('Time Series Plot:')
            time_series_option = st.selectbox('Select data', system_total_df.columns, key='time_series_option')
            time_series_plot = px.line(system_total_df,x=system_total_df.index,y=time_series_option)
            time_series_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
            st.plotly_chart(time_series_plot, use_container_width=True)
        with col2:
            st.subheader('Weekdays Distribution:')
            weekdays_options = st.selectbox('Select data', system_total_df.columns, key='weekdays_options')
            weekdays_rate_plot = px.box(system_total_df,x='weekdays',y=weekdays_options)
            weekdays_rate_plot.update_layout(margin={'l':0,'r':0,'b':0,'t':0})
            st.plotly_chart(weekdays_rate_plot, use_container_width=True)
        
        ta_productivity = system_total_df[['Actual TPMH','Actual SPMH','TA']]
        st.write(ta_productivity)

