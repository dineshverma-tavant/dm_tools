import streamlit as st
import pandas as pd
from simple_salesforce import Salesforce
from io import BytesIO
import plotly.express as px
import pyodbc
from sqlalchemy import create_engine

# Set page configuration
st.set_page_config(page_title="Data Upload and Preview", layout="wide")

# Initialize session state
if "df" not in st.session_state:
    st.session_state.df = None
if "data_source" not in st.session_state:
    st.session_state.data_source = "File Upload"
if "sf_connection" not in st.session_state:
    st.session_state.sf_connection = None
if "sf_credentials" not in st.session_state:
    st.session_state.sf_credentials = {}

# Title and description
st.title("Data Upload and Preview App")
st.markdown("Select a data source and view options from the sidebar to explore data.")

# Sidebar for navigation and data source
st.sidebar.header("Data Source Selection")
data_source = st.sidebar.selectbox(
    "Choose Data Source",
    ["File Upload", "Salesforce"],
    index=["File Upload", "Salesforce"].index(st.session_state.data_source)
)

st.sidebar.header("Navigation")
view_option = st.sidebar.radio("Select View", ["Data Preview", "Data Statistics", "Data Cleaning", "Save Data", "Graph"], index=0)

# Handle data source
if data_source != st.session_state.data_source:
    st.session_state.df = None
    st.session_state.data_source = data_source
    st.session_state.sf_connection = None  # Reset Salesforce connection

if data_source == "File Upload":
    uploaded_file = st.file_uploader("Choose a file", type=["csv", "xls", "xlsx"])
    if uploaded_file is not None:
        try:
            file_extension = uploaded_file.name.split(".")[-1].lower()
            if file_extension in ["xls", "xlsx"]:
                st.session_state.df = pd.read_excel(uploaded_file)
            elif file_extension == "csv":
                st.session_state.df = pd.read_csv(uploaded_file)
            else:
                st.error("Unsupported file format. Please upload a CSV or Excel file.")
                st.stop()
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.session_state.df = None
    elif uploaded_file is None and st.session_state.df is None:
        st.info("Please upload a CSV or Excel file to begin.")

elif data_source == "Salesforce":
    st.sidebar.header("Salesforce Credentials")
    sf_username = st.sidebar.text_input("Username", type="default", key="sf_username")
    sf_password = st.sidebar.text_input("Password", type="password", key="sf_password")
    sf_security_token = st.sidebar.text_input(
        "Security Token", type="password", help="Leave blank if not required for sandbox", key="sf_token"
    )
    sf_domain = st.sidebar.selectbox(
        "Environment", ["login", "test"], help="Select 'login' for production or 'test' for sandbox", key="sf_domain"
    )
    sf_query = st.sidebar.text_area("Enter SOQL Query", value="SELECT Id, Name FROM Account LIMIT 100", key="sf_query")

    if st.sidebar.button("Connect and Query", key="connect_query"):
        if sf_username and sf_password and sf_query:
            with st.spinner("Querying Salesforce..."):
                try:
                    sf = Salesforce(
                        username=sf_username,
                        password=sf_password,
                        security_token=sf_security_token or None,
                        domain=sf_domain
                    )
                    st.session_state.sf_connection = sf
                    st.session_state.sf_credentials = {
                        "username": sf_username,
                        "password": sf_password,
                        "security_token": sf_security_token,
                        "domain": sf_domain
                    }
                    query_result = sf.query_all(sf_query)
                    st.session_state.df = pd.DataFrame(query_result["records"])
                    if "attributes" in st.session_state.df.columns:
                        st.session_state.df = st.session_state.df.drop(columns=["attributes"])
                    st.success("Successfully queried Salesforce data!")
                except Exception as e:
                    st.error(f"Error connecting to Salesforce or executing query: {str(e)}")
                    st.session_state.df = None
                    st.session_state.sf_connection = None
        else:
            st.error("Please provide Salesforce username, password, and a valid SOQL query.")
    elif st.session_state.df is None:
        st.info("Enter Salesforce credentials and a SOQL query, then click 'Connect and Query'.")

# Process and display data if available
if st.session_state.df is not None:
    try:
        # Display basic information
        st.subheader("Dataset Info")
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"Rows: {st.session_state.df.shape[0]}")
        with col2:
            st.write(f"Columns: {st.session_state.df.shape[1]}")

        # Display content based on selected view
        if view_option == "Data Preview":
            st.subheader("Data Preview")
            num_rows = st.slider(
                "Select number of rows to preview",
                min_value=1,
                max_value=min(100, st.session_state.df.shape[0]),
                value=min(10, st.session_state.df.shape[0]),
                step=1,
                key="preview_slider"
            )
            st.dataframe(st.session_state.df.head(num_rows), use_container_width=True)

        elif view_option == "Data Statistics":
            st.subheader("Data Statistics")
            st.dataframe(st.session_state.df.describe(), use_container_width=True)

        elif view_option == "Data Cleaning":
            st.subheader("Data Cleaning")
            cleaning_option = st.selectbox(
                "Select Cleaning Operation",
                ["Sort Data", "Remove Null", "Group By"],
                key="cleaning_option"
            )

            if cleaning_option == "Sort Data":
                st.write("Sort Data")
                sort_column = st.selectbox("Select column to sort by", st.session_state.df.columns, key="sort_column")
                sort_order = st.radio("Sort order", ["Ascending", "Descending"], index=0, key="sort_order")
                sorted_df = st.session_state.df.sort_values(
                    by=sort_column,
                    ascending=(sort_order == "Ascending")
                )
                st.dataframe(sorted_df, use_container_width=True)

            elif cleaning_option == "Remove Null":
                st.write("Remove Null Values")
                null_column = st.multiselect(
                    "Select columns to check for null values",
                    st.session_state.df.columns,
                    key="null_column"
                )
                if st.button("Remove Nulls", key="remove_nulls"):
                    if null_column:
                        cleaned_df = st.session_state.df.dropna(subset=null_column)
                        st.session_state.df = cleaned_df
                        st.success(f"Removed rows with null values in {', '.join(null_column)}. New row count: {cleaned_df.shape[0]}")
                    else:
                        st.warning("Please select at least one column.")
                st.dataframe(st.session_state.df, use_container_width=True)

            elif cleaning_option == "Group By":
                st.write("Group By")
                group_column = st.selectbox("Select column to group by", st.session_state.df.columns, key="group_column")
                agg_columns = st.multiselect(
                    "Select columns to aggregate",
                    [col for col in st.session_state.df.columns if col != group_column],
                    key="agg_columns"
                )
                agg_function = st.selectbox(
                    "Select aggregation function",
                    ["count", "sum", "mean", "min", "max"],
                    key="agg_function"
                )
                if st.button("Group Data", key="group_data"):
                    if agg_columns:
                        try:
                            grouped_df = st.session_state.df.groupby(group_column)[agg_columns].agg(agg_function).reset_index()
                            st.session_state.df = grouped_df
                            st.success(f"Grouped by {group_column} with {agg_function} aggregation.")
                        except Exception as e:
                            st.error(f"Error grouping data: {str(e)}")
                    else:
                        st.warning("Please select at least one column to aggregate.")
                st.dataframe(st.session_state.df, use_container_width=True)

        elif view_option == "Save Data":
            st.subheader("Save Data")
            save_format = st.selectbox("Select save format", ["CSV", "Excel", "SQL Server", "Salesforce"], key="save_format")
            total_rows = len(st.session_state.df)
            st.info(f"Total rows in data: {total_rows}")
            num_rows_input = st.number_input(
                "Enter number of rows to save (leave blank for all rows)",
                min_value=0,
                max_value=total_rows,
                value=0,
                step=1,
                key="save_num_rows"
            )

            # Determine DataFrame to save
            df_to_save = st.session_state.df if num_rows_input == 0 else st.session_state.df.head(num_rows_input)

            if save_format in ["CSV", "Excel"]:
                filename = st.text_input("Enter filename (without extension)", value="output", key="save_filename")
                if save_format == "CSV":
                    buffer = BytesIO()
                    df_to_save.to_csv(buffer, index=False)
                    file_extension = "csv"
                    mime_type = "text/csv"
                else:  # Excel
                    buffer = BytesIO()
                    df_to_save.to_excel(buffer, index=False, engine="openpyxl")
                    file_extension = "xlsx"
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                st.download_button(
                    label=f"Download as {save_format}",
                    data=buffer.getvalue(),
                    file_name=f"{filename}.{file_extension}",
                    mime=mime_type
                )

            elif save_format == "SQL Server":
                table_name = st.text_input("Enter table name", value="data_table", key="sql_table")
                server = st.text_input("SQL Server name (e.g., localhost)", key="sql_server")
                database = st.text_input("Database name", key="sql_database")
                sql_username = st.text_input("Username", key="sql_username")
                sql_password = st.text_input("Password", type="password", key="sql_password")
                driver = st.text_input("Driver", value="ODBC Driver 17 for SQL Server", key="sql_driver")
                if st.button("Save to SQL Server", key="save_sql"):
                    if server and database and table_name:
                        try:
                            # Create connection string
                            connection_string = (
                                f"mssql+pyodbc://{sql_username}:{sql_password}@{server}/{database}"
                                f"?driver={driver.replace(' ', '+')}"
                            )
                            engine = create_engine(connection_string)
                            df_to_save.to_sql(table_name, engine, if_exists="replace", index=False)
                            st.success(f"Data saved to {table_name} in {database} on {server}")
                        except Exception as e:
                            st.error(f"Error saving to SQL Server: {str(e)}")
                    else:
                        st.error("Please provide server, database, and table name.")

            elif save_format == "Salesforce":
                st.write("Load Data to Salesforce")
                use_existing_creds = st.checkbox("Use credentials from Salesforce query section", value=True, key="use_sf_creds")
                if not use_existing_creds:
                    sf_username_save = st.text_input("Username", type="default", key="sf_username_save")
                    sf_password_save = st.text_input("Password", type="password", key="sf_password_save")
                    sf_security_token_save = st.text_input(
                        "Security Token", type="password", help="Leave blank if not required for sandbox", key="sf_token_save"
                    )
                    sf_domain_save = st.selectbox(
                        "Environment", ["login", "test"], help="Select 'login' for production or 'test' for sandbox", key="sf_domain_save"
                    )
                else:
                    sf_username_save = st.session_state.sf_credentials.get("username", "")
                    sf_password_save = st.session_state.sf_credentials.get("password", "")
                    sf_security_token_save = st.session_state.sf_credentials.get("security_token", "")
                    sf_domain_save = st.session_state.sf_credentials.get("domain", "login")

                sf_object = st.text_input("Enter Salesforce object name (e.g., Account, Contact)", value="Account", key="sf_object")
                operation = st.selectbox("Select operation", ["Insert", "Update"], key="sf_operation")
                if operation == "Update":
                    id_column = st.selectbox("Select ID column", st.session_state.df.columns, key="sf_id_column")
                else:
                    id_column = None

                # Allow column mapping, skip blank mappings
                st.write("Map DataFrame columns to Salesforce fields (leave blank to exclude)")
                field_mappings = {}
                for col in st.session_state.df.columns:
                    default_value = "BillingState" if col == "State" else col
                    sf_field = st.text_input(f"Map '{col}' to Salesforce field", value=default_value, key=f"map_{col}")
                    if sf_field.strip():  # Only include non-empty mappings
                        field_mappings[col] = sf_field.strip()

                if st.button("Load to Salesforce", key="load_sf"):
                    if not field_mappings:
                        st.error("Please provide at least one field mapping.")
                        st.stop()
                    if not use_existing_creds and not (sf_username_save and sf_password_save):
                        st.error("Please provide Salesforce username and password.")
                        st.stop()
                    if use_existing_creds and not (st.session_state.sf_credentials.get("username") and st.session_state.sf_credentials.get("password")):
                        st.error("No valid credentials found from Salesforce query section. Please provide credentials above.")
                        st.stop()

                    # Establish Salesforce connection if not already connected
                    if not st.session_state.sf_connection:
                        try:
                            sf = Salesforce(
                                username=sf_username_save,
                                password=sf_password_save,
                                security_token=sf_security_token_save or None,
                                domain=sf_domain_save
                            )
                            st.session_state.sf_connection = sf
                            if use_existing_creds:
                                st.session_state.sf_credentials = {
                                    "username": sf_username_save,
                                    "password": sf_password_save,
                                    "security_token": sf_security_token_save,
                                    "domain": sf_domain_save
                                }
                        except Exception as e:
                            st.error(f"Failed to connect to Salesforce: {str(e)}")
                            st.stop()

                    try:
                        sf = st.session_state.sf_connection
                        success_count = 0
                        error_count = 0
                        for _, row in df_to_save.iterrows():
                            record = {field_mappings[col]: row[col] for col in field_mappings}
                            try:
                                if operation == "Update" and id_column:
                                    record_id = row[id_column]
                                    sf_obj = getattr(sf, sf_object)
                                    sf_obj.update(record_id, record)
                                else:
                                    sf_obj = getattr(sf, sf_object)
                                    sf_obj.create(record)
                                success_count += 1
                            except Exception as e:
                                error_count += 1
                                st.warning(f"Error processing row {row.get(id_column, 'unknown')}: {str(e)}")
                        # st.success(f"Loaded {success_count} records to {sf_object}. Errors: {error_count}")
                    except Exception as e:
                        st.error(f"Error loading data to Salesforce: {str(e)}")

        elif view_option == "Graph":
            st.subheader("Graph")
            graph_type = st.selectbox("Select graph type", ["Bar", "Line", "Scatter"], key="graph_type")
            x_column = st.selectbox("Select X-axis column", st.session_state.df.columns, key="x_column")
            y_column = st.selectbox("Select Y-axis column", st.session_state.df.columns, key="y_column")

            try:
                if graph_type == "Bar":
                    fig = px.bar(st.session_state.df, x=x_column, y=y_column, title=f"{y_column} vs {x_column}")
                elif graph_type == "Line":
                    fig = px.line(st.session_state.df, x=x_column, y=y_column, title=f"{y_column} vs {x_column}")
                elif graph_type == "Scatter":
                    fig = px.scatter(st.session_state.df, x=x_column, y=y_column, title=f"{y_column} vs {x_column}")
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Error generating graph: {str(e)}")

    except Exception as e:
        st.error(f"Error displaying data: {str(e)}")

# Add styling
st.markdown("""
<style>
    .stFileUploader {
        margin-top: 20px;
    }
    .stDataFrame {
        margin-bottom: 20px;
    }
    .stSidebar .stRadio > div {
        padding: 10px;
    }
    .stSidebar .stSelectbox {
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)
