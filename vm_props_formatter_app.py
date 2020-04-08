import argparse
import base64
import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table as dt
import io
import pandas as pd
import webbrowser
from dash.dependencies import Input, State, Output
from flask import send_file
from vm_props_formatter.utils.logger import format_logs

# Set up the app
app = dash.Dash(__name__)
app.title = 'VM Props Formatter App'
server = app.server
app_url = 'http://127.0.0.1:8050/'

# Define global variables
current_start_order_check_click_count = 0
rejected_data = {}
report_filename = 'report.xlsx'

# Define upload default text
upload_default_text = html.Div([
    'Drag and drop or ',
    html.A('select file')
])
# Define UI styles
input_area_style = {
    'width': '23%',
    'display': 'inline-block',
    'vertical-align': 'top'
}
separator_area_style = {
    'width': '2%',
    'display': 'inline-block',
    'vertical-align': 'top'
}
output_area_style = {
    'width': '75%',
    'display': 'inline-block',
    'vertical-align': 'top'
}
upload_box_style = {
    'width': '100%',
    'height': '60px',
    'lineHeight': '60px',
    'borderWidth': '1px',
    'borderStyle': 'dashed',
    'borderRadius': '10px',
    'textAlign': 'center',
    'margin': '0px'
}
start_button_style = {
    'width': '50%',
    'height': '60px',
    'textAlign': 'center'
}
center_placement_style = {
    'display': 'flex',
    'align-items': 'center',
    'justify-content': 'center'
}

# UI defs
def generate_datatable_title(title):
    """
    Generate datatable title

    Parameters
    ----------
    title : str
        Title

    Returns
    -------
    output : dash_html_components.Div
        Datatable title
    """
    return html.Div(
        children=[html.P(title)],
        style=center_placement_style
    )


def generate_datatable(data, id):
    """
    Generate datatable

    Parameters
    ----------
    data : pandas.DataFrame
        Data
    id : str
        Datatable ID

    Returns
    -------
    output : dash_html_components.Div
        Datatable
    """
    return html.Div(
        children=[
            dt.DataTable(
                id=id,
                data=data.to_dict('rows'),
                columns=[{'name': i, 'id': i} for i in data.columns],
                sort_action="native",
                sort_mode='multi',
                style_header={
                    'backgroundColor': 'rgb(220, 220, 220)',
                    'fontWeight': 'bold'
                },
                style_table={
                    'maxHeight': '500px',
                    'overflowX': 'scroll',
                    'overflowY': 'scroll',
                    'border': 'thin lightgrey solid'
                },
                style_data_conditional=[{
                    'if': {'row_index': 'odd'},
                    'backgroundColor': 'rgb(250, 250, 250)'
                }]
            )
        ],
        style=center_placement_style
    )


def generate_empty_datatable(id, max_cell_width=500):
    """
    Generate an empty datatable

    Parameters
    ----------
    id : str
        Datatable ID
    max_cell_width : int
        Max cell width

    Returns
    -------
    output : dash_table.DataTable
        Datatable
    """
    return dt.DataTable(
        id=id,
        sort_action="native",
        sort_mode='multi',
        row_selectable='multi',
        style_header={
            'backgroundColor': 'rgb(220, 220, 220)',
            'fontWeight': 'bold'
        },
        style_table={
            'maxWidth': '1000px',
            'maxHeight': '400px',
            'overflowX': 'scroll',
            'overflowY': 'scroll',
            'border': 'thin lightgrey solid'
        },
        style_cell={
            'height': 'auto',
            'minWidth': '0px',
            'maxWidth': '%dpx' % max_cell_width,
            'whiteSpace': 'normal'
        },
        style_data={
            'height': 'auto',
            'whiteSpace': 'normal'
        },
        style_data_conditional=[{
            'if': {'row_index': 'odd'},
            'backgroundColor': 'rgb(250, 250, 250)'
        }]
    )


def generate_file_error_message(filenames):
    """
    Generate file error message

    Parameters
    ----------
    filenames : str
        Filenames

    Returns
    -------
    output : dash_html_components.Div
        File error message
    """
    return html.Div(
        children=[html.P(
            'Error: Incorrect file format, table columns, or table contents, please check the following file(s): ' + filenames)],
        style=center_placement_style
    )


def generate_no_error_message():
    """
    Generate no error message

    Parameters
    ----------
    None

    Returns
    -------
    output : dash_html_components.Div
        No error message
    """
    return html.Div(
        children=[html.P('Status: Checking has completed successfully, no errors have been found')],
        style=center_placement_style
    )


# Define app layout
app.layout = html.Div(
    id='app-body',
    children=[
        html.H1('Order Checking'),
        html.Div(
            id='tabs-area',
            children=[
                dcc.Tabs(
                    id='tabs',
                    children=[
                        dcc.Tab(
                            id='shipping-order-tab',
                            label='Shipping Order',
                            children=[
                                html.Div(
                                    id='shipping-order-input-area',
                                    children=[
                                        html.P('VM Props Order Summary File'),
                                        dcc.Upload(
                                            id='upload-vm-props-order-summary',
                                            style=upload_box_style
                                        ),
                                        html.P('Country-Warehouse Naming File'),
                                        dcc.Upload(
                                            id='upload-country-whs-names',
                                            style=upload_box_style
                                        ),
                                        html.P('VM Props-Batch Naming File'),
                                        dcc.Upload(
                                            id='upload-props-batch-names',
                                            style=upload_box_style
                                        ),
                                        html.P(''),
                                        html.Div(
                                            children=[
                                                html.Button(
                                                    id='start-order-check-button',
                                                    n_clicks=0,
                                                    children='Start checking',
                                                    style=start_button_style
                                                )
                                            ],
                                            style=center_placement_style
                                        ),
                                        html.P(''),
                                        html.Div(
                                            id='download-report-area',
                                            style=center_placement_style
                                        )
                                    ],
                                    style=input_area_style
                                ),
                                html.Div(
                                    style=separator_area_style
                                ),
                                html.Div(
                                    id='shipping-order-output-area',
                                    children=[
                                        dcc.Loading(
                                            id='loading',
                                            children=[
                                                html.Div(
                                                    children=[
                                                        html.Div(
                                                            id='so-format-datatable-area',
                                                            children=[
                                                                html.H4('Matched Stock Transfers'),
                                                                html.Div([generate_empty_datatable('so-format-datatable')], style=center_placement_style)
                                                            ]
                                                        )
                                                    ]
                                                )
                                            ],
                                            type='circle',
                                            style={'vertical-align': 'middle'}
                                        ),
                                    ],
                                    style=output_area_style
                                )
                            ]
                        )
                    ]
                )
            ]
        )
    ]
)

@app.callback(
    Output('upload-vm-props-order-summary', 'children'),
    [Input('upload-vm-props-order-summary', 'contents')],
    [State('upload-vm-props-order-summary', 'filename')]
)
def display_vm_props_order_summary_filename(vm_props_order_summary_content, vm_props_order_summary_filename):
    """
    Display inventory order summary filename

    Parameters
    ----------
    vm_props_order_summary_content : str
        File content
    vm_props_order_summary_filename : str
        Filename

    Returns
    -------
    output : dash_html_components.Div
        Filename
    """
    if None not in (vm_props_order_summary_content, vm_props_order_summary_filename):
        return html.Div([vm_props_order_summary_filename])
    else:
        return upload_default_text

@app.callback(
    Output('upload-country-whs-names', 'children'),
    [Input('upload-country-whs-names', 'contents')],
    [State('upload-country-whs-names', 'filename')]
)
def display_repeat_order_filename(country_whs_content, ountry_whs_filename):
    """
    Display repeat order filename

    Parameters
    ----------
    repeat_order_content : str
        File content
    repeat_order_filename : str
        Filename

    Returns
    -------
    output : dash_html_components.Div
        Filename
    """
    if None not in (country_whs_content, ountry_whs_filename):
        return html.Div([ountry_whs_filename])
    else:
        return upload_default_text


@app.callback(
    Output('upload-props-batch-names', 'children'),
    [Input('upload-props-batch-names', 'contents')],
    [State('upload-props-batch-names', 'filename')]
)
def display_props_batch_filename(props_batch_content, props_batch_filename):
    """
    Display business rules filename

    Parameters
    ----------
    business_rules_content : str
        File content
    business_rules_filename : str
        Filename

    Returns
    -------
    output : dash_html_components.Div
        Filename
    """
    if None not in (props_batch_content, props_batch_filename):
        return html.Div([props_batch_filename])
    else:
        return upload_default_text

@app.callback(
    [
        Output('so-format-datatable', 'data'),
        Output('so-format-datatable', 'columns'),
        Output('so-format-datatable', 'selected_rows'),
        Output('download-report-area', 'children')
    ],
    [
        Input('upload-vm-props-order-summary', 'contents'),
        Input('upload-country-whs-names', 'contents'),
        Input('upload-props-batch-names', 'contents'),
        Input('start-order-check-button', 'n_clicks')
    ],
    [
        State('upload-vm-props-order-summary', 'filename'),
        State('upload-country-whs-names', 'filename'),
        State('upload-props-batch-names', 'filename'),
    ]
)
def check_order(vm_props_order_summary_content, country_whs_content, props_batch_content, start_clicks, 
                vm_props_order_summary_filename, country_whs_content_filename, props_batch_content_filename):
    """
    Perform checking of the files

    Parameters
    ----------
    vm_props_order_summary_content : str
        VM Props order summary file
    purchase_order_summary_content : str
        Purchase order summary file
    po_details_summary_content : str
        Po details summary file
    repeat_order_content : str
        Repeat order file
    business_rules_content : str
        Business rules file
    start_clicks : int
        Total clicks of start button
    vm_props_order_summary_filename : str
        VM Props order summary filename

    Returns
    -------
    duplicated_output : list
        List of visualization outputs for duplicated data row check
    download_report_output: list
        Report download link
    """
    global current_start_order_check_click_count
    global so_format_data
    so_format_data = pd.DataFrame({'A':[1,3,4], 'B':[2,5,6]})
    download_report_output = []
    are_outputs_available = False
    if start_clicks is not None and start_clicks > 0 and start_clicks > current_start_order_check_click_count:
        # Run checking

        # Format outputs
        if so_format_data is not None and not so_format_data.empty:
            so_format_datatable_data = so_format_data.to_dict('rows')
            so_format_datatable_columns = [{'name': i, 'id': i} for i in so_format_data.columns]
            are_outputs_available = True

        # Generate report download link
        if are_outputs_available:
            download_report_output = [
                html.A(
                    'Download report',
                    id='download-report-link',
                    href='/downloads/'
                )
            ]
        # Update current clicks
        current_start_order_check_click_count = start_clicks

    return so_format_datatable_data, so_format_datatable_columns, [], download_report_output


# Download the report
@app.server.route('/downloads/')
def download_report():
    """
    Download the report

    Parameters
    ----------
    None

    Returns
    -------
    send_file : flask.send_file
        File sending function
    """
    global so_format_data
    buffer = io.BytesIO()
    excel_writer = pd.ExcelWriter(buffer, engine='openpyxl')
    for key in so_format_data.keys():
        so_format_data[key].to_excel(excel_writer, sheet_name=key, index=False)
    excel_writer.save()
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=report_filename,
        as_attachment=True,
        cache_timeout=0
    )


# Run the Dash app server
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='VM Props Formatter App')
    parser.add_argument('--debug', help='Run the app in debug mode', action='store_true')
    arguments = parser.parse_args()
    format_logs('Store Consolidation', True)
    webbrowser.open(app_url)
    # Run the Dash server
    app.run_server(debug=arguments.debug)