import argparse
import base64
import dash
import datetime
import dash_core_components as dcc
import dash_bootstrap_components as dbc
import dash_html_components as html
import dash_table as dt
import io
import pandas as pd
import webbrowser
from dash.dependencies import Input, State, Output
from flask import send_file
from vm_props_formatter.vm_props_manager import VMPropsManager
from vm_props_formatter.utils.logger import format_logs
from vm_props_formatter.utils.json_parser import read_json, write_json

# Set up the app
app = dash.Dash(__name__)
app.title = 'VM Props Formatter App'
server = app.server
app_url = 'http://127.0.0.1:8050/'

# Define global variables
current_start_analysis_clicks = 0
current_load_settings_clicks = 0
current_save_settings_clicks = 0
current_load_default_settings_clicks = 0
so_format_data = pd.DataFrame()
checked_data = pd.DataFrame()
entity = None
settings = {}
settings_path = 'settings/default_settings.json'
outputs_path = 'outputs/'
image_filename = 'settings/ck_logo.png'
encoded_image = base64.b64encode(open(image_filename, 'rb').read())
hover_text = {
    'settings-shape-main-header-row-text':
        ['e.g. Row 8 where the main data row starts'],
    'settings-shape-props-header-start-col-text':
        ['e.g. Col 6 or F where the props column starts at BIG BAG STAND'],
    'settings-shape-props-header-tally-first-text':
        ['e.g. How many columns to tally from the first props col stand to find summary below'],
    'settings-shape-no-summary-table-rows-text':
        ['e.g. How many rows does summary table have'],
    'settings-shape-summary-table-sum-row-text':
        ['e.g. Row 2 in summary table contains the total values to check'],
    'settings-shape-props-header-end-col-text': 
        ['e.g. Last column (REMARKS) to drop that do not require to tallying'],
    'settings-names-sheet-name-text':
        ['If left as blank, app auto-reads sheet name that has any of the entity names indicated below'],
    'settings-names-store-col-text':
        ['Column name that represents Store'],
    'settings-names-country-col-text':
        ['Column name that represents Country'],
    'settings-names-main-cols-text':
        ['Main cols that represents index of SO Table.',html.Br(),
         'Tip: Put only Country if the Store Locations given are wrong.'],
    'settings-names-entity-list-text':
        ['This list indicates how to find sheet name based on if entity is present in file name.']
}
entity_list = ['CKS','CKI','CKC']

# Define upload default text
upload_default_text = html.Div([
    'Drag and drop or ',
    html.A('select file')
])
# Define UI styles
input_area_style = {
    'width': '23%',
    'display': 'inline-block',
    'vertical-align': 'top',
    'font-family': 'Helvetica'
}
separator_area_style = {
    'width': '2%',
    'display': 'inline-block',
    'vertical-align': 'top',
    'font-family': 'Helvetica'
}
output_area_style = {
    'width': '75%',
    'display': 'inline-block',
    'vertical-align': 'top',
    'font-family': 'Helvetica'
}
left_output_area_style = {
    'width': '47%',
    'display': 'inline-block',
    'vertical-align': 'top'
}
right_output_area_style = {
    'width': '47%',
    'display': 'inline-block',
    'vertical-align': 'top',
    'font-family': 'Helvetica'
}
upload_box_style = {
    'width': '100%',
    'height': '60px',
    'lineHeight': '60px',
    'borderWidth': '1px',
    'borderStyle': 'dashed',
    'borderRadius': '10px',
    'textAlign': 'center',
    'margin': '0px',
    'font-family': 'Helvetica'
}
button_style = {
    'width': '50%',
    'height': '60px',
    'textAlign': 'center',
    'font-family': 'Helvetica'
}
center_placement_style = {
    'display': 'flex',
    'align-items': 'center',
    'justify-content': 'center',
    'font-family': 'Helvetica'
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


def generate_hover_text(target_name):
    return dbc.Tooltip(
        hover_text[target_name],
        target=target_name,
        placement='top-start',
        style={'width': '220pt',
               'line-height': '150%',
               'background-color': 'white',
               'font-family': 'Helvetica',
               'border': '2px solid #89CFF0',
               'border-radius': '5px',
               'padding-top': '5px',
               'padding-bottom': '5px',
               'padding-right': '5px',
               'padding-left': '10px'}
    )


def generate_slider(id):
    m = 5 # multiple 
    max_v = 20 # max bar
    count = int(round(max_v/m,0)) # number of marks
    markers = {}
    for i in range(m,(count+1)*m,m):
        markers[i]= str(i) # add items to dict
    return dcc.Slider(
        id=id,
        min=0,
        max=max_v,
        step=1,
        marks=markers
    )

# Define app layout
app.layout = html.Div(
    id='app-body',
    children=[
        html.Div([
            html.H1('VM Props Formatter')
        ], style={'text-align': 'left',
                  'margin-right': 20,
                  'display': 'inline-block',
                  'font-family': 'Helvetica'}),
        html.Div([
            html.Img(
                src='data:image/png;base64,{}'.format(encoded_image.decode()),
                style={'height': '60%',
                       'width': '60%'})
        ], style={'float': 'right',
                  'display': 'inline-block'}),
        html.Div(
            id='tabs-area',
            children=[
                dcc.Tabs(
                    id='tabs',
                    children=[
                        dcc.Tab(
                            id='analysis-tab',
                            label='Analysis',
                            children=[
                                html.P(''),
                                html.Div(
                                    id='analysis-input-area',
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
                                                    style=button_style
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
                                    id='analysis-output-area',
                                    children=[
                                        dcc.Loading(
                                            id='loading',
                                            children=[
                                                html.Div(
                                                    children=[
                                                        html.Div(
                                                            id='so-format-datatable-area',
                                                            children=[
                                                                html.H4('SO Table'),
                                                                html.Div(
                                                                    [generate_empty_datatable('so-format-datatable')],
                                                                    style=center_placement_style
                                                                )
                                                            ]
                                                        ),
                                                        html.Div(
                                                            id='checked-datatable-area',
                                                            children=[
                                                                html.H4('Cross Checked Table'),
                                                                html.Div(
                                                                    [generate_empty_datatable('checked-datatable')],
                                                                    style=center_placement_style
                                                                )
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
                        ),
                        dcc.Tab(
                            id='settings-tab',
                            label='Settings',
                            children=[
                                html.Div(
                                    id='settings-input-area',
                                    children=[
                                        html.P(''),
                                        html.Div(
                                            children=[
                                                html.Button(
                                                    id='settings-load-button',
                                                    n_clicks=0,
                                                    children='Load Settings',
                                                    style=button_style
                                                )
                                            ],
                                            style=center_placement_style
                                        ),
                                        html.P(''),
                                        html.Div(
                                            children=[
                                                html.Button(
                                                    id='settings-save-button',
                                                    n_clicks=0,
                                                    children='Save Settings',
                                                    style=button_style
                                                )
                                            ],
                                            style=center_placement_style
                                        ),
                                        html.P(''),
                                        html.Div(
                                            children=[
                                                html.Button(
                                                    id='default-settings-load-button',
                                                    n_clicks=0,
                                                    children='Load Default Settings',
                                                    style=button_style
                                                )
                                            ],
                                            style=center_placement_style
                                        )
                                    ],
                                    style=input_area_style
                                ),
                                html.Div(
                                    id='settings-separator-area',
                                    style=separator_area_style
                                ),
                                html.Div(
                                    id='settings-output-area',
                                    children=[
                                        dcc.Loading(
                                            id='settings-output-loading',
                                            children=[
                                                html.Div(
                                                    id='left-separator-settings-output-area',
                                                    style=separator_area_style
                                                ),
                                                html.Div(
                                                    id='left-settings-output-area',
                                                    children=[
                                                        html.H4(
                                                            'Data Shape Definitions',
                                                            id='settings-shape'),
                                                        html.P(
                                                            id='settings-shape-main-header-row-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-shape-main-header-row-text'
                                                        ),
                                                        generate_slider(
                                                            'settings-shape-main-header-row-slider'
                                                        ),
                                                        html.P(
                                                            id='settings-shape-props-header-start-col-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-shape-props-header-start-col-text'
                                                        ),
                                                        generate_slider(
                                                            'settings-shape-props-header-start-col-slider'
                                                        ),
                                                        html.P(
                                                            id='settings-shape-props-header-tally-first-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-shape-props-header-tally-first-text'
                                                        ),
                                                        generate_slider(
                                                            'settings-shape-props-header-tally-first-slider'
                                                        ),
                                                        html.P(
                                                            id='settings-shape-no-summary-table-rows-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-shape-no-summary-table-rows-text'
                                                        ),
                                                        generate_slider(
                                                            'settings-shape-no-summary-table-rows-slider'
                                                        ),
                                                        html.P(
                                                            id='settings-shape-summary-table-sum-row-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-shape-summary-table-sum-row-text'
                                                        ),
                                                        generate_slider(
                                                            'settings-shape-summary-table-sum-row-slider'
                                                        ),
                                                        html.P(
                                                            id='settings-shape-props-header-end-col-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-shape-props-header-end-col-text'
                                                        ),
                                                        generate_slider(
                                                            'settings-shape-props-header-end-col-slider'
                                                        ),

                                                    ],
                                                    style=left_output_area_style
                                                ),
                                                html.Div(
                                                    id='separator-settings-output-area',
                                                    style=separator_area_style
                                                ),
                                                html.Div(
                                                    id='right-settings-output-area',
                                                    children=[
                                                        html.H4(
                                                            'Data Naming Definitions',
                                                            id='settings-names'),
                                                        html.P(
                                                            id='settings-names-sheet-name-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-names-sheet-name-text'
                                                        ),
                                                        dcc.Input(
                                                            id='settings-names-sheet-name-input',
                                                            type='text',
                                                            placeholder='input text',
                                                            debounce=True
                                                        ),
                                                        html.P(
                                                            id='settings-names-country-col-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-names-country-col-text'
                                                        ),
                                                        dcc.Input(
                                                            id='settings-names-country-col-input',
                                                            type='text',
                                                            placeholder='input text',
                                                            debounce=True
                                                        ),
                                                        html.P(
                                                            id='settings-names-store-col-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-names-store-col-text'
                                                        ),
                                                        dcc.Input(
                                                            id='settings-names-store-col-input',
                                                            type='text',
                                                            placeholder='input text',
                                                            debounce=True
                                                        ),
                                                        html.P(
                                                            id='settings-names-main-cols-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-names-main-cols-text'
                                                        ),
                                                        dcc.Dropdown(
                                                            id='settings-names-main-cols-dropdown',
                                                            multi=True
                                                        ),
                                                        html.P(
                                                            id='settings-names-entity-list-text'
                                                        ),
                                                        generate_hover_text(
                                                            'settings-names-entity-list-text'
                                                        ),
                                                        dcc.Dropdown(
                                                            id='settings-names-entity-list-dropdown',
                                                            multi=True
                                                        )
                                                    ],
                                                    style=right_output_area_style
                                                )
                                            ]
                                        )
                                    ],
                                    style=output_area_style
                                ),
                                html.P(
                                    id='settings-placeholder'
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
def display_country_whs_filename(country_whs_content, country_whs_filename):
    """
    Display repeat order filename

    Parameters
    ----------
    country_whs_content : str
        File content
    country_whs_filename : str
        Filename

    Returns
    -------
    output : dash_html_components.Div
        Filename
    """
    if None not in (country_whs_content, country_whs_filename):
        return html.Div([country_whs_filename])
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
        Output('settings-names-main-cols-dropdown', 'options'),
        Output('settings-names-entity-list-dropdown', 'options')
    ],
    [
        Input('settings-names-country-col-input', 'value'),
        Input('settings-names-store-col-input', 'value')
    ]
)
def update_settings_dropdown(country_col, store_col):
    global entity_list
    if country_col is not None and store_col is not None:
        main_cols = [{'label': x, 'value': x} for x in [country_col, store_col]]
        entities = [{'label': x, 'value': x} for x in entity_list]
        return main_cols, entities
    else:
        return [], []

@app.callback(
    [
        Output('settings-shape-main-header-row-text', 'children'),
        Output('settings-shape-props-header-start-col-text', 'children'),
        Output('settings-shape-props-header-tally-first-text', 'children'),
        Output('settings-shape-no-summary-table-rows-text', 'children'),
        Output('settings-shape-summary-table-sum-row-text', 'children'),
        Output('settings-shape-props-header-end-col-text', 'children'),
        Output('settings-names-sheet-name-text', 'children'),
        Output('settings-names-country-col-text', 'children'),
        Output('settings-names-store-col-text', 'children'),
        Output('settings-names-main-cols-text', 'children'),
        Output('settings-names-entity-list-text', 'children')
     ],
    [
        Input('settings-shape-main-header-row-slider', 'value'),
        Input('settings-shape-props-header-start-col-slider', 'value'),
        Input('settings-shape-props-header-tally-first-slider', 'value'),
        Input('settings-shape-no-summary-table-rows-slider', 'value'),
        Input('settings-shape-summary-table-sum-row-slider', 'value'),
        Input('settings-shape-props-header-end-col-slider', 'value'),
        Input('settings-names-sheet-name-input', 'value'),
        Input('settings-names-country-col-input', 'value'),
        Input('settings-names-store-col-input', 'value'),
        Input('settings-names-main-cols-dropdown', 'value'),
        Input('settings-names-entity-list-dropdown', 'value')
    ]
)
def update_settings(main_header_row, props_header_start_col, props_header_tally_first, no_summary_table_rows, 
                    summary_table_sum_row, props_header_end_col, sheet_name, country_col, store_col, main_cols, entity_list):
    # load settings
    global settings
    
    # define texts
    main_header_row_text = 'SkipRows to Main Header'
    props_header_start_col_text = 'SkipCols to First Props Column'
    props_header_tally_first_text = 'Number of First X Props Columns to Tally'
    no_summary_table_rows_text = 'Number of Rows of Summary Table'
    summary_table_sum_row_text = 'Summary Table Sum Row Number '
    props_header_end_col_text = 'Number of Last X Props Columns to Ignore'
    sheet_name_text = 'Sheet to Read From'
    country_col_text = 'Column name of Country'
    store_col_text = 'Column name of Stores'
    main_cols_text = 'Main Columns'
    entity_list_text = 'Entities to find default sheet name'
    
    # update texts
    if main_header_row is not None:
        main_header_row_text += ': %s'%main_header_row
    if props_header_start_col is not None:
        props_header_start_col_text += ': %s'%props_header_start_col
    if props_header_tally_first is not None:
        props_header_tally_first_text += ': %s'%props_header_tally_first
    if no_summary_table_rows is not None:
        no_summary_table_rows_text += ': %s'%no_summary_table_rows
    if summary_table_sum_row is not None:
        summary_table_sum_row_text += ': %s'%summary_table_sum_row
    if props_header_end_col is not None:
        props_header_end_col_text += ': %s'%props_header_end_col
    if sheet_name is not None:
        sheet_name_text += ': %s'%sheet_name
    if country_col is not None:
        country_col_text += ': %s'%country_col
    if store_col is not None:
        store_col_text += ': %s'%store_col
    if main_cols is not None:
        main_cols_text += ': %s'%main_cols
    if entity_list is not None:
        entity_list_text += ': %s'%entity_list
    
    # update settings (no output)
    if len(settings) > 0:
        settings['shape']['main_header_row'] = main_header_row
        settings['shape']['props_header_start_col'] = props_header_start_col
        settings['shape']['props_header_tally_first'] = props_header_tally_first
        settings['shape']['no_summary_table_rows'] = no_summary_table_rows
        settings['shape']['summary_table_sum_row'] = summary_table_sum_row
        settings['shape']['props_header_end_col'] = props_header_end_col
        settings['names']['sheet_name'] = sheet_name
        settings['names']['country_col'] = country_col
        settings['names']['store_col'] = store_col
        settings['names']['main_cols'] = main_cols
        settings['names']['entity_list'] = entity_list
        
    # return texts
    return main_header_row_text, props_header_start_col_text, props_header_tally_first_text, no_summary_table_rows_text, \
           summary_table_sum_row_text, props_header_end_col_text, sheet_name_text, country_col_text, store_col_text, \
           main_cols_text, entity_list_text

@app.callback(
    Output('settings-placeholder', 'children'),
    [
        Input('settings-save-button', 'n_clicks')
    ]
)
def save_settings(save_settings_clicks):
    """"""
    global current_save_settings_clicks
    global settings
    if save_settings_clicks > current_save_settings_clicks:
        filename = settings_path
        write_json(settings, filename)
        current_save_settings_clicks = save_settings_clicks
    return None

@app.callback(
    [
        Output('settings-shape-main-header-row-slider', 'value'),
        Output('settings-shape-props-header-start-col-slider', 'value'),
        Output('settings-shape-props-header-tally-first-slider', 'value'),
        Output('settings-shape-no-summary-table-rows-slider', 'value'),
        Output('settings-shape-summary-table-sum-row-slider', 'value'),
        Output('settings-shape-props-header-end-col-slider', 'value'),
        Output('settings-names-sheet-name-input', 'value'),
        Output('settings-names-country-col-input', 'value'),
        Output('settings-names-store-col-input', 'value'),
        Output('settings-names-main-cols-dropdown', 'value'),
        Output('settings-names-entity-list-dropdown', 'value')
    ]
    ,
    [
        Input('settings-load-button', 'n_clicks'),
        Input('default-settings-load-button', 'n_clicks')
    ]
)
def load_settings(load_settings_clicks, load_default_settings_clicks):
    global current_load_settings_clicks
    global current_load_default_settings_clicks
    global settings
    if load_settings_clicks is not None and load_settings_clicks > 0 and load_settings_clicks > current_load_settings_clicks:
        print('[Status]', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ' Loading the settings ...')
        filename = settings_path
        values = read_json(filename)
        current_load_settings_clicks = load_settings_clicks
    elif load_default_settings_clicks is not None and load_default_settings_clicks > 0 and load_default_settings_clicks > current_load_default_settings_clicks:
        print('[Status]', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ' Loading the default settings ...')
        values = VMPropsManager().get_default_parameters()
        current_load_default_settings_clicks = load_default_settings_clicks
    else:
        values = None
    print('[Status]', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ' Loaded settings: ', values)
    if values is not None:
        settings = values
        return \
            settings['shape']['main_header_row'],\
            settings['shape']['props_header_start_col'], \
            settings['shape']['props_header_tally_first'], \
            settings['shape']['no_summary_table_rows'], \
            settings['shape']['summary_table_sum_row'], \
            settings['shape']['props_header_end_col'], \
            settings['names']['sheet_name'], \
            settings['names']['country_col'], \
            settings['names']['store_col'], \
            settings['names']['main_cols'], \
            settings['names']['entity_list']
    else:
        return None, None, None, None, None, None, None, None, None, [], []


@app.callback(
    [
        Output('so-format-datatable', 'data'),
        Output('so-format-datatable', 'columns'),
        Output('so-format-datatable', 'selected_rows'),
        Output('checked-datatable', 'data'),
        Output('checked-datatable', 'columns'),
        Output('checked-datatable', 'selected_rows'),
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
def run_analysis(vm_props_order_summary_content, country_whs_content, props_batch_content, start_analysis_clicks,
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
    start_analysis_clicks : int
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
    global current_start_analysis_clicks
    global so_format_data
    global checked_data
    global settings
    global sheet_name
    
    so_format_datatable_data = {}
    so_format_datatable_columns = []
    checked_datatable_data = {}
    checked_datatable_columns = []
    download_report_output = []
    are_outputs_available = False

    if start_analysis_clicks > 0 and start_analysis_clicks > current_start_analysis_clicks:
        filename = settings_path
        settings = read_json(filename)
        print('[Status]', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ' Updating the settings ...')
        
        # Run checking
        if None not in (vm_props_order_summary_content, vm_props_order_summary_filename):
            vm_props_order_summary_file = io.BytesIO(base64.b64decode(vm_props_order_summary_content.split(',')[-1]))
        else:
            vm_props_order_summary_file = None
            
        if None not in (vm_props_order_summary_file, vm_props_order_summary_filename):
            # Initialise
            vm = VMPropsManager(settings)
            # Run analysis
            print('[Status]', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ' Updating the settings ...')

            data, data_sh_colours, sheet_name = vm.load_dataset(vm_props_order_summary_file, vm_props_order_summary_filename)
            main_data = vm.get_main_data(data)
            main_data = vm.dropna_rows_cols(main_data)
            # load parameters if not specified
            col_name_list = vm.get_index_to_split_tables(main_data)
            if len(col_name_list) == 1:
                data_1, data_2 = vm.get_split_data(main_data, col_name_list)
            elif len(col_name_list) == 2:
                data_1, data_2, data_3 = vm.get_split_data(main_data, col_name_list)
            data_1_clean = vm.clean_main_data(data_1.copy())
            col_name_list = vm.get_index_to_split_tables2(data_1_clean)
            data_1_head, data_1_body = vm.get_split_data(data_1_clean, col_name_list)
            summary_df = vm.shorten_table_w_max_rows(data_2)
            df = vm.format_main_data(data_1_body)
            checked_data = vm.main_and_summary_checker(df, summary_df)
            so_table = vm.main_table_to_so_converter(df)
            so_format_data = vm.get_cell_colour_col(so_table, data_sh_colours)
        # Format outputs
        if so_format_data is not None and not so_format_data.empty:
            # table not showing until second click
            so_format_datatable_data = so_format_data.to_dict('rows')
            so_format_datatable_columns = [{'name': i, 'id': i} for i in so_format_data.columns]
            are_outputs_available = True
        if checked_data is not None and not checked_data.empty:
            # table not showing until second click; why?
            checked_datatable_data = checked_data.to_dict('rows')
            checked_datatable_columns = [{'name': i, 'id': i} for i in checked_data.columns]
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
        current_start_analysis_clicks = start_analysis_clicks
        
    return so_format_datatable_data, so_format_datatable_columns, [], \
           checked_datatable_data, checked_datatable_columns, [], download_report_output


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
    global checked_data
    global sheet_name

    def format_and_save_excel(summary_df, so_table, keep_cols=None):

        buffer = io.BytesIO()

        if keep_cols is not None:
            so_table = so_table[keep_cols]

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(buffer, engine='xlsxwriter')

        # Convert dataframes to an XlsxWriter Excel object. 
        so_table.to_excel(writer, sheet_name='SO_Table', encoding='utf8')
        summary_df.to_excel(writer, sheet_name='Summary', encoding='utf8')

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['SO_Table']

        for i in list(so_table[so_table['Cell_Colour'] != '00000000'].index):
            hex_code = '#' + str(so_table['Cell_Colour'].loc[i][-6:])
            cell_format = workbook.add_format()
            cell_format.set_pattern(1)
            cell_format.set_bg_color(hex_code)
            worksheet.set_row(i + 1,  # +1 due to cells start from 1 but python 0
                              None,  # do not change row height
                              cell_format  # add bg colour
                              )

        for i in list(so_table[so_table['COUNTRY NAME'] == 'TOTAL'].index):
            cell_format = workbook.add_format({'bold': True, 'border': 3})
            cell_format.set_pattern(1)
            cell_format.set_bg_color('#e5e5e5')
            worksheet.set_row(i + 1,  # +1 due to cells start from 1 but python 0
                              None,  # do not change row height
                              cell_format  # add bold and grey bg for row
                              )

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        buffer.seek(0)
        return buffer

    buffer = format_and_save_excel(checked_data, so_format_data)
    vm = VMPropsManager()

    return send_file(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=vm.get_file_name(sheet_name),
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



