import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import plotly.express as px

# Load the Excel data
df = pd.read_excel('Collection Final 1.xlsx', sheet_name='Sheet5')

# Define the custom order for the y-axis
custom_order = ['A', 'B', 'C', 'D', 'E']

month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

# Create Dash app
app = dash.Dash(__name__)
server=app.server

# Define layout
app.layout = html.Div([
    html.H1("Customer Rating Dashboard"),

    # Dropdown menu
    dcc.Dropdown(
        id='dropdown-menu',
        options=[
            {'label': column, 'value': column}
            for column in df.columns
        ],
        value=df.columns[1],  # Default selected item
        style={'width': '50%'}
    ),

    # Line chart
    dcc.Graph(id='line-chart')

])


# Define callback to update the line chart based on the selected item
@app.callback(
    Output('line-chart', 'figure'),
    [Input('dropdown-menu', 'value')]
)
def update_line_chart(selected_item):
    if selected_item not in df.columns:
        # Handle the error, perhaps return a default chart or display an error message.
        return dash.no_update

    figure = px.line(df, x=df.index, y=selected_item, title=f'Rating Line Chart of {selected_item}',
                     labels={'x': 'Month', 'y': 'Grade'},
                     category_orders={'Grade':custom_order,'Month':month_order})
    return figure


# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
