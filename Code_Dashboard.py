import pandas as pd
import dash
import dash_core_components as dcc
import dash_html_components as html
import plotly.graph_objs as go
import plotly.express as px 
from dash.dependencies import Input, Output


# Read the Excel file
df1 = pd.read_excel('Monthly_visitor.xlsx')

months = df1.iloc[1:, 0]
visitors = df1.iloc[1:, 2]
visitors_21 = df1.iloc[1:, 4]

# Read the Excel file
df2 = pd.read_excel('Nationality_visitor.xlsx')

x_months = df2.iloc[2, 1:]
y_val1 = df2.iloc[3, 1:]
y_val2 = df2.iloc[4, 1:]
y_val3 = df2.iloc[5, 1:]
y_val4 = df2.iloc[6, 1:]
y_val5 = df2.iloc[7, 1:]
y_val1_21 = df2.iloc[35, 1:]
y_val2_21 = df2.iloc[36, 1:]
y_val3_21 = df2.iloc[37, 1:]
y_val4_21 = df2.iloc[38, 1:]
y_val5_21 = df2.iloc[39, 1:]

# Read the Excel file
df3 = pd.read_excel('Additional_info_visitor.xlsx')

purpose =list(df3.iloc[8:13, 0])
size_p = list(df3.iloc[8:13, 2])
size_p_21 = list(df3.iloc[8:13, 4])
age = list(df3.iloc[15:22, 0])
size_a= list(df3.iloc[15:22, 2])
size_a_21 = list(df3.iloc[15:22, 4])
days = list(df3.iloc[24:29, 0])
size_d= list(df3.iloc[24:29, 2])
size_d_21 = list(df3.iloc[24:29, 4])
colors_5 = ['yellowgreen', 'orange', 'gold', 'plum', 'salmon']
colors_7 = ['yellowgreen', 'orange', 'gold', 'plum', 'salmon', 'skyblue', 'pink']

# Create the app
app = dash.Dash(__name__)

# Define the layout
app.layout = html.Div(
    children=[
        html.H1("Travel Dashboard"),
        html.Div(
            id="top-container",
            children=[
                html.Div(
                    id="dropdown-container",
                    children=[
                        dcc.Dropdown(
                            id="year-dropdown",
                            options=[
                                {"label": "2019", "value": "2019"},
                                {"label": "2021", "value": "2021"},
                            ],
                            value="2019",
                        ),
                    ],
                ),
                html.Div(
                    id="graph-container",
                    children=[
                        html.Div(id="line-chart"),
                        html.Div(id="bar-chart"),
                    ],
                ),
                html.Div(
                    id="pie-charts-container",
                    style={"display": "flex", "flex-wrap": "wrap"},
                    children=[
                        html.Div(id="pie-charts"),
                    ],
                ),
            ],
        ),
    ]
)

@app.callback(
    Output("line-chart", "children"),
    Output("bar-chart", "children"),
    [Input("year-dropdown", "value")]
)


def update_graph(year):
    if year == "2019":
        return [
            dcc.Graph(
                figure=px.bar(
                    x=months,
                    y=visitors,
                    labels={"x": "Month", "y": "The number of visitors"},
                    title="The number of Tourists in NZ",
                )
            ),
            dcc.Graph(
                figure=go.Figure(
                    data=[
                        go.Scatter(
                            x=x_months,
                            y=y_val1,
                            mode="lines",
                            name="Australia",
                            line=dict(color="blue"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val2,
                            mode="lines",
                            name="China",
                            line=dict(color="red"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val3,
                            mode="lines",
                            name="America",
                            line=dict(color="black"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val4,
                            mode="lines",
                            name="UK",
                            line=dict(color="orange"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val5,
                            mode="lines",
                            name="Germany",
                            line=dict(color="green"),
                        ),
                    ],
                    layout=go.Layout(
                        title="The number of Tourists by Nationality in NZ",
                        xaxis=dict(title="Month"),
                        yaxis=dict(title="The number of visitors"),
                        showlegend=True,
                        legend=dict(x=0, y=1, bgcolor="rgba(0,0,0,0)"),
                        autosize=True,
                        margin=dict(l=40, r=40, t=40, b=40),
                    ),
                )
            ),
        ]
    elif year == "2021":
        return [
            dcc.Graph(
                figure=px.bar(
                    x=months,
                    y=visitors_21,
                    labels={"x": "Month", "y": "The number of visitors"},
                    title="The number of Tourists in NZ"
                )
            ),
            dcc.Graph(
                figure=go.Figure(
                    data=[
                        go.Scatter(
                            x=x_months,
                            y=y_val1_21,
                            mode="lines",
                            name="Australia",
                            line=dict(color="blue"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val2_21,
                            mode="lines",
                            name="UK",
                            line=dict(color="red"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val3_21,
                            mode="lines",
                            name="America",
                            line=dict(color="black"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val4_21,
                            mode="lines",
                            name="Samoa",
                            line=dict(color="orange"),
                        ),
                        go.Scatter(
                            x=x_months,
                            y=y_val5_21,
                            mode="lines",
                            name="Cook Island",
                            line=dict(color="green"),
                        ),
                    ],
                    layout=go.Layout(
                        title="The number of Tourists by Nationality in NZ",
                        xaxis=dict(title="Month"),
                        yaxis=dict(title="The number of visitors"),
                        showlegend=True,
                        legend=dict(x=0, y=1, bgcolor="rgba(0,0,0,0)"),
                        autosize=True,
                        margin=dict(l=40, r=40, t=40, b=40),
                    ),
                )
            ),
        ]
    else:
        return None


@app.callback(
    Output("pie-charts-container", "children"),
    [Input("year-dropdown", "value")]
)
def update_pie_charts(year):
    if year == "2019":
        return html.Div(
            style={"display": "flex", "justify-content": "space-between"},
            children=[
                dcc.Graph(
                    figure=go.Figure(
                        data=[
                            go.Pie(
                                labels=purpose,
                                values=size_p,
                                marker=dict(
                                    colors=colors_5,
                                ),
                                hole=0.3,
                                hoverinfo="label+percent",
                                textinfo="value",
                                textfont=dict(size=10)
                            )
                        ],
                        layout=go.Layout(
                            title="Travel Purpose",
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                x=0.2,
                                y=-0.2,
                                bgcolor="rgba(0,0,0,0)",
                                traceorder='normal'
                            ),
                        ),
                    ),
                    config={"displayModeBar": False},
                    style={"height": "550px", "width": "550px"}
                ),
                dcc.Graph(
                    figure=go.Figure(
                        data=[
                            go.Pie(
                                labels=age,
                                values=size_a,
                                marker=dict(
                                    colors=colors_7,
                                ),
                                hole=0.3,
                                hoverinfo="label+percent",
                                textinfo="value",
                                textfont=dict(size=10)
                            )
                        ],
                        layout=go.Layout(
                            title="Age group",
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                x=0.1,
                                y=-0.2,
                                bgcolor="rgba(0,0,0,0)",
                                traceorder='normal'
                            ),
                        ),
                    ),
                    config={"displayModeBar": False},
                    style={"height": "550px", "width": "550px"}
                ),
                dcc.Graph(
                    figure=go.Figure(
                        data=[
                            go.Pie(
                                labels=days,
                                values=size_d,
                                marker=dict(
                                    colors=colors_5,
                                ),
                                hole=0.3,
                                hoverinfo="label+percent",
                                textinfo="value",
                                textfont=dict(size=10)
                            )
                        ],
                        layout=go.Layout(
                            title="Length of Stay (Days)",
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                x=0.1,
                                y=-0.2,
                                bgcolor="rgba(0,0,0,0)",
                                traceorder='normal'
                            ),
                        ),
                    ),
                    config={"displayModeBar": False},
                    style={"height": "550px", "width": "550px"}
                )
            ]
        )
    elif year == "2021":
                return html.Div(
            style={"display": "flex", "justify-content": "space-between"},
            children=[
                dcc.Graph(
                    figure=go.Figure(
                        data=[
                            go.Pie(
                                labels=purpose,
                                values=size_p_21,
                                marker=dict(
                                    colors=colors_5,
                                ),
                                hole=0.3,
                                hoverinfo="label+percent",
                                textinfo="value",
                                textfont=dict(size=10)
                            )
                        ],
                        layout=go.Layout(
                            title="Travel Purpose",
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                x=0.2,
                                y=-0.2,
                                bgcolor="rgba(0,0,0,0)",
                                traceorder='normal'
                            ),
                        ),
                    ),
                    config={"displayModeBar": False},
                    style={"height": "550px", "width": "550px"}
                ),
                dcc.Graph(
                    figure=go.Figure(
                        data=[
                            go.Pie(
                                labels=age,
                                values=size_a_21,
                                marker=dict(
                                    colors=colors_7,
                                ),
                                hole=0.3,
                                hoverinfo="label+percent",
                                textinfo="value",
                                textfont=dict(size=10)
                            )
                        ],
                        layout=go.Layout(
                            title="Age group",
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                x=0.1,
                                y=-0.2,
                                bgcolor="rgba(0,0,0,0)",
                                traceorder='normal'
                            ),
                        ),
                    ),
                    config={"displayModeBar": False},
                    style={"height": "550px", "width": "550px"}
                ),
                dcc.Graph(
                    figure=go.Figure(
                        data=[
                            go.Pie(
                                labels=days,
                                values=size_d_21,
                                marker=dict(
                                    colors=colors_5,
                                ),
                                hole=0.3,
                                hoverinfo="label+percent",
                                textinfo="value",
                                textfont=dict(size=10)
                            )
                        ],
                        layout=go.Layout(
                            title="Length of Stay (Days)",
                            showlegend=True,
                            legend=dict(
                                orientation="h",
                                x=0.1,
                                y=-0.2,
                                bgcolor="rgba(0,0,0,0)",
                                traceorder='normal'
                            ),
                        ),
                    ),
                    config={"displayModeBar": False},
                    style={"height": "550px", "width": "550px"}
                )
            ]
        )
    else:
        return None


# Run the app
if __name__ == "__main__":
    app.run_server(debug=True, port= 8089)