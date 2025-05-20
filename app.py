import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html
from dash.dependencies import Input, Output

vendas_2020 = pd.read_excel("Base Vendas - 2020.xlsx")
vendas_2021 = pd.read_excel("Base Vendas - 2021.xlsx")
vendas_2022 = pd.read_excel("Base Vendas - 2022.xlsx")
clientes = pd.read_excel("Cadastro Clientes.xlsx")
lojas = pd.read_excel("Cadastro Lojas.xlsx")
produtos = pd.read_excel("Cadastro Produtos.xlsx")
vendas = pd.concat([vendas_2020, vendas_2021, vendas_2022], ignore_index=True)
clientes["Nome Completo"] = clientes["Primeiro Nome"].str.strip() + " " + clientes["Sobrenome"].str.strip()
vendas = vendas.merge(clientes[["ID Cliente", "Nome Completo"]], on="ID Cliente", how="left")
vendas = vendas.merge(produtos, on="SKU", how="left")
vendas = vendas.merge(lojas, on="ID Loja", how="left")

vendas['Data da Venda'] = pd.to_datetime(vendas['Data da Venda'])
vendas['Ano'] = vendas['Data da Venda'].dt.year
vendas['Mês'] = vendas['Data da Venda'].dt.month_name()

app = Dash(__name__)

server = app.server

app.layout = html.Div([
    html.H1("Análise de Vendas - Dashboard Interativo", style={'textAlign': 'center'}),

    html.Div([
        dcc.Dropdown(
            id='ano-filtro',
            options=[{'label': str(ano), 'value': ano} for ano in sorted(vendas['Ano'].unique())],
            multi=True,
            placeholder="Filtrar por ano..."
        ),
        dcc.Dropdown(
            id='produto-filtro',
            options=[{'label': p, 'value': p} for p in sorted(vendas['Produto'].unique())],
            multi=True,
            placeholder="Filtrar por produto..."
        ),
        dcc.Dropdown(
            id='loja-filtro',
            options=[{'label': l, 'value': l} for l in sorted(vendas['Nome da Loja'].unique())],
            multi=True,
            placeholder="Filtrar por loja..."
        ),
        dcc.Dropdown(
            id='cliente-filtro',
            options=[{'label': c, 'value': c} for c in sorted(vendas['Nome Completo'].unique())],
            multi=True,
            placeholder="Filtrar por cliente..."
        ),
        dcc.Dropdown(
            id='tipo-produto-filtro',
            options=[{'label': t, 'value': t} for t in sorted(vendas['Tipo do Produto'].unique())],
            multi=True,
            placeholder="Filtrar por tipo de produto..."
        ),
        dcc.Dropdown(
            id='marca-filtro',
            options=[{'label': m, 'value': m} for m in sorted(vendas['Marca'].unique())],
            multi=True,
            placeholder="Filtrar por marca de produto..."
        )
    ], style={'display': 'grid', 'gridTemplateColumns': '1fr 1fr 1fr', 'gap': '10px', 'margin': '20px'}),

    dcc.Graph(id='vendas-ano'),
    dcc.Graph(id='vendas-cliente'),
    dcc.Graph(id='vendas-produto'),
    dcc.Graph(id='vendas-loja'),
    dcc.Graph(id='vendas-tipo-produto'),
    dcc.Graph(id='vendas-mensais'),

    html.Hr(),
    html.H2("Análise por Tipo e Marca"),
    html.Div([
        dcc.Dropdown(
            id='tipo-produto-dropdown',
            options=[{'label': tipo, 'value': tipo} for tipo in sorted(vendas['Tipo do Produto'].unique())],
            placeholder='Selecione o Tipo de Produto'
        ),
        dcc.Dropdown(
            id='marca-dropdown',
            placeholder='Selecione a Marca',
            multi=True
        )
    ], style={'width': '50%', 'margin': 'auto'}),
    dcc.Graph(id='grafico-vendas-marca')
])

@app.callback(
    Output('marca-dropdown', 'options'),
    Input('tipo-produto-dropdown', 'value')
)
def atualizar_marcas(tipo_produto):
    if tipo_produto:
        marcas = vendas[vendas['Tipo do Produto'] == tipo_produto]['Marca'].unique()
        return [{'label': m, 'value': m} for m in sorted(marcas)]
    return []

@app.callback(
    [Output('vendas-ano', 'figure'),
     Output('vendas-cliente', 'figure'),
     Output('vendas-produto', 'figure'),
     Output('vendas-loja', 'figure'),
     Output('vendas-tipo-produto', 'figure'),
     Output('vendas-mensais', 'figure')],
    [Input('ano-filtro', 'value'),
     Input('produto-filtro', 'value'),
     Input('loja-filtro', 'value'),
     Input('cliente-filtro', 'value'),
     Input('tipo-produto-filtro', 'value'),
     Input('marca-filtro', 'value')]
)
def update_graphs(anos, produtos, lojas, clientes, tipos, marcas):
    df = vendas.copy()
    if anos:
        df = df[df['Ano'].isin(anos)]
    if produtos:
        df = df[df['Produto'].isin(produtos)]
    if lojas:
        df = df[df['Nome da Loja'].isin(lojas)]
    if clientes:
        df = df[df['Nome Completo'].isin(clientes)]
    if tipos:
        df = df[df['Tipo do Produto'].isin(tipos)]
    if marcas:
        df = df[df['Marca'].isin(marcas)]

    fig1 = px.bar(df.groupby('Ano')['Qtd Vendida'].sum().reset_index(), x='Ano', y='Qtd Vendida', title='Vendas por Ano')
    top_clientes = df.groupby('Nome Completo')['Qtd Vendida'].sum().nlargest(10).reset_index()
    fig2 = px.pie(top_clientes, names='Nome Completo', values='Qtd Vendida', title='Top 10 Clientes')
    vendas_prod = df.groupby('Produto')['Qtd Vendida'].sum().nlargest(15).reset_index()
    fig3 = px.bar(vendas_prod, y='Produto', x='Qtd Vendida', orientation='h', title='Top 15 Produtos')
    fig4 = px.bar(df.groupby('Nome da Loja')['Qtd Vendida'].sum().reset_index(), x='Nome da Loja', y='Qtd Vendida', title='Vendas por Loja')
    fig5 = px.treemap(df, path=['Tipo do Produto', 'Marca', 'Produto'], values='Qtd Vendida', title='Vendas por Tipo e Marca')
    vendas_mensais = df.groupby(['Ano', 'Mês'])['Qtd Vendida'].sum().reset_index()
    meses_ordem = ['January', 'February', 'March', 'April', 'May', 'June','July', 'August', 'September', 'October', 'November', 'December']
    vendas_mensais['Mês'] = pd.Categorical(vendas_mensais['Mês'], categories=meses_ordem, ordered=True)
    vendas_mensais = vendas_mensais.sort_values(['Ano', 'Mês'])
    fig6 = px.imshow(pd.pivot_table(vendas_mensais, values='Qtd Vendida', index='Ano', columns='Mês'),
                     title='Heatmap de Vendas Mensais', labels=dict(x="Mês", y="Ano", color="Vendas"))
    return fig1, fig2, fig3, fig4, fig5, fig6

@app.callback(
    Output('grafico-vendas-marca', 'figure'),
    [Input('tipo-produto-dropdown', 'value'),
     Input('marca-dropdown', 'value')]
)
def update_grafico_marca(tipo, marcas):
    df = vendas.copy()
    if tipo:
        df = df[df['Tipo do Produto'] == tipo]
    if marcas:
        df = df[df['Marca'].isin(marcas)]
    df_grouped = df.groupby('Marca')['Qtd Vendida'].sum().reset_index()
    fig = px.bar(df_grouped, x='Marca', y='Qtd Vendida', title=f'Vendas por Marca ({tipo})')
    return fig

if __name__ == '__main__':
    app.run(debug=True)
