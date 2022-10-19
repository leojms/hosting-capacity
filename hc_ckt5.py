#importação de bibliotecas
import py_dss_interface
import random
import functions
import numpy as np
import pandas as pd
import math
from bokeh.io import show, curdoc
from bokeh.layouts import gridplot
from bokeh.plotting import figure, output_file
from bokeh.models import Range1d
from bokeh.themes import built_in_themes

u_senai = "leonardo.simoes"
u_note = "leona"
diretorio_excel = fr'C:\Users\{u_senai}\OneDrive - Sistema FIEB\centro_comp\host_cap'

#localização do arquivo
arquivo = r"C:\Users\leonardo.simoes\OneDrive - Sistema FIEB\centro_comp\host_cap\circuito\ckt5\Master_ckt5.dss"

#criação do objeto para referenciar ao OpenDSS
dss = py_dss_interface.DSSDLL()

#tensão padronizada do equivalente de thevenin
thevenin_pu = 1.045
#percentual de redução da carga
load_mult = 0.2
#fator de potencia
pf=1
#passo de crescimento da potência em kva
p_step = 5 #kva
#....duvida
kva_to_kw = 1
#Valor inteiro para padronizar a randomização da distribuição de fotovoltaicos
valor_seed = 1685

perc_barras_list = list()
fv_kva = list()
fv_kw = list()
fv_kvar = list()
thev_kw_com_fv = list()
thev_kvar_com_fv = list()
tensao_min = list()
tensao_max = list()
sobretensao_lista = list()
sobrecarga_lista = list()
i_list = list()
barras_fv = list()
q_barras = list()

for barras_perc in np.arange(0.1, 1.1, 0.1):
    #limpar a memoria e compilar o arquivo
    dss.text("clear")
    dss.text(f"Compile [{arquivo}]")

    #Editando impedancia para um valor pequeno 
    #de elementos que ficam próximos a fonte, a fim de manter a tensão de entrada sem variações
    dss.text("edit Reactor.MDV_SUB_1_HSB x=0.0000001 r=0.0000001")
    dss.text(r"edit Transformer.MDV_SUB_1 %loadloss=0.000001 xhl=0.0000001")
    #modificando a tensão pu da fonte
    dss.text(f"edit Vsource.SOURCE pu={thevenin_pu}")
    #redução da carga para uma porcentagem da instalada, a fim de aumentar a potência fluindo nas linhas 
    dss.text(f"set loadmult = {load_mult}")

    dss.solution_solve()

    #todas as barras do circuito trifásicas de média tensão
    barras_hc_geral = list()
    #tensoes associadas as barras do circuito
    barras_kv = dict()
    #todas as barras
    barras = dss.circuit_all_bus_names()

    #loop para mapear as barras
    for barra in barras:
        #setar a barra que vai ser avaliada
        dss.circuit_set_active_bus(barra)
        #mapeando a tensão da barra
        barra_kv = dss.bus_kv_base()
        #avaliando o número de fases
        num_fases = len(dss.bus_nodes())
        
        #condicional para avaliar se são trifásicas e de média tensão
        if barra_kv > 1.0 and num_fases == 3 and barra!= 'sourcebus':
            barras_hc_geral.append(barra)
            barras_kv[barra] = barra_kv

    #selecionando um percentual das barras para fazer HC
    random.seed(valor_seed)        
    barras_hc = random.sample(barras_hc_geral, int(barras_perc*len(barras_hc_geral)))

    #aplicando um sistema FV nas barras que foram escolhidas
    barras_fv_atual = list()
    for barra in barras_hc:
        functions.define_3ph_pvsystem(dss=dss, bus=barra, kv=barras_kv[barra], kva=kva_to_kw * p_step, pmpp = p_step)
        functions.add_bus_marker(dss=dss, bus=barra, color="red", size_marker=4)
        barras_fv_atual.append(barra)
    q_barra = len(barras_fv_atual)

    #interpolando os pontos de coordenada para adicionar as coordenadas as barras que estão sem
    dss.text("interpolate")
    #calculando o fluxo
    dss.solution_solve()

    #flags para indicar se há violação na rede
    sobretensao = False
    sobrecarga = False
    #limite de tensao avaliado
    v_limite = 1.05

    #numero de iterações
    i=0
    #loop de cálculo do HC
    while not sobretensao and not sobrecarga and i<1000:
        #incremento das iterações
        i += 1
        #aumento do sistema fotovoltaico de acordo com o passo definido
        functions.increment_pv_size(dss=dss, p_step=p_step, kva_to_kw=kva_to_kw, pf=pf, i=i)
        #calculo do fluxo
        dss.solution_solve()
        #armazenamento das tensoes
        tensoes = dss.circuit_all_bus_vmag_pu()
        #avaliação do ponto que possui a maior tensão
        v_max = max(tensoes)
        
        #condicional para avaliar a existência de sobretensão
        if v_max > v_limite:
            sobretensao = True

        #lógica para avaliar sobrecarga 
        #Selecionada a primeira linha do circuito
        dss.lines_first()
        #loop para varrer todas as linhas do circuito
        for _ in range(dss.lines_count()):
            #Selecionar a linha que vai ser avaliada
            dss.circuit_set_active_element(f"line.{dss.lines_read_name()}")
            #armazenar a corrente atual da linha
            i_linha = dss.cktelement_currents_mag_ang()
            #armazenar a corrente nominal da linha
            i_nom_linha = dss.lines_read_norm_amps()
            
            #avaliar se a corrente atual é maior que a corrente nominal
            #se for, há sobrecarga
            if (max(i_linha[0:12:2]) / i_nom_linha) > 1:
                sobrecarga = True
                break
        
            #ir para a próxima linha    
            dss.lines_next()

    #Selecionar o valor anterior do fotovoltaico, para não ter violação
    functions.increment_pv_size(dss=dss, p_step=p_step, kva_to_kw=kva_to_kw, pf=pf, i=i-1)
    #calculo do fluxo
    dss.solution_solve()
    perc_barras_list.append(barras_perc)
    fv_kva.append((i-1) * len(barras_hc) * p_step)
    fv_kw.append(functions.get_total_pv_powers(dss)[0]) 
    fv_kvar.append(functions.get_total_pv_powers(dss)[1])
    thev_kw_com_fv.append(-1 * dss.circuit_total_power()[0])
    thev_kvar_com_fv.append(-1 * dss.circuit_total_power()[1])
    v_pu = dss.circuit_all_bus_vmag_pu()
    tensao_min.append(min(v_pu))
    tensao_max.append(max(v_pu))
    sobretensao_lista.append(sobretensao)
    sobrecarga_lista.append(sobrecarga)
    i_list.append(i)
    barras_fv.append(barras_fv_atual)
    q_barras.append(q_barra)
    
dicio = dict()
dicio["percentual de barras"] = perc_barras_list
dicio["potencia fv (kVA)"] = fv_kva
dicio["potencia fv (kW)"] = fv_kw
dicio["potencia fv (kVAr)"] = fv_kvar
dicio["potencia subestacao (kW)"] = thev_kw_com_fv
dicio["potencia subestacao (kVAr)"] = thev_kvar_com_fv
dicio["tensao minima"] = tensao_min
dicio["tensao maxima"] = tensao_max
dicio["sobretensao"] = sobretensao_lista
dicio["sobrecarga"] = sobrecarga_lista
dicio["numero de iteracoes"] = i_list
dicio["barras com fotovoltaico inserido"] = barras_fv
dicio["quantidade de barras utilizadas"] = q_barras
df = pd.DataFrame.from_dict(dicio)
nome = f'Dados HC ckt 5 original.xlsx'
df.to_excel(fr'{diretorio_excel}\{nome}')
print(f"O arquivo de dados {nome} foi salvo em excel com sucesso")

output_file(fr'{diretorio_excel}\hc ckt 5 original.html')
curdoc().theme = 'dark_minimal'
potencia_fv = figure(x_axis_label="Quantidade de barras", title="Potência FV (kW)")
potencia_fv.line(x=df["quantidade de barras utilizadas"], y=df["potencia fv (kW)"], color='white')
potencia_sub = figure(x_axis_label="Quantidade de barras", title="Potência subestação (kW)")
potencia_sub.line(x=df["quantidade de barras utilizadas"], y=df["potencia subestacao (kW)"], color='hotpink')
num_it = figure(x_axis_label="Quantidade de barras", title="Número de iterações")
num_it.line(x=df["quantidade de barras utilizadas"], y=df["numero de iteracoes"], color='lawngreen')
grid_layout = gridplot([[potencia_fv, potencia_sub], [num_it]])
show(grid_layout)
