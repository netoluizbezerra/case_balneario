#!/usr/bin/env python
"""""Real Estate Analysis"""
import pandas as pd
import numpy as np
from recordtype import recordtype
import collections
venda = recordtype('venda', ['id', 'tipologia', 'num_unds', 'm2', 'area_venda', 'preco_unit'])
obra = recordtype('obra', ['andar', 'tipologia', 'privativa', 'coberto', 'descoberto',
                           'equiv_coberto', 'equiv_desc', 'total', 'custo_raso', 'custo_cheio'])
# sensitivity_param_cc -> Construction Cost; sensitivity_param_sp -> Selling Price
# sensitivity_test_cc -> Construction Cost; sensitivity_test_sp -> Selling Price

class Empreendimento:
    def __init__(self, file, sensitivity_test_cc=False, sensitivity_test_sp=False, sensitivity_param_cc=1,
                 sensitivity_param_sp=1):
        self.file = file
        data = pd.read_excel(self.file, index_col=None, sheet_name='produto')
        initial_info = pd.read_excel(self.file, index_col=None, sheet_name='informacoes_iniciais')
        rows = data.index
        self.sensitivity_param_sp = sensitivity_param_sp

        if sensitivity_test_sp:
            self._venda = [venda(id=data['id'][row], tipologia=data['tipologia'][row],
                                 num_unds=data['num_unds'][row], m2=data['m2'][row],
                                 area_venda=data['area_venda'][row], preco_unit=data['preco_unit'][row]) for row in
                           rows]
            tipos = self._venda
            for tipo in tipos:
                tipo.preco_unit *= sensitivity_param_sp
            self._venda = tipos

        else:
            self._venda = [venda(id=data['id'][row], tipologia=data['tipologia'][row],
                                 num_unds=data['num_unds'][row], m2=data['m2'][row],
                                 area_venda=data['area_venda'][row], preco_unit=data['preco_unit'][row]) for row in
                           rows]


        data = pd.read_excel(self.file, index_col=None, sheet_name='custo_de_obra')
        rows = data.index
        self._obra = [obra(andar=data['andar'][row], tipologia=data['tipologia'][row],
                           privativa=data['privativa'][row], coberto=data['coberto'][row],
                           descoberto=data['descoberto'][row], equiv_coberto=data['equiv_coberto'][row],
                           equiv_desc=data['equiv_desc'][row], total=data['total'][row],
                           custo_raso=data['custo_raso'][row], custo_cheio=data['custo_cheio'][row]) for row in rows]
        temp = []
        [temp.append(value.custo_cheio) for value in self._obra]

        if sensitivity_test_cc:
            self.custo_cheio = sum(temp) * sensitivity_param_cc
        else:
            self.custo_cheio = sum(temp)

        self.data = pd.read_excel(self.file, index_col=None, sheet_name='premissas_gerais', header=None)
        self.prazo_pre_obra = int(self.data[1][0])
        self.prazo_obra = int(self.data[1][1])
        self.prazo_pos_obra = int(self.data[1][2])

        self.initial_info_list = [(initial_info.iloc[i, 0], initial_info.iloc[i, 1])
                                  for i in initial_info.index]
        self.prazo_total = self.prazo_pre_obra + self.prazo_obra + self.prazo_pos_obra

        financing_simul = pd.read_excel(self.file, sheet_name='premissa_financiamento', header=None)
        self.financing_pairs = {}
        for i in financing_simul.index:
            self.financing_pairs.update({financing_simul.iloc[i][0]: financing_simul.iloc[i][1]})

    def import_data(self):
        try:
            data = pd.read_excel(self.file)
            return data
        except:
            raise ImportError

    def calc_vgv(self):
        temp = []
        [temp.append(value.area_venda * value.preco_unit) for value in self._venda]
        return sum(temp)

    def cronograma_receitas(self):
        premissas_venda = pd.read_excel(self.file, index_col=None, sheet_name='premissa_venda')
        vendas = [(premissas_venda.iloc[i].tipologia, premissas_venda.iloc[i].qtd_mes) for i in premissas_venda.index]
        i0_venda = pd.read_excel(self.file, sheet_name='premissas_gerais', header=None).loc[5][1]

        produto_venda = pd.read_excel(self.file, sheet_name='produto')
        produto = [(produto_venda.loc[i].tipologia, produto_venda.loc[i].num_unds, produto_venda.loc[i].valor_total)
                   for i in produto_venda.index]
        self.lista_header = []
        [self.lista_header.append(produto[i][0]) for i in range(len(produto))]

        lista_temp = []
        for i in range(len(produto)):
            resto = produto[i][1] % vendas[i][1]
            n_meses_venda = int(produto[i][1] / vendas[i][1])
            lista_venda = [0] * int((i0_venda - 1)) + [vendas[i][1]] * n_meses_venda
            lista_venda.append(resto)
            lista_venda = lista_venda + [0] * (self.prazo_total - len(lista_venda))
            lista_temp.append(lista_venda)

        matrix_vendas = pd.DataFrame(lista_temp, index=self.lista_header)
        return matrix_vendas

    def financing_schedule(self):
        matrix_vendas = self.cronograma_receitas()
        produto_venda = pd.read_excel(self.file, sheet_name='produto')
        produto_pairs = {}
        for i in produto_venda.index:
            produto_pairs.update({produto_venda.iloc[i][1]: produto_venda.iloc[i][6] * self.sensitivity_param_sp})

        financing_simul = pd.read_excel(self.file, sheet_name='premissa_financiamento', header=None)
        list_temp = []
        financing_pairs = {}
        headers_temp = []

        for i in financing_simul.index:
            financing_pairs.update({financing_simul.iloc[i][0]: financing_simul.iloc[i][1]})

        for apartment_type in self.lista_header:
            sales_schedule = matrix_vendas.loc[apartment_type]
            for i in range(len(sales_schedule)):
                qty_sales_month = sales_schedule[i]

                if qty_sales_month == 0:
                    continue

                elif qty_sales_month > 1:
                    months_flow = len(sales_schedule) - i - 2
                    list_for_div = np.array([0] * i +  # Months before the sale
                                            [financing_pairs['sinal'] * qty_sales_month] +  # Sinal
                                            [(qty_sales_month / months_flow) * financing_pairs['fluxo']] * months_flow +
                                            [financing_pairs['chaves'] * qty_sales_month])
                    new_list = list_for_div/qty_sales_month

                    for j in range(qty_sales_month):
                        list_temp.append(new_list * produto_pairs[apartment_type])
                        headers_temp.append(apartment_type)

                else:
                    months_flow = len(sales_schedule) - i - 2
                    new_list = np.array([0] * i +  # Months before the sale
                                        [financing_pairs['sinal'] * qty_sales_month] +  # Sinal
                                        [(qty_sales_month / months_flow) * financing_pairs['fluxo']] * months_flow +
                                        [financing_pairs['chaves'] * qty_sales_month])

                    list_temp.append(new_list * produto_pairs[apartment_type])
                    headers_temp.append(apartment_type)

        return pd.DataFrame(list_temp, index=headers_temp)

    def cronograma_despesas(self):
        """ I = Inicial (Parcela Única) / L = Fluxo Linear / F = Fim (Parcela Única) """
        desp_incorp = pd.read_excel(self.file, index_col=None, sheet_name='custo_incorporacao', header=None)
        desp_pos_obra = pd.read_excel(self.file, index_col=None, sheet_name='pos_obra', header=None)
        current_expenses = pd.read_excel(self.file, index_col=None, sheet_name='desp_corrente', header=None)
        broker_commission = pd.read_excel(self.file, sheet_name='premissas_gerais', header=None).loc[7][1]


        temp_headers = []
        [temp_headers.append(desp_incorp[0][i]) for i in desp_incorp.index]
        temp_values = []

        # Development Expenses:
        for i in desp_incorp.index:
            if desp_incorp[2][i] == 'i':
                temp = [desp_incorp[1][i]]
                temp = temp + [0] * (self.prazo_pre_obra - 1) + [0] * self.prazo_obra + [0] * self.prazo_pos_obra
                temp_values.append(temp)
            elif desp_incorp[2][i] == 'l':
                temp = desp_incorp[1][i]
                temp = [temp * (1 / self.prazo_pre_obra)] * int(self.prazo_pre_obra) + [0] * self.prazo_obra + \
                       [0] * self.prazo_pos_obra
                temp_values.append(temp)
            elif desp_incorp[2][i] == 'f':
                temp = [desp_incorp[1][i]]
                temp = [0] * (self.prazo_pre_obra - 1) + temp + [0] * self.prazo_obra + [0] * self.prazo_pos_obra
                temp_values.append(temp)

        c_obra_timeline = list(np.array([0] * self.prazo_pre_obra + [1/self.prazo_obra] * int(self.prazo_obra) +
                                        [0] * self.prazo_pos_obra) * self.custo_cheio)

        # Construction Expenses
        temp_headers.append('Custo de Obra')
        temp_values.append(c_obra_timeline)

        [temp_headers.append(desp_pos_obra[0][i]) for i in desp_pos_obra.index]

        for i in desp_pos_obra.index:
            if desp_pos_obra[2][i] == 'i':
                temp = [desp_pos_obra[1][i]]
                temp = [0] * self.prazo_pre_obra + [0] * self.prazo_obra + temp + [0] * (self.prazo_pos_obra - 1)
                temp_values.append(temp)
            elif desp_pos_obra[2][i] == 'l':
                temp = desp_pos_obra[1][i]
                temp = [0] * self.prazo_pre_obra + [0] * self.prazo_obra + \
                       [temp * (1 / self.prazo_pos_obra)] * int(self.prazo_pos_obra)
                temp_values.append(temp)
            elif desp_pos_obra[2][i] == 'f':
                temp = [desp_pos_obra[1][i]]
                temp = [0] * self.prazo_pre_obra + [0] * self.prazo_obra + [0] * (self.prazo_pos_obra - 1) + temp
                temp_values.append(temp)

        # Expenses with Brokerage

        sales_matrix = np.array(self.cronograma_receitas()).transpose() * \
                       np.array(pd.read_excel(self.file, sheet_name='produto').valor_total)

        sales_matrix = sales_matrix.transpose()
        broker_commission_vector = sales_matrix.sum(axis=0) * broker_commission
        temp_values.append(broker_commission_vector)
        temp_headers.append('Comissão Corretagem')

        # Current Expenses:
        current_expenses_values = []
        current_expenses_index = []
        for i in current_expenses.index:
            current_expenses_values.append([current_expenses.iloc[i][1]] * self.prazo_total)
            current_expenses_index.append(current_expenses.iloc[i][0])

        for i in range(len(current_expenses_values)):
            temp_values.append(current_expenses_values[i])
            temp_headers.append(current_expenses_index[i])

        df = pd.DataFrame(temp_values, index=temp_headers)

        return df

    def __repr__(self):
        return "O arquivo processado pra análise é {}".format(self.file)

    def __str__(self):
        return "file".format()


if __name__ == "__main__":
    file = input("Input File: ")
    projeto = Empreendimento(file)
    print(projeto.file)
    print(projeto.import_data())


