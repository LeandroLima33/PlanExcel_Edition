from openpyxl import load_workbook

planilha = load_workbook(filename='Digitação Inovação Completo com CNAE.xlsx')

digitacao = planilha['Digitação']



class Excel():
    def __init__(self, v1=str, v2=str):
        self.campo = digitacao[v1:v2]
        self.lista = {}
        print('Campo adicionado')
        self.adiciona_valores()




    def adiciona_valores(self):
        for i in self.campo:
            self.lista[i.row] = i.value



tel1_digitacao = Excel('e', 'e')

print(tel1_digitacao.lista)
        