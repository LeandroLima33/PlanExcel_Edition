import json

class Excel:
    def __init__(self, p, v1=str, v2=str):
        self.campo = p[v1:v2]
        self.lista = {}
        self.lista_iguais = {}
        self.contador = {}
       
        print('Campo adicionado')
        self.adiciona_valores()

        
    def busca_duplicatas(self, i=dict):

       
    
        for l,v in self.lista.items():
            for l1, v1 in i.items():
                if v == v1:
                    print(f'{l1} {v1} Duplicado')
                    print('{v1} Adicionado na lista')
                    self.contador[l1] = v1
                    self.lista_iguais[l1] = v1
                  
           
             
                else:
                    print(l,'<= Linha planilha |', 'Duplicatas encontradas =>', len(self.contador))
        
        json_object = json.dumps(self.lista_iguais, indent=4)
        with open('telefones', 'w') as arquivo_saida:
            arquivo_saida.write(json_object)
            print('Arquivo finalizado')





    def adiciona_valores(self):
        for i in self.campo:
            self.lista[i.row] = str(i.value)
