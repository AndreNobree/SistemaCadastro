from ast import Return
from mimetypes import init
import openpyxl
from turtle import color
import win32com.client as win32
import mysql.connector
import re

from PySimpleGUI import PySimpleGUI as sg # principal
from PySimpleGUI import PySimpleGUI as erro #erro
from PySimpleGUI import PySimpleGUI as exito # exito
from PySimpleGUI import PySimpleGUI as consulta #para consulta
from PySimpleGUI import PySimpleGUI as deleta #para deletar
from PySimpleGUI import PySimpleGUI as exporta #para exportar
from PySimpleGUI import PySimpleGUI as editar #para perguntar se deseja editar
from PySimpleGUI import PySimpleGUI as editar2 #para editar

class TelaCad:

#ERROS DE PREENCHIMENTO    
    def erros(self):
        erro.theme('DarkRed2')
        print("ERRO")
        layout = [
            [
                sg.popup_ok('ERRO DE DIGITAÇÃO, VERIFIQUE OS CAMPOS')
            ]
        ]

    def erro(self):
        erro.theme('DarkRed2')
        print("ERRO")
        layout = [
            [
                sg.popup_ok('VERIFIQUE SE OS CAMPOS FORAM PREENCHIDOS CORRETAMENTE')
            ]
        ]

#PROGRAMA PRINCIPAL
    def cadastro(self):

        sg.theme('Reddit')

        conexao = mysql.connector.connect(
            host='localhost',
            user='root',
            password='password',
            database='teste'
        )
        cursor = conexao.cursor()

        # colocar o campo CARGO
        layout = [
            [
                sg.Text('NumCart'), sg.Input(key='cartao', size=10),
            
                sg.Text('Membro'), sg.Input(key='membro', size=51),

                sg.Button('Consultar'),
            ],
            [
                sg.Text('Nascimento'), 
                sg.Combo(["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
            "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"]),
                sg.Combo(["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]),
                sg.Combo(["2022", "2021", "2020", "2019", "2018", "2017", "2016", "2015", "2014", "2013", "2012", "2011",
            "2010", "2009", "2008", "2007", "2006", "2005", "2004", "2003", "2002", "2001", "2000", "1999", "1998", "1997", "1996", "1995", 
            "1994", "1993", "1992", "1991", "1990", "1989","1988", "1987", "1986", "1985", "1984", "1983",
            "1982", "1981", "1980", "1979", "1978", "1977", "1976", "1975", "1974", "1973", "1972", "1971", "1970", "1969", "1968", "1967", 
            "1966", "1965", "1964", "1963", "1962", "1961", "1960", "1959", "1958", "1957", "1956", "1955", "1954", "1953", "1952", "1951",
            "1950", "1949", "1948", "1947", "1946", "1945", "1944", "1943", "1942", "1941", "1940", "1939", "1938", "1937", "1936", "1935", 
            "1934", "1933", "1932", "1931", "1930","1929", "1928", "1927", "1926", "1925", "1924", "1923", "1922", "1921", "1920",
            "1919","1918", "1917", "1916", "1915", "1914", "1913", "1912", "1911", "1910",]),

                sg.Text('Admissão'), 
                sg.Combo(["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
            "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"]),
                sg.Combo(["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]),
                sg.Combo(["2022", "2021", "2020", "2019", "2018", "2017", "2016", "2015", "2014", "2013", "2012", "2011",
            "2010", "2009", "2008", "2007", "2006", "2005", "2004", "2003", "2002", "2001", "2000", "1999", "1998", "1997", "1996", "1995", 
            "1994", "1993", "1992", "1991", "1990", "1989","1988", "1987", "1986", "1985", "1984", "1983",
            "1982", "1981", "1980", "1979", "1978", "1977", "1976", "1975", "1974", "1973", "1972", "1971", "1970", "1969", "1968", "1967", 
            "1966", "1965", "1964", "1963", "1962", "1961", "1960", "1959", "1958", "1957", "1956", "1955", "1954", "1953", "1952", "1951",
            "1950", "1949", "1948", "1947", "1946", "1945", "1944", "1943", "1942", "1941", "1940", "1939", "1938", "1937", "1936", "1935", 
            "1934", "1933", "1932", "1931", "1930","1929", "1928", "1927", "1926", "1925", "1924", "1923", "1922", "1921", "1920",
            "1919","1918", "1917", "1916", "1915", "1914", "1913", "1912", "1911", "1910",]),
                sg.Text('Cargo'), 
                sg.Combo(["Membro", "Auxiliar", "Diácono", "Presbítero", "Evangelista", "Pastor"]),
            ],

            [
                sg.Text('  Batismo    '), 
                sg.Combo(["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
            "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"]),
                sg.Combo(["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]),
                sg.Combo(["2022", "2021", "2020", "2019", "2018", "2017", "2016", "2015", "2014", "2013", "2012", "2011",
            "2010", "2009", "2008", "2007", "2006", "2005", "2004", "2003", "2002", "2001", "2000", "1999", "1998", "1997", "1996", "1995", 
            "1994", "1993", "1992", "1991", "1990", "1989","1988", "1987", "1986", "1985", "1984", "1983",
            "1982", "1981", "1980", "1979", "1978", "1977", "1976", "1975", "1974", "1973", "1972", "1971", "1970", "1969", "1968", "1967", 
            "1966", "1965", "1964", "1963", "1962", "1961", "1960", "1959", "1958", "1957", "1956", "1955", "1954", "1953", "1952", "1951",
            "1950", "1949", "1948", "1947", "1946", "1945", "1944", "1943", "1942", "1941", "1940", "1939", "1938", "1937", "1936", "1935", 
            "1934", "1933", "1932", "1931", "1930","1929", "1928", "1927", "1926", "1925", "1924", "1923", "1922", "1921", "1920",
            "1919","1918", "1917", "1916", "1915", "1914", "1913", "1912", "1911", "1910",]),
                sg.Text('  Endereço  '), sg.Input(key='endereco', size=42),
            ],
            [
                sg.Text('  Bairro   '), sg.Input(key='bairro', size=20),

                sg.Text('Telefone'), sg.Input(key='telefone', size=14),

                sg.Text('Telefone2'), sg.Input(key='telefone2', size=14),
                sg.Button('Exportar', button_color=('darkorange')),
            ],
            [
                sg.Text('*OBS: Se caso não houver Telefone, apenas coloque  "0" no campo', text_color='red'),
            ],

            [
                sg.Button('Cadastrar'),
                sg.Button('Limpar', button_color=('Black', 'white')),
                sg.Button('Terminar', button_color=('red'))
            ]
        ]

        
        janela = sg.Window('Cadastro Membros', layout)

        while True:
            self.eventos, self.valores = janela.read()
            print(self.valores)
            if self.eventos == sg.WINDOW_CLOSED:
                break
            
            if self.eventos == 'Terminar':
                break
        #DECLARAÇÃO DE VARIÁVEIS DA PRINCIPAL

                #nome e numero de cartão
            valA = (self.valores['cartao'])
            valB = (self.valores['membro'])
            print('Cartão: '+ valA)
            print('Nome: '+ valB)

                #data nascimento
            valC = (self.valores[0])# dia
            valD = (self.valores[1])# mes
            valE = (self.valores[2])# ano

            valF = (self.valores[3])
            valG = (self.valores[4])
            valH = (self.valores[5])

            valI = (self.valores[7])
            valJ = (self.valores[8])
            valK = (self.valores[9])

                #cargo
            valP = (self.valores[6])
                
                #endereço
            valL = (self.valores['endereco'])
            print('endereço ' + valL)

                #bairro
            valM = (self.valores['bairro'])
            print('bairro ' + valM)

                #telefone
            valN = (self.valores['telefone'])
            print('telefone ' + valN)
            
            if valN == '':
                valN = 0

                #telefone2
            valO = (self.valores['telefone2'])
            print('telefone2 ' + valO)
            if valO == '':
                valO = 0


            if self.eventos == 'Limpar':
                janela['cartao'].update('')
                janela['membro'].update('')
                janela[0].update('')
                janela[1].update('')
                janela[2].update('')
                janela[3].update('')
                janela[4].update('')
                janela[5].update('')
                janela[6].update('')
                janela[7].update('')
                janela[8].update('')
                janela[9].update('')
                janela['endereco'].update('')
                janela['bairro'].update('')
                janela['telefone'].update('')
                janela['telefone2'].update('')

#validação de datas
            try:
            #nascimento
                valC = re.sub('[^0-9]', '', valC)

                valD = re.sub('[^0-9]', '', valD)


                valE = re.sub('[^0-9]', '', valE)


                valT1 = valC + '-' + valD + '-' + valE

                #admissão
                valF = re.sub('[^0-9]', '', valF)


                valG = re.sub('[^0-9]', '', valG)


                valH = re.sub('[^0-9]', '', valH)


                valT2 = valF + '-' + valG + '-' + valH

                #batismo
                valI = re.sub('[^0-9]', '', valI)


                valJ = re.sub('[^0-9]', '', valJ)


                valK = re.sub('[^0-9]', '', valK)

                valT3 = valI + '-' + valJ + '-' + valK
                

            except:
                q = TelaCad().erros()


            if self.eventos == 'Cadastrar':

                if valA != '' and valB != '' and valC != '' and valD != '' and valE != '' and valF != '' and valG != '' and valH != '' and valI != '' and valJ != '' and valK != '' and valL != '' and valM != '' and valP != '':
                   
                    print(self.valores)

                    comando = f'INSERT INTO tabelateste (IDCartao, Nome, Cargo, DataNasc, DataAdmissao, DataBatismo, Endereco, Bairro, Telefone, Telefone2) VALUES ({valA}, "{valB}", "{valP}", "{valT1}", "{valT2}", "{valT3}", "{valL}", "{valM}", {valN}, {valO})'

                    try:
                        cursor.execute(comando)
                        conexao.commit() 

                        exito.theme('DarkGreen1')
                        print("Aceito")

                        layout = [
                            [
                                sg.popup_ok('Dados Cadastrados com Sucesso')
                            ]
                        ]        

#ERRO DE PREENCHIMENTO            
                    except:
                        q = TelaCad().erros()


#***ERRO DE CAMPOS FALTANDO            
                else:
                    q2 = TelaCad().erro()
#CONSULTAR
            if self.eventos == 'Consultar':
                if valB != '':

                    comando = f'SELECT * FROM tabelateste WHERE Nome LIKE "{valB}%"'

                    cursor.execute(comando)        
                    self.resultado = cursor.fetchall()   

                    
                    consulta.theme('DarkTeal12')
                    layout = [
                        [
                            consulta.Text(valB, size=50),
                        ],

                        [
                            consulta.Output(size=(70, 40)),
                        ],

                        [
                            consulta.Button('Mostrar Dados', button_color=('green')),
                            consulta.Button('Alterar', button_color=('darkblue')),
                            consulta.Button('Deletar', button_color=('red')),
                            consulta.Button('Terminar', button_color=('Black', 'white')),
                        ],


                    ]
                    janela4 = consulta.Window(valB, layout)
                    while True:
                        self.eventos4, self.valores4 = janela4.read()

                        if self.eventos4 == 'Terminar':
                            janela4.close()
                            break
                        if self.eventos4 == consulta.WINDOW_CLOSED:
                            janela4.close()
                            break
                        
#CONFIRMAÇÃO PARA EDIÇÃO                 
                        if self.eventos4 == 'Alterar':
                            editar.theme('DarkRed')
                            layout2 = [
                                [
                                    editar.Text('DESEJA ALTERAR ALGUM CONTATO?'), editar.Input(key='edita', size=11),

                                    editar.Text('*Digite  CONFIRMO ', text_color='white'),

                                    editar.Button('Continuar', button_color=('green')),
                                    editar.Button('Cancelar', button_color=('red')),
                                ],
                            ]
                            janela5 = editar.Window('EDITOR', layout2)

                            while True:
                                self.eventos5, self.valores5 = janela5.read()

                                vall = (self.valores5['edita'])

                                if self.eventos5 == 'Cancelar':
                                    janela5.close()
                                    break

                                if self.eventos5 == editar.WINDOW_CLOSED:
                                    janela5.close()
                                    break

                                if self.eventos5 == 'Continuar' and vall == 'CONFIRMO':
                                    janela5.close()

#**EDIÇÃO DE DADOS

                                    editar2.theme('DarkBlue17')

                                    layout3 = [
                                        [
                                            editar2.Text('EDIÇÃO DE DADOS ', font=("Helvetica", 12), text_color='white'),
                                        ],
                                        [
                                            editar2.Text("Cartão", size=15), editar2.Input(key='Card1', size=42),
                                        ],
                                        [
                                            editar2.Text("Nome", size=15), editar2.Input(key='Nome1', size=42),
                                        ],
                                        [
                                            editar2.Text("Cargo", size=15), editar2.Input(key='Cargo1', size=42),
                                        ],
                                        [
                                            editar2.Text("Data Nascimento", size=15), editar2.Input(key='Nasc', size=42),
                                        ],
                                        [
                                            editar2.Text("Data Admissão", size=15), editar2.Input(key='Admi', size=42),
                                        ],
                                        [
                                            editar2.Text("Data Batismo", size=15), editar2.Input(key='Bat', size=42),
                                        ],
                                        [
                                            editar2.Text("Endereço", size=15), editar2.Input(key='End', size=42),
                                        ],
                                        [
                                            editar2.Text("Bairro", size=15), editar2.Input(key='Barr', size=42),
                                        ],
                                        [
                                            editar2.Text("Telefone1", size=15), editar2.Input(key='Tell', size=42),
                                        ],
                                        [
                                            editar2.Text("Telefone2", size=15), editar2.Input(key='Telll', size=42),
                                        ],
                                            
                                        [
                                            editar2.Button('Salvar', button_color=('darkblue')),
                                            editar2.Button('Sair', button_color=('red')),
                                        ],
                                    ] 
                                    janela6 = editar2.Window('EDITOR', layout3)
                                    while True:
                                        self.eventos6, self.valores6 = janela6.read()     
                                        
                                        if self.eventos6 == 'Salvar':      
                                            valQ = (self.valores6['Card1'])
                                            print(valQ)
                                            valR = (self.valores6['Nome1'])
                                            print(valR) 
                                            valS = (self.valores6['Cargo1'])
                                            print(valS) 
                                            valU = (self.valores6['Nasc'])
                                            print(valU) 
                                            valV = (self.valores6['Admi'])
                                            print(valV) 
                                            valX = (self.valores6['Bat'])
                                            print(valX) 
                                            valW = (self.valores6['End'])
                                            print(valW) 
                                            valY = (self.valores6['Barr'])
                                            print(valY) 
                                            valZ = (self.valores6['Tell'])
                                            print(valZ) 
                                            valZ2 = (self.valores6['Telll'])
                                            print(valZ2) 
                                            
                                            #VERIFICAÇÃO DA DATA
                                            #nascimento
                                            valU = re.sub('[^0-9]', '', valU)

                                            #admissão
                                            valV = re.sub('[^0-9]', '', valV)

                                            #batismo
                                            valX = re.sub('[^0-9]', '', valX)
                                            
                                            #COLOCAR 0 SE NÃO HOUVER TELEFONE
                                            if valZ2 == '':
                                                valZ2 = 0
                                            if valZ == '':
                                                valZ = 0    

                                            if valQ != '' and valR != '' and valS != '' and valU != '' and valV != '' and valX != '' and valW != '' and valY != '' and valZ != '' and valZ2 != '':

                                                
                                                comando = f'UPDATE tabelateste SET Nome = "{valR}", Cargo = "{valS}", DataNasc = "{valU}", DataAdmissao = "{valV}", DataBatismo = "{valX}", Endereco = "{valW}", Bairro = "{valY}", Telefone = "{valZ}", Telefone2 = "{valZ2}" WHERE IDCartao = {valQ}'
                                              
                                                try:
                                                    cursor.execute(comando)
                                                    conexao.commit()
                                                    
                                                    exito.theme('DarkGreen1')
                                                    print("Aceito")

                                                    layout = [
                                                        [
                                                            sg.popup_ok('Registro Alterado com Sucesso')
                                                        ]
                                                    ]
                                                except:
                                                    q = TelaCad().erros()
                                            else:
                                                q2 = TelaCad().erro()

                                        if self.eventos6 == editar2.WIN_CLOSED:
                                            janela6.close()
                                            break
                                            
                                        if self.eventos6 == 'Sair':
                                            janela6.close()
                                            break
                                
#mostrar resultado        
                        if self.eventos4 == 'Mostrar Dados':
                            for self.result in self.resultado:
                                print("==========================================")    
                                print("ID CARTÃO: ")
                                print(self.result[1])
                                print("\n")
                                print("NOME: ")
                                print(self.result[2])
                                print("\n")
                                print("Cargo: ")
                                print(self.result[3])
                                print("\n")
                                print("DATA NASCIMENTO: ")
                                print(self.result[4])
                                print("\n")
                                print("DATA ADMISSÃO: ")
                                print(self.result[5])
                                print("\n")
                                print("DATA BATISMO: ")
                                print(self.result[6])
                                print("\n")
                                print("ENDEREÇO: ")
                                print(self.result[7])
                                print("\n")
                                print("BAIRRO: ")
                                print(self.result[8])
                                print("\n")
                                print("TELEFONE: ")
                                print(self.result[9])
                                print("\n")
                                print("TELEFONE2: ")
                                print(self.result[10])
                                print("==========================================")   


                        if self.eventos4 == 'Deletar':
                            deleta.theme('DarkRed')
                            layout2 = [
                                [
                                    deleta.Text('COLOQUE A ID DO CARTÃO DE MEMBRO REGISTRADA QUE DESEJA EXCLUIR'), deleta.Input(key='IdDel', size=11),

                                    deleta.Button('Continuar', button_color=('green')),
                                    deleta.Button('Cancelar', button_color=('red')),
                                ],

                            ]
                            janela5 = deleta.Window('DELETAR?', layout2)

                            while True:
                                self.eventos5, self.valores5 = janela5.read()

                                vall = (self.valores5['IdDel'])

                                if self.eventos5 == 'Continuar' and vall != '':
                                    try:

                                        comando = f'DELETE FROM tabelateste WHERE IdCartao = "{vall}" ' 
                                        cursor.execute(comando)
                                        conexao.commit()  
                                        
                                        janela5.close()
                                        
                                        layout = [
                                            [
                                                sg.popup_ok('O Registro foi Apagado com Sucesso!')
                                            ]
                                        ]
                                    except:
                                        q = TelaCad().erros()
                                elif self.eventos5 == 'Continuar' and vall == '':
                                    q2 = TelaCad().erro()
                                        
                                if self.eventos5 == 'Cancelar':
                                    janela5.close()
                                    break

                                if self.eventos5 == deleta.WINDOW_CLOSED:
                                    janela5.close()
                                    break
                                    
           

            #exportar
            if self.eventos == 'Exportar':
               
                coma = 'SELECT * FROM tabelateste'
                cursor.execute(coma)        
                self.total = cursor.fetchall()  
                
                book = openpyxl.Workbook()
                print(book.sheetnames)
                
                cadieadtc = book['Sheet'] 
                cadieadtc.append(['IDMembro', 'IDCartão', 'Nome', 'Cargo', 'Nascimento', 'Admissão', 'Batismo', 'Endereço', 'Bairro', 'Telefone', 'Telefone 2'])
                try:
                    for self.tot in self.total:
                        cadieadtc.append([self.tot[0],self.tot[1],self.tot[2],self.tot[3],self.tot[4],self.tot[5],self.tot[6],self.tot[7],self.tot[8],self.tot[9],self.tot[10]])
                        exito.theme('DarkGreen1')
                        layout = [
                            [
                                sg.popup_ok('Exportacão Bem Sucedida')
                            ]
                        ]
                except:
                    erro.theme('DarkRed2')
                    layout = [
                        [
                            sg.popup_ok('ERRO NA EXPORTACÃO')
                        ]
                    ]
                    
                book.save('Cadastro IEADTC.xlsx')
                


        cursor.close()
        conexao.close()

p = TelaCad().cadastro()
