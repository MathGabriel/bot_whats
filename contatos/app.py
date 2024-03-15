#importacao dos pacotes
import openpyxl
# importando o pacote para formatar a url certinho na mensagem
from urllib.parse import quote
# para abrir o navegador
import webbrowser
# parada pra ver a imagem de enviar
import pyautogui

# para pausar a aplicacao por alguns segundos
from time import sleep 
#webbrowser.open("https://web.whatsapp.com/")
#sleep(30)

# 1. Ler planilha e guardar informacoes sobre nome, telefone e data limite
# abrindo a planilha no codigo
planilha = openpyxl.load_workbook('clientes.xlsx')
# especificando a pagina a usar da planilha
pagina = planilha['Pagina1']

#conectar primeiro no whatsapp no navegador
#webbrowser
# for colocando pra começar na linha 2, pois a 1 é de nome de cada coluna
for linha in pagina.iter_rows(min_row=2):
    
    # extrair da planilha o nome, telefone e vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Mostrando os valores pegos
    print(nome)
    print(telefone)
    print(vencimento)

    #Criacao da mensagem formatada para inserir na URL
    mensagem = f'Olá {nome} seus produtos já estão disponíveis para retirada ate o dia {vencimento.strftime("%d/%m/%Y")}\nMas lembrando ao senhores que após o prazo será revendido seu produto'


# 2.Criar links personalizados do whatszapp e enviar mensagens para cada cliente
    
    # Descobrir se algum deu erro
    try: 
        # geração da url para nmr especifico
        # https://web.whatsapp.com/send?phone={}&text={}
        link_msg_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

        # para abrir no navegador usando minha url
        webbrowser.open(link_msg_whats)

        # Pausar para poder a escrita entrar e aparecer o boatao
        sleep(10)

        # salvanvo na variavel a localização do botao de enviar
        seta = pyautogui.locateCenterOnScreen('seta.png')
        
        # Pausa apos achar a seta
        sleep(10) 

        # clicar na ao encontrar o botao de enviar
        pyautogui.click(seta[0],seta[1])
        
        # nova pausa
        sleep(5) 
        
        #fechar a aba
        pyautogui.hotkey('ctrl','w')
        
        # nova pausa
        sleep(5)

    except:
        print(f"Não foi possivel enviar mensagem para {nome}")
        with open ('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}\n')
