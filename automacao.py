
import pandas as pd
import pyautogui
import time

                                                  
def preencher_campo(dado, intervalo=0.1):                                                  
    """     

    Preenche um campo simulando a digitação.
    :param dado: O dado que será digitado.
    :param intervalo: O intervalo entre as teclas (segundos).
    """
    pyautogui.typewrite(dado, interval=intervalo)

def preencher_data(dado, intervalo=0.1):
    # Certifique-se de que dado é uma string
    dado_str = dado.strftime('%m/%d/%Y')  # Exemplo de formato de data: dia/mês/ano
    pyautogui.typewrite(dado_str, interval=intervalo)

df = pd.read_excel('C:PATH')
for index, row in df.iterrows():                    
    # O foco da página já deve estar na página de cadastro da Oracle Cloud.
    time.sleep(10)
    pyautogui.press('tab')
    pyautogui.press('tab')

    preencher_campo(row['Nome'])

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    preencher_campo(row['Ativo ou Item'])

    pyautogui.press('tab')
    time.sleep(2)
    pyautogui.press('tab')

    preencher_data(row['Data Inicial'])

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.press('space')
    time.sleep(2)

    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.press('space')

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.press('down')

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    preencher_campo(str(row['BaseDias']))

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.press('space')
    time.sleep(2)

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('down')
    time.sleep(2)

    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(2)

    pyautogui.press('tab')
    pyautogui.press('tab')
    preencher_campo(row['CodigoDefTrab'])
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(1)

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(1)

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('down')
    time.sleep(1)
    pyautogui.press('enter')

# Avisar que o processo foi concluído
print("Cadastro concluído com sucesso!")
