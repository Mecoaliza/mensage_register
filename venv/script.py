import pandas as pd
import pyautogui as py
import time
import pygetwindow as gw
import pyperclip
import logging

logging.basicConfig(level=logging.INFO)

def read_excel(filepath):
    try:
        df = pd.read_excel(filepath, dtype={'conta_txt': str})
        if 'conta_txt' in df.columns and 'status' in df.columns:
            contas_pendentes = df[df['status'].isna()]['conta_txt'].tolist()
            return contas_pendentes, df
        else:
            logging.error('As colunas esperadas não foram encontradas no arquivo.')
            return [], None
    except Exception as e:
        logging.error(f'Erro ao ler o arquivo Excel: {e}')
        return [], None
    

def get_mensage_siat():
    try:
        janela_siat = gw.getWindowsWithTitle('teoff-exe')[0]
        janela_siat.activate()
        py.moveTo(janela_siat.left + 250, janela_siat.top + 20)
        py.dragTo(x=250, y=20, duration=0.5, button='left')
        py.moveTo(50, 400)
        py.dragTo(500, 400, duration=1.0)
        time.sleep(2)
        py.rightClick()
        mensagem = pyperclip.paste()
        return mensagem
    except Exception as e:
        logging.error(f"Erro ao copiar mensagem no SIAT: {e}")
        return ""


def start_siat(path_imgs):
    siat_app = OpenAppLogin("password", "user")
    siat_app.enter_acclient()
    time.sleep(10)
    siat_interface = EnterSiat()
    siat_interface.find_siattext(path_imgs, "SIAT")
    time.sleep(15)
    atalho_teclado = DigitKeysboard()
    atalho_teclado.msg_keys()
    time.sleep(2)



def register_mensage_siat(contas, df, nome_arquivo):
    for conta in contas:
        logging.info(f'Registrando conta: {conta}')
        py.write(str(conta))
        time.sleep(2)
        mensagem = get_mensage_siat()

        if "Você deseja Alterar ou Excluir" in mensagem:
            py.press("a")
            time.sleep(2)
            py.press("esc")
            time.sleep(2)
            py.press("e")
            status = "Já existe uma mensagem"

        elif "Digite a primeira linha da mensagem." in mensagem:
            py.write("capital baixado para pgto de dividas.")
            py.press("enter")
            time.sleep(2)
            py.press("enter")
            time.sleep(2)
            py.write("s")
            py.press("esc")
            time.sleep(2)
            py.press("e")
            status = "Mensagem registrada"
        
        else: 
            status = "Já existe uma mensagem"
        
        df.loc[df['conta_txt'] == conta, 'status'] = status
        df.to_excel(nome_arquivo, index=False)
        logging.info(f"Conta {conta} processada e salva com status: {status}")

