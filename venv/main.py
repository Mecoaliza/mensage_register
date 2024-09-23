from script import start_siat, read_excel, register_mensage_siat
import time
from dotenv import load_dotenv
import os

load_dotenv()

def main():

    login = os.getenv("LOGIN")
    password = os.getenv("PASSWORD")
    excel_path = os.getenv("EXCEL_PATH")
    img_path = os.getenv("IMG_PATH")

    abrir_siat = OpenAppLogin(login, password)
    abrir_siat.enter_acclient()
    time.sleep(10)

    entrar_siat = EnterSiat()
    entrar_siat.find_siattext(img_path, "SIAT")
    time.sleep(15)

    atalho = DigitKeysboard()
    atalho.msg_keys()
    time.sleep(2)


    lista_contas, df = read_excel(excel_path)

    if lista_contas:
        register_mensage_siat(lista_contas, df, excel_path)
    else:
        print("Nenhuma conta para processar.")

if __name__ == "__main__":
    main()