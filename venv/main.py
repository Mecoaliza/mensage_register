from script import iniciar_siat, read_excel, registrar_mensagens_siat
import time

excel_path = "Mensagem_SIAT.xlsx"
path_imgs = "C:\\Users\\user\\Documents\\Python\\images\\"

def main():

    abrir_siat = OpenAppLogin("passwords", "user")
    abrir_siat.enter_acclient()
    time.sleep(10)

    entrar_siat = EnterSiat()
    entrar_siat.find_siattext(path_imgs, "SIAT")
    time.sleep(15)

    atalho = DigitKeysboard()
    atalho.msg_keys()
    time.sleep(2)


    lista_contas, df = read_excel(excel_path)

    if lista_contas:
        registrar_mensagens_siat(lista_contas, df, excel_path)
    else:
        print("Nenhuma conta para processar.")

if __name__ == "__main__":
    main()