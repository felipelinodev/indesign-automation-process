import customtkinter
import json
customtkinter.set_appearance_mode("dark")
import os

#JANELA


def MainConfigFunction():
    #TITULO
    janela = customtkinter.CTk()
    janela.geometry("280x430")

    titulo = customtkinter.CTkLabel(janela, text="CONFIGURAÇÃO", font=("Arial Bold", 25))
    titulo.grid(column=1, row=1, padx=30, pady=30, sticky="w")  # Alinhado à esquerda


    #---------------------- CAMPO ATALHOS -------------------------
    titulo = customtkinter.CTkLabel(janela, text="ATALHOS INDESIGN ", font=("Arial Bold", 15))
    titulo.grid(column=1, row=2, padx=40, pady=7, sticky="w")  # Alinhado à esquerda

    # SELECIONAR
    label1 = customtkinter.CTkLabel(janela, text="SELECIONAR: ", font=("Arial Bold", 10), anchor="w")
    label1.grid(column=1, row=3, padx=40, sticky="w")

    atalho_selecionar = None
    atalho_selecionar = customtkinter.CTkEntry(janela, placeholder_text="Ex: f11", width=80, height=40)
    atalho_selecionar.grid(column=1, row=4, padx=40, sticky="w")

    # ATALHO ALINHAR
    label1 = customtkinter.CTkLabel(janela, text="ALINHAR: ", font=("Arial Bold", 10), anchor="w")
    label1.grid(column=1, row=6, padx=40, sticky="w")

    atalho_alinhar = None
    atalho_alinhar = customtkinter.CTkEntry(janela, placeholder_text="Ex: f11", width=80, height=40)
    atalho_alinhar.grid(column=1, row=7, padx=40, sticky="w")

    # ATALHO ATUALIZAR TUDO
    label1 = customtkinter.CTkLabel(janela, text="ATUALIZAR TUDO: ", font=("Arial Bold", 10), anchor="w")
    label1.grid(column=1, row=9, padx=40, sticky="w")

    atalho_atulizar_tudo = None
    atalho_atulizar_tudo = customtkinter.CTkEntry(janela, placeholder_text="Ex: f11", width=80, height=40)
    atalho_atulizar_tudo.grid(column=1, row=10, padx=40, sticky="w")

    
    def AddConfigInJasonFile():
        selecionar = atalho_selecionar.get()
        alinhar = atalho_alinhar.get()
        AtualizarTudo = atalho_atulizar_tudo.get()
        
        def SalveConsigJason():
            # LER O JASON
            # O with garante que o arquivo seja fechado adequadamente mesmo se der merda
            with open('atalhos.json', 'r', encoding='utf-8') as file:
                data = json.load(file)
                
            data["Selecionar"] = selecionar 
            data["Alinhar"] = alinhar
            data["Atualizar Tudo"] = AtualizarTudo
            
            # RESCREVE O JASON
            with open('atalhos.json', 'w', encoding='utf-8') as json_file:
                json.dump(data, json_file, ensure_ascii=False, indent=4)
        SalveConsigJason()
        
    btn_formatadr = customtkinter.CTkButton(janela, text="SALVAR", width=200, height=40, fg_color="#1071E5", command=AddConfigInJasonFile)
    btn_formatadr.grid(column=1, row=11, pady=10, padx=40, sticky="w")

    janela.mainloop()