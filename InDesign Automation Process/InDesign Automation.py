import win32com.client
import win32gui
import win32con
import os
import pyautogui
import time
import customtkinter
import json
customtkinter.set_appearance_mode("dark")
from  Atalhos_Config import MainConfigFunction

# Criação da janela
janela = customtkinter.CTk()
janela.geometry("790x580")
janela.title(" InDesign Automation")
janela.iconbitmap("logo.ico")

# Configuração do título
titulo = customtkinter.CTkLabel(janela, text="FORMATAÇÃO", font=("Arial Bold", 25))
titulo.grid(column=1, row=2, padx=30, pady=30, sticky="w")  # Alinhado à esquerda

config_button = customtkinter.CTkButton(janela, text="CONFIG", width=170, height=25, fg_color="#353535", hover_color="#212121", command=MainConfigFunction)
config_button.grid(column=1, row=3, padx=30, sticky="w")


# Configuração do subtítulo
subtitulo = customtkinter.CTkLabel(janela, text="PREENCHA OS TODOS CAMPOS", font=("Arial Bold", 15))
subtitulo.grid(column=2, row=2, padx=40, pady=20, sticky="w")  # Alinhado à esquerda


#----------------------CAMPO CÓDIGO-------------------------
# Configuração dos labels e entradas
label1 = customtkinter.CTkLabel(janela, text="DIGITE O CÓDIGO", font=("Arial Bold", 10), anchor="w")
label1.grid(column=2, row=3, padx=40, sticky="w")

campo_codigo = None
campo_codigo = customtkinter.CTkEntry(janela, placeholder_text="Ex: ADV00", width=250, height=40)
campo_codigo.grid(column=2, row=4, padx=40, sticky="w")



#----------------------CAMPOS CAMINHOS-------------------------
label2 = customtkinter.CTkLabel(janela, text="CAMINHO ARQUIVOS IMPRESSÃO", font=("Arial Bold", 10), anchor="w")
label2.grid(column=2, row=5, padx=40, sticky="w")

caminho_arquivos_impres = None
caminho_arquivos_impres = customtkinter.CTkEntry(janela, placeholder_text=r"Ex: C:\CAMISA\IMPRESSÃO", width=250, height=40)
caminho_arquivos_impres.grid(column=2, row=6, padx=40, sticky="w")

label3 = customtkinter.CTkLabel(janela, text="CAMINHO ARQUIVOS LAYOUT", font=("Arial Bold", 10), anchor="w")
label3.grid(column=2, row=7, padx=40, sticky="w")

caminho_arquivos_layout = None
caminho_arquivos_layout = customtkinter.CTkEntry(janela, placeholder_text=r"Ex: C:\CAMISA\LAYOUT", width=250, height=40)
caminho_arquivos_layout.grid(column=2, row=8, padx=40, sticky="w")

label4 = customtkinter.CTkLabel(janela, text="CAMINHO EXPORTAÇÃO", font=("Arial Bold", 10), anchor="w")
label4.grid(column=2, row=9, padx=40, sticky="w")

caminho_exportacao = None
caminho_exportacao = customtkinter.CTkEntry(janela, placeholder_text=r"Ex: C:\IMPRESSÃO\ADV", width=250, height=40)
caminho_exportacao.grid(column=2, row=10, padx=40, sticky="w")

label5 = customtkinter.CTkLabel(janela, text="TIPO FORMATAÇÃO:", font=("Arial Bold", 10), anchor="w")
label5.grid(column=2, row=11, padx=40, sticky="w")


Array_velores_combo = ['CAMISA MASTER', 'CAMISA MASTER CURTA', 'CAMISA CAPUZ', 'CAMISA CASUAL CURTA', 'CAMISA CASUAL LONGA', 'CAMISA RAGLAN']
combo = customtkinter.CTkComboBox(janela, values= Array_velores_combo, font=("Arial Bold", 10), width=250)
combo.grid(column=2, row=12, padx=40, sticky="w")


#----------------------CAMPOS CHECKBOX-------------------------
check_box = customtkinter.CTkCheckBox(janela, text="MASC", onvalue="OK", offvalue=False)
check_box.grid(column=2, row=13, padx=40, pady=20, sticky="w")

check_box2 = customtkinter.CTkCheckBox(janela, text="FEM", onvalue="OK", offvalue=False)
check_box2.grid(column=2, row=13, padx=120, pady=20, sticky="w")

check_box3 = customtkinter.CTkCheckBox(janela, text="INF", onvalue="OK", offvalue=False)
check_box3.grid(column=2, row=13, padx=195, pady=20, sticky="w")

#----------------------CAPTURAS DADOS-------------------------

def dialogBox(titulo, mensagem):
    dialogo = customtkinter.CTkToplevel()
    dialogo.title(titulo)
    dialogo.geometry("300x125")
    
    #Define como janela temporária da janela principal
    dialogo.transient(janela)  
    dialogo.grab_set()       # Captura todos os eventos de entrada (modal)
    dialogo.focus_set()      # Define o foco na janela de alerta
    # Label para a mensagem
    
    label = customtkinter.CTkLabel(dialogo, text=mensagem, wraplength=250, font=("Arial Bold", 14))
    label.pack(padx= 10, pady=10)
    
    # Botão para fechar a janela
    botao_ok = customtkinter.CTkButton(dialogo, text="VERIFICAR", fg_color="#ed0c4c", width=250, height=35, hover_color="#c9003a", command=dialogo.destroy)
    botao_ok.pack(pady=20, padx=0, side="bottom")

def Main():
    def minimizar_janela():
        time.sleep(.3)  # Esperar 3 segundos
        janela.iconify() 
     
    codigo = campo_codigo.get()
    caminho_arquivos_ids = caminho_arquivos_impres.get()
    camino_exportacao_pdf = caminho_exportacao.get()
    caminho_layout = caminho_arquivos_layout.get()
    tipo_camisa = check_box.get()
    tipo_camisa2 = check_box2.get()
    tipo_camisa3 = check_box3.get()
    categoria_camisa = combo.get()
        
    if (not codigo) or (not caminho_arquivos_ids) or (not camino_exportacao_pdf) or (not caminho_layout) or ((tipo_camisa == False) and (tipo_camisa2 == False) and (tipo_camisa3 == False)):
        dialogBox("DADOS AUSENTES", "Por favor verifique as informações!")

    elif not codigo in camino_exportacao_pdf:
        dialogBox("CODIGO IMCOMPATÍVEL", "O código informado não coincide com código escrito na pasta!")
    
    else:    
        minimizar_janela() 
        def clicar():
            center_x = 1920 // 2
            center_y = 1080 // 2
            pyautogui.click(center_x, center_y)
        
        def MainFormat(ids_molde_name, tipo_formatacoes):

            def PegarAtalhos():
                atalhos_jason_ids = os.path.join(os.path.dirname(__file__),'atalhos.json')
                
                with open(atalhos_jason_ids) as file:
                    atalho = json.load(file)
                
                return atalho
                    
            Selecionar_atalho = PegarAtalhos()["Selecionar"]
            alinhar_atalho = PegarAtalhos()["Alinhar"]
            
            #INICIALIZAÇÃO
            try:
                app = win32com.client.Dispatch("InDesign.Application")
            except Exception as e:
                app = win32com.client.Dispatch("InDesign.Application.2024")
             
            if app != True:
                def CloseErrorWindow():
                    def encontrar_janela_por_titulo(title):
                        identificador_janela = win32gui.FindWindow(None, title) 
                        return identificador_janela #Retorna um inteiro se encontrar e 0 se não encontrar janela

                    # Fecha a janela do Creative Cloud
                    def close_creative_cloud():
                        identificador_janela = encontrar_janela_por_titulo("Creative Cloud") #quarada o retorno
                        if identificador_janela:
                            # Envia o comando para clicar no botão "Sair"
                            win32gui.PostMessage(identificador_janela, win32con.WM_CLOSE, 0, 0)
                            
                    # Fecha a janela do Creative Cloud
                    close_creative_cloud()
                    
                time.sleep(5)
                CloseErrorWindow()

                caminho = caminho_arquivos_ids+"\\"+ids_molde_name
                doc = app.Open(caminho)
                        
                #FUNÇÃO QUE COLOCA O CODIGO EM CADA MOLDE
                def inputCode(codigo_subs):
                    app.FindTextPreferences.FindWhat = 'CODIGO'
                    app.ChangeTextPreferences.ChangeTo = codigo_subs
                    app.ActiveDocument.ChangeText()
                    idNothing = 1851876449  
                    app.FindTextPreferences.FindWhat = idNothing
                    app.ChangeTextPreferences.ChangeTo = idNothing
                inputCode(codigo)

                #FUNÇÃO QUE ATUALIZA TODAS INSTÂNCIAS
                def RevinculaArquivoUsandoJavascript(arquivos_indesign, caminho_arquivos_png):
                    
                    # Código JavaScript formatado para substituir o caminho
                    javascript_code = f"""
                    
                    //DEFINI QUE O SCRIPT DEVE SER EXECUTADO DENTRO DO (INDESIGN)
                    #target indesign
                    
                    //ESSA FUNÇÃO VAI REVINCULAR TODOS OS ARQUIVOS (INDESIGN)
                    function relinkLinks(caminho_arquivos_png) {{
                        
                        var doc = app.activeDocument;
                        
                        //ISSO É UM ARRAY QUE COTÉM TODOS OS VINCULOS DO ARQUIVO (INDESIGN)
                        var links = doc.links;
                    
                        for (var i = 0; i < links.length; i++) {{
                            var link = links[i];
                            
                            //EU USEI O CONTRUCTOR (New) POIS, O MÉTODOS INCLUSIVE(relink) ESPERA UM OBJETO FILE, AO INVÉS DE UMA STRING NORMAL
                            var linkFile = new File(caminho_arquivos_png + "/" + link.name);
                            
                            //Verifica se o arquivo existe antes de relinkar
                            if (linkFile.exists) {{
                                link.relink(linkFile);
                                link.update();
                            }} else {{
                                //ESCREVE UMA MENSAGEM NA TELA DO (INDESIGN)
                                $.writeln("Arquivo não encontrado para relink: " + linkFile.fsName);
                            }}
                        }}
                    }}

                    relinkLinks("{caminho_arquivos_png.replace('\\', '/')}");
                    """

                    #O NÚMERO (1246973031) ESPECIFICA AO INDESIGN QUE QUERO TRABALHAR ESPECIFICAMENTE COM (JavasCript OU ExtendScript)
                    app.DoScript(javascript_code, 1246973031)

                RevinculaArquivoUsandoJavascript(caminho_arquivos_ids, caminho_layout)                                          
                if tipo_camisa != "GOLA":
                    time.sleep(3)
                    clicar()
                    def selecionar_prancheta(prancheta_numero):
                        try:
                            prancheta = doc.Pages(prancheta_numero)
                            prancheta.Select()
                            pyautogui.press('v')
                            pyautogui.press(Selecionar_atalho) #SELECIONAR TUDO
                            pyautogui.press(alinhar_atalho) #ALINHAR TUDO
                        except Exception as e:
                            print(f"Erro: {e}")
                    for pg in range(doc.Pages.Count+1):
                        selecionar_prancheta(pg)

                def exportComp(filename, pastlename):
                    def Exportar():
                        # Definindo o caminho de exportação
                        output_path = fr'{camino_exportacao_pdf}\{filename}.pdf'

                        idPDFType = 1952403524
                        # 1=[High Quality Print], 2=[PDF/X-1a:2001] etc..
                        myPDFPreset = app.PDFExportPresets.Item("PRESET_BRK")
                        doc.Export(idPDFType, output_path, False, myPDFPreset)
                    try:
                        Exportar()
                        def RenameFilesPdf():
                            for file in os.listdir(f"{camino_exportacao_pdf}\{filename}"):
                                string_slicing = file[:-6]+".pdf"
                                file_format = f"{camino_exportacao_pdf}\{filename}\\{string_slicing}"
                                file_complete = f"{camino_exportacao_pdf}\{filename}\\{file}"
                                os.rename(file_complete, file_format)
                        RenameFilesPdf()
                        
                        def RenameGola():      
                            for file in os.listdir(f"{camino_exportacao_pdf}\{filename}"):
                                if "GOLA" in file:
                                    if tipo_formatacoes == "MASC":
                                        gola_name = f"{codigo}_GOLA_MASTER.pdf"
                                        gola_file_format = f"{camino_exportacao_pdf}\{filename}\\{gola_name}"
                                        gola_file_complete = f"{camino_exportacao_pdf}\{filename}\\{file}"
                                        os.rename(gola_file_complete, gola_file_format)
                                    else:
                                        gola_name = f"{codigo}_GOLA_{tipo_formatacoes}.pdf"
                                        gola_file_format = f"{camino_exportacao_pdf}\{filename}\\{gola_name}"
                                        gola_file_complete = f"{camino_exportacao_pdf}\{filename}\\{file}"
                                        os.rename(gola_file_complete, gola_file_format)
                        if tipo_formatacoes == "MASC" or tipo_formatacoes == "FEM":
                            RenameGola()
                            
                    except Exception as e:
                        print(f"Ocorreu um erro durante a exportação: {str(e)}")
                    def RenamePasta():
                        caminho_pasta_completo = f"{camino_exportacao_pdf}\{filename}"
                        subs_pastlename = f"{camino_exportacao_pdf}\{pastlename}"
                        os.rename(caminho_pasta_completo,subs_pastlename)
                    RenamePasta()        
                
                match categoria_camisa:
                    case 'CAMISA MASTER':
                        if tipo_formatacoes == "FEM":    
                            exportComp(f"{codigo}_FULL_FEM", "FEM")
                        elif tipo_formatacoes == "INF":  
                            exportComp(f"{codigo}_FULL_INF", "INF")
                        else:
                            exportComp(f"{codigo}_FULL_MASC", "MASC")
                            
                    case 'CAMISA CAPUZ':
                        if tipo_formatacoes == "MASC":
                            exportComp(f"{codigo}_FULL_MASC", "MASC")
                            
                    case 'CAMISA MASTER CURTA':
                        if tipo_formatacoes == "FEM":    
                            exportComp(f"{codigo}_FULL_FEM", "FEM")
                        else:
                            exportComp(f"{codigo}_FULL_MASC", "MASC")
                            
                    case 'CAMISA CASUAL CURTA':
                        if tipo_formatacoes == "FEM":    
                            exportComp(f"{codigo}_FULL_FEM", "FEM")
                        elif tipo_formatacoes == "INF":  
                            exportComp(f"{codigo}_FULL_INF", "INF")
                        else:
                            exportComp(f"{codigo}_FULL_MASC", "MASC")
                            
                    case 'CAMISA CASUAL LONGA':
                        if tipo_formatacoes == "FEM":    
                            exportComp(f"{codigo}_FULL_FEM", "FEM")
                        elif tipo_formatacoes == "INF":  
                            exportComp(f"{codigo}_FULL_INF", "INF")
                        else:
                            exportComp(f"{codigo}_FULL_MASC", "MASC")

        def MainRunning():
            match categoria_camisa:
                case 'CAMISA MASTER':
                    if tipo_camisa == "OK":
                        MainFormat("FULL MASC.indd","MASC")
                    if tipo_camisa2 == "OK":
                        MainFormat("FULL FEM.indd","FEM")
                    if tipo_camisa3 == "OK":
                        MainFormat("FULL INF.indd","INF")
                        
                case 'CAMISA CAPUZ':
                    if tipo_camisa == "OK": 
                        MainFormat("CAPUZ FULL MASC.indd","MASC")
                        
                case 'CAMISA MASTER CURTA':
                    if tipo_camisa == "OK":
                        MainFormat("FULL MASC CURTA.indd","MASC")
                    if tipo_camisa2 == "OK":
                        MainFormat("FULL FEM CURTA.indd","FEM")
                        
                case 'CAMISA CASUAL CURTA':
                    if tipo_camisa == "OK":
                        MainFormat("CASUAL MASC CURTA.indd","MASC")
                    if tipo_camisa2 == "OK":
                        MainFormat("CASUAL FEM CURTA.indd","FEM")
                    if tipo_camisa3 == "OK":
                        MainFormat("CASUAL INF CURTA.indd","INF")
                
                case 'CAMISA CASUAL LONGA':   
                    if tipo_camisa == "OK":
                        MainFormat("CASUAL LONGA MASC.indd","MASC")
                    if tipo_camisa2 == "OK":
                        MainFormat("CASUAL LONGA FEM.indd","FEM")
                    if tipo_camisa3 == "OK":
                        MainFormat("CASUAL LONGA INF.indd","INF")
                    
        MainRunning()
    
#-----------------BOTÃO FORMATAÇÃO-------------------------
btn_formatadr = customtkinter.CTkButton(janela, text="FORMATAR", width=250, height=40, fg_color="#1071E5", command=Main)
btn_formatadr.grid(column=2, row=14, padx=40, sticky="w")

# Exibe a janela
os.system('cls')
janela.mainloop()
