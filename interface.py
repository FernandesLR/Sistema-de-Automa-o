import customtkinter as ct
from validar import validacoes as vali
from Automatização import iniciar

def confirma(vp, vf, vl):
    cor = '#ff006a' # Vermelho
    cor2 = '#00ff37' # Verde

    # Verifica se o valor inserido pelo usuário está correto e retorna mensagem de erro
    if not vali.planilha(vp):
        valor_planilha.configure(border_color=cor, border_width=2)
        text_erro_planilha = ct.CTkLabel(frame2, text='Formato errado ou o arquivo não existe', text_color=cor)
        text_erro_planilha.place(x=610, y=105) # Posiciona ao lado do campo de entrada
        frame2.after(3000, lambda: text_erro_planilha.destroy())
    else:
        valor_planilha.configure(border_color=cor2, border_width=2) # Muda a cor para verde

    # Retorna uma mensagem de erro
    if not vali.folha(vp, vf):
        valor_folha.configure(border_color=cor, border_width=2)
        text_erro_folha = ct.CTkLabel(frame2, text='Folha não existe na planilha', text_color=cor)
        text_erro_folha.place(x=610, y=215) # Posiciona ao lado do campo de entrada
        frame2.after(2000, lambda: text_erro_folha.destroy())
    else:
        valor_folha.configure(border_color=cor2, border_width=2) # Muda a cor para verde
    
    # Retorna mensagem em caso
    if not vali.linhas(vl):
        valor_qtd.configure(border_color=cor, border_width=2)
        text_erro_qtd = ct.CTkLabel(frame2, text='Valor inserido deve ser um número positivo', text_color=cor)
        text_erro_qtd.place(x=610, y=325) # Posiciona ao lado do campo de entrada
        frame2.after(2000, lambda: text_erro_qtd.destroy())
    else:
        valor_qtd.configure(border_color=cor2, border_width=2) # Muda a cor para verde

    # Verifica se todos os valores estão corretos e finaliza o app
    if vali.planilha(vp) and vali.folha(vp, vf) and vali.linhas(vl):
        iniciar(vp, vf, vl)
        app.after(2000, lambda: app.quit())

app = ct.CTk()
app.title('Cadastro de dados')
app.geometry('800x600')
app._set_appearance_mode('light')

# Criando um frame para agrupar conteúdo
frame = ct.CTkFrame(master=app, fg_color='#0043b0', height=100)
frame.pack(fill='x')

# Criando o título
text = ct.CTkLabel(frame, text='Cadastro de Dados', font=('Arial', 30), text_color='white', height=100)
text.pack(pady=20)

# Criando segundo frame
frame2 = ct.CTkFrame(master=app, border_color='black', border_width=2, fg_color='white')
frame2.pack(pady=40, padx=20, fill='both', expand=True)

# Nome da planilha
text = ct.CTkLabel(frame2, text='Nome da planilha:', font=('Arial', 15), text_color='black')
text.place(x=20, y=100)
valor_planilha = ct.CTkEntry(frame2, placeholder_text='Nome da planilha', height=35, width=400, text_color='black', fg_color='white')
valor_planilha.place(x=200, y=100)

# Nome da folha
text = ct.CTkLabel(frame2, text='Nome da folha:', font=('Arial', 15), text_color='black')
text.place(x=20, y=210)
valor_folha = ct.CTkEntry(frame2, placeholder_text='Nome da folha', height=35, width=400, text_color='black', fg_color='white')
valor_folha.place(x=200, y=210)

# Quantidade de linhas
text = ct.CTkLabel(frame2, text='Quantidade de linhas:', font=('Arial', 15), text_color='black')
text.place(x=20, y=320)
valor_qtd = ct.CTkEntry(frame2, placeholder_text='0', height=35, width=400, text_color='black', fg_color='white')
valor_qtd.place(x=200, y=320)

# Botão para começar
btn = ct.CTkButton(frame2, text='Começar', width=100, command=lambda: confirma(valor_planilha.get(), valor_folha.get(), valor_qtd.get()))
btn.place(x=350, y=400)

# Encerra o app
app.mainloop()
