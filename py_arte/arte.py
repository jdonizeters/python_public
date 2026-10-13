import random
import time
import tkinter as tk
from tkinter import messagebox

# Configurações da Tela (100 linhas x 232 colunas - correspondente até a coluna HL)
LIN_MAX = 100
COL_MAX = 232
TOTAL_PONTOS = LIN_MAX * COL_MAX

# Tamanho de cada pixel na tela gráfica (em pixels de vídeo)
PIXEL_SIZE = 5 
LARGURA_TELA = COL_MAX * PIXEL_SIZE
ALTURA_TELA = LIN_MAX * PIXEL_SIZE

def sortear_cor_por_opcao(opcao):
    """
    Retorna diretamente a cor em formato hexadecimal (#RRGGBB) com base na escolha atual,
    com alta variação e canais RGB garantidos na ordem correta.
    """
    if opcao == 0:  # Original (VBA / Cores totalmente aleatórias)
        r = random.randint(0, 255)
        g = random.randint(0, 255)
        b = random.randint(0, 255)
    elif opcao == 1:   # Vermelho
        r = random.randint(150, 255)
        g = random.randint(0, 50)
        b = random.randint(0, 50)
    elif opcao == 2: # Laranja
        r = random.randint(200, 255)
        g = random.randint(80, 160)
        b = random.randint(0, 30)
    elif opcao == 3: # Amarelo
        r = random.randint(200, 255)
        g = random.randint(200, 255)
        b = random.randint(0, 50)
    elif opcao == 4: # Verde
        r = random.randint(0, 50)
        g = random.randint(150, 255)
        b = random.randint(0, 50)
    elif opcao == 5: # Azul
        r = random.randint(0, 50)
        g = random.randint(50, 150)
        b = random.randint(180, 255)
    elif opcao == 6: # Anil / Azul Escuro
        r = random.randint(10, 60)
        g = random.randint(10, 80)
        b = random.randint(130, 220)
    elif opcao == 7: # Violeta / Roxo
        r = random.randint(120, 220)
        g = random.randint(0, 60)
        b = random.randint(180, 255)
        
    return f"#{r:02x}{g:02x}{b:02x}"

class ObraDeArteApp:
    def __init__(self, root):
        self.root = root
        # Título atualizado conforme solicitado
        self.root.title("Poeira, Nuves e Árvores")
        self.root.geometry(f"{LARGURA_TELA + 40}x{ALTURA_TELA + 180}")
        self.root.configure(bg="#222222")

        # Rótulo de Status / Cronômetro
        self.status_label = tk.Label(root, text="Selecione uma cor e clique em 'Iniciar Arte'", fg="white", bg="#222222", font=("Arial", 11, "bold"))
        self.status_label.pack(pady=6)

        # Canvas onde os pixels serão pintados
        self.canvas = tk.Canvas(root, width=LARGURA_TELA, height=ALTURA_TELA, bg="white", highlightthickness=0)
        self.canvas.pack()

        # --- PAINEL DE BOTÕES DE CORES ---
        paleta_frame = tk.LabelFrame(root, text=" Paletas de Cores (Mude a qualquer momento) ", fg="yellow", bg="#222222", font=("Arial", 9, "bold"))
        paleta_frame.pack(pady=6, fill="x", padx=10)

        self.opcao_cor_escolhida = tk.IntVar(value=0)
        self.botoes_cores = {}

        cores_botoes = [
            ("Original", 0, "#cccccc", "black"),
            ("Vermelho", 1, "#ff4444", "white"),
            ("Laranja", 2, "#ff8800", "white"),
            ("Amarelo", 3, "#ffcc00", "black"),
            ("Verde", 4, "#00cc44", "white"),
            ("Azul", 5, "#3399ff", "white"),
            ("Anil", 6, "#3333ff", "white"),
            ("Violeta", 7, "#9933ff", "white")
        ]

        cores_inner_frame = tk.Frame(paleta_frame, bg="#222222")
        cores_inner_frame.pack(pady=4)

        for texto, val, bg_col, fg_col in cores_botoes:
            b = tk.Button(cores_inner_frame, text=texto, command=lambda v=val: self.mudar_cor_tempo_real(v),
                          bg=bg_col, fg=fg_col, font=("Arial", 9, "bold"), relief=tk.RAISED, padx=6, pady=2)
            b.pack(side=tk.LEFT, padx=3)
            self.botoes_cores[val] = b

        self.atualizar_estilo_botoes_cor()

        # --- PAINEL DE CONTROLE DE EXECUÇÃO ---
        controle_frame = tk.Frame(root, bg="#222222")
        controle_frame.pack(pady=6)

        self.btn_iniciar = tk.Button(controle_frame, text="Iniciar Obra de Arte", command=self.iniciar_processo, font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", padx=10, pady=5)
        self.btn_iniciar.pack(side=tk.LEFT, padx=5)

        self.btn_parar = tk.Button(controle_frame, text="Parar Processo", command=self.parar_processo, font=("Arial", 10, "bold"), bg="#ff9800", fg="white", padx=10, pady=5, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=5)

        # Botão Sobre criado com as informações solicitadas
        self.btn_sobre = tk.Button(controle_frame, text="Sobre", command=self.mostrar_sobre, font=("Arial", 10, "bold"), bg="#2196F3", fg="white", padx=10, pady=5)
        self.btn_sobre.pack(side=tk.LEFT, padx=5)

        self.btn_sair = tk.Button(controle_frame, text="Sair", command=root.quit, font=("Arial", 10, "bold"), bg="#f44336", fg="white", padx=10, pady=5)
        self.btn_sair.pack(side=tk.LEFT, padx=5)

        # Inicialização da matriz na memória
        self.matriz = [[None for _ in range(COL_MAX)] for _ in range(LIN_MAX)]
        self.pixel_ids = [[None for _ in range(COL_MAX)] for _ in range(LIN_MAX)]
        self.inicializar_grid_visual()
        
        self.executando = False

    def inicializar_grid_visual(self):
        """Cria os pixels em branco na tela gráfica para otimizar a performance"""
        for r in range(LIN_MAX):
            for c in range(COL_MAX):
                x1 = c * PIXEL_SIZE
                y1 = r * PIXEL_SIZE
                x2 = x1 + PIXEL_SIZE
                y2 = y1 + PIXEL_SIZE
                p_id = self.canvas.create_rectangle(x1, y1, x2, y2, outline="", fill="white")
                self.pixel_ids[r][c] = p_id

    def mudar_cor_tempo_real(self, nova_opcao):
        """Altera a paleta ativamente enquanto o programa roda ou antes de iniciar"""
        self.opcao_cor_escolhida.set(nova_opcao)
        self.atualizar_estilo_botoes_cor()

    def atualizar_estilo_botoes_cor(self):
        """Destaca visualmente qual botão de cor está selecionado no momento"""
        ativo = self.opcao_cor_escolhida.get()
        for val, btn in self.botoes_cores.items():
            if val == ativo:
                btn.config(relief=tk.SUNKEN, bd=3)
            else:
                btn.config(relief=tk.RAISED, bd=1)

    def mostrar_sobre(self):
        """Apresenta a janela com as informações de autoria"""
        mensagem = (
            "Criado por: Donizete - jdonizeters@gmail.com - 30/01/20 - V1.06\n\n"
            "Original em VBA, em python e com melhorias com ajuda da IA"
        )
        messagebox.showinfo("Sobre - Poeira e Nuvens", mensagem)

    def iniciar_processo(self):
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_parar.config(state=tk.NORMAL)
        
        # Reinicializa variáveis de estado
        self.n_pontos = 0
        self.n_arvores = 0
        self.n_nuvens = 0
        self.cima = 0
        self.baixo = 0
        self.nE = 0
        self.nD = 0
        self.l_acelera = False
        self.executando = True
        
        # Marca o tempo de início para o cronômetro
        self.tempo_inicio = time.time()
        
        # Limpa matriz interna e pinta tela de branco
        for r in range(LIN_MAX):
            for c in range(COL_MAX):
                self.matriz[r][c] = None
                self.canvas.itemconfig(self.pixel_ids[r][c], fill="white")

        # Dispara o laço principal
        self.root.after(10, self.passo_geracao)

    def parar_processo(self):
        """Interrompe o loop de geração imediatamente"""
        self.executando = False
        self.btn_iniciar.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED)
        self.status_label.config(text="Processo interrompido pelo usuário.")

    def passo_geracao(self):
        """Executa um ciclo do laço processando um evento por vez com troca de cor dinâmica"""
        
        if not self.executando:
            return

        if self.n_pontos >= TOTAL_PONTOS:
            tempo_total = time.time() - self.tempo_inicio
            minutos = int(tempo_total // 60)
            segundos = int(tempo_total % 60)
            tempo_str = f"{minutos}m {segundos}s" if minutos > 0 else f"{segundos} segundos"
            
            self.status_label.config(text=f"Concluído em {tempo_str}! Árvores: {self.n_arvores:,} | Nuvens: {self.n_nuvens:,}")
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_parar.config(state=tk.DISABLED)
            self.executando = False
            return

        if not self.l_acelera and self.n_pontos > (TOTAL_PONTOS / 2):
            self.l_acelera = True

        # Sorteia a cor com base na paleta selecionada no momento exato da iteração
        hex_cor = sortear_cor_por_opcao(self.opcao_cor_escolhida.get())
        
        n_lin = random.randint(0, LIN_MAX - 1)
        n_col = random.randint(0, COL_MAX - 1)
        n_qual = random.randint(1, 2)       # 1 = Árvore (vertical), 2 = Nuvem (horizontal)
        n_direcao = random.randint(1, 2)

        if self.matriz[n_lin][n_col] is None:
            self.n_pontos += 1
            self.matriz[n_lin][n_col] = "."
            self.canvas.itemconfig(self.pixel_ids[n_lin][n_col], fill=hex_cor)
        else:
            n_passo = random.randint(1, 10)
            
            if n_qual == 1: # Árvore (Vertical)
                self.n_arvores += 1
                if n_direcao == 1: # Baixo
                    self.baixo += 1
                    inicio = LIN_MAX - 1 if self.l_acelera else n_lin
                    cond = (lambda l2: l2 >= 0) if self.l_acelera else (lambda l2: l2 < LIN_MAX)
                    passo_loop = -1 if self.l_acelera else n_passo
                    
                    n_lin2 = inicio
                    while cond(n_lin2):
                        if self.matriz[n_lin2][n_col] != ".":
                            self.n_pontos += 1
                            self.matriz[n_lin2][n_col] = "."
                            self.canvas.itemconfig(self.pixel_ids[n_lin2][n_col], fill=hex_cor)
                        n_lin2 += passo_loop
                else: # Cima
                    self.cima += 1
                    inicio = 0 if self.l_acelera else n_lin
                    cond = (lambda l2: l2 < LIN_MAX) if self.l_acelera else (lambda l2: l2 >= 0)
                    passo_loop = 1 if self.l_acelera else -n_passo
                    
                    n_lin2 = inicio
                    while cond(n_lin2):
                        if self.matriz[n_lin2][n_col] != ".":
                            self.n_pontos += 1
                            self.matriz[n_lin2][n_col] = "."
                            self.canvas.itemconfig(self.pixel_ids[n_lin2][n_col], fill=hex_cor)
                        n_lin2 += passo_loop
            
            else: # Nuvem (Horizontal)
                self.n_nuvens += 1
                if n_direcao == 1: # Esquerda para Direita
                    self.nE += 1
                    inicio = 0 if self.l_acelera else n_col
                    passo_loop = 1 if self.l_acelera else n_passo
                    
                    n_col2 = inicio
                    while n_col2 < COL_MAX:
                        if self.matriz[n_lin][n_col2] != ".":
                            self.n_pontos += 1
                            self.matriz[n_lin][n_col2] = "."
                            self.canvas.itemconfig(self.pixel_ids[n_lin][n_col2], fill=hex_cor)
                        if self.l_acelera:
                            break
                        n_col2 += passo_loop
                else: # Direita para Esquerda
                    self.nD += 1
                    inicio = 0 if self.l_acelera else n_col
                    passo_loop = 1 if self.l_acelera else n_passo
                    
                    n_col2 = inicio
                    while n_col2 < COL_MAX:
                        idx_col = max(0, min(COL_MAX - 1, n_col2))
                        if self.matriz[n_lin][idx_col] != ".":
                            self.n_pontos += 1
                            self.matriz[n_lin][idx_col] = "."
                            self.canvas.itemconfig(self.pixel_ids[n_lin][idx_col], fill=hex_cor)
                        if not self.l_acelera and (n_col2 - n_passo < 0):
                            break
                        n_col2 += passo_loop

        # Cálculo do Tempo Decorrido para o Cronômetro
        tempo_decorrido = int(time.time() - self.tempo_inicio)
        min_dec = tempo_decorrido // 60
        seg_dec = tempo_decorrido % 60
        cronometro_str = f"{min_dec:02d}:{seg_dec:02d}"

        # Atualiza a barra de status
        percentual = int((self.n_pontos / TOTAL_PONTOS) * 100)
        status_texto = (f"Tempo: {cronometro_str} | Progresso: {percentual}% | "
                        f"Pontos: {self.n_pontos:,} | Árvores: {self.n_arvores:,} | Nuvens: {self.n_nuvens:,}")
        
        self.status_label.config(text=status_texto)

        # Ritmo gradual suave mantido (fator 24)
        progresso_ratio = self.n_pontos / TOTAL_PONTOS
        tempo_pausa = int(24 * (1.0 - progresso_ratio))

        self.root.after(max(0, tempo_pausa), self.passo_geracao)

if __name__ == "__main__":
    root = tk.Tk()
    app = ObraDeArteApp(root)
    root.mainloop()