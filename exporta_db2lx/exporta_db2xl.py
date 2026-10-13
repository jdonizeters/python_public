import configparser
import logging
import os
import platform
import re
import subprocess
import threading
import customtkinter as ctk
import pandas as pd
from sqlalchemy import create_engine, text

# --- Configurações de Log ---
LOG_FILE = "erros_importacao.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.ERROR,
    format="%(asctime)s - [ERRO] - %(message)s",
    encoding="utf-8",
)

# --- Arquivo de Configuração INI ---
CONFIG_FILE = "config.ini"


class CustomMessageBox:
  """Classe utilitária para criar caixas de diálogo personalizadas em CTkToplevel."""

  @staticmethod
  def mostrar_mensagem(
      parent,
      titulo,
      mensagem,
      tipo="info",
      callback_sim=None,
      callback_nao=None,
  ):
    top = ctk.CTkToplevel(parent)
    top.title(titulo)
    top.geometry("400x200")
    top.resizable(False, False)
    top.grab_set()

    top.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() // 2) - (400 // 2)
    y = parent.winfo_y() + (parent.winfo_height() // 2) - (200 // 2)
    top.geometry(f"+{x}+{y}")

    lbl = ctk.CTkLabel(
        top,
        text=mensagem,
        font=ctk.CTkFont(size=13),
        justify="center",
        wraplength=360,
    )
    lbl.pack(pady=(30, 20), padx=20)

    frame_botoes = ctk.CTkFrame(top, fg_color="transparent")
    frame_botoes.pack(pady=10)

    if tipo == "pergunta":
      btn_sim = ctk.CTkButton(
          frame_botoes,
          text="Sim",
          width=100,
          command=lambda: (top.destroy(), callback_sim and callback_sim()),
      )
      btn_sim.pack(side="left", padx=10)

      btn_nao = ctk.CTkButton(
          frame_botoes,
          text="Não",
          width=100,
          command=lambda: (top.destroy(), callback_nao and callback_nao()),
      )
      btn_nao.pack(side="left", padx=10)
    else:
      btn_ok = ctk.CTkButton(
          frame_botoes,
          text="OK",
          width=120,
          command=lambda: (top.destroy(), callback_sim and callback_sim()),
      )
      btn_ok.pack()


class AppExportacao(ctk.CTk):

  def __init__(self):
    super().__init__()

    self.title("Exportador Protheus para Excel")
    self.geometry("480x610")
    self.resizable(False, False)

    self.cancelar_execucao = False

    self.carregar_configuracoes_iniciais()

    self.label_titulo = ctk.CTkLabel(
        self,
        text="Exportar Tabelas do Protheus",
        font=ctk.CTkFont(size=15, weight="bold"),
    )
    self.label_titulo.pack(pady=(8, 4))

    # --- Seção de Configurações do Banco de Dados ---
    self.frame_db = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_db.pack(pady=(0, 4), padx=15, fill="x")

    self.label_db_titulo = ctk.CTkLabel(
        self.frame_db,
        text="Configurações de Conexão (SQL Server)",
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.label_db_titulo.pack(anchor="w", pady=(0, 1))

    self.entries = {}
    campos = [
        ("server", "Servidor:", "localhost"),
        ("database", "Banco de Dados:", "protheus2510_padrao"),
        ("username", "Usuário:", "sa"),
        ("password", "Senha:", "sa"),
        ("driver", "Driver ODBC:", "SQL Server Native Client 11.0"),
    ]

    for chave, rotulo, padrao in campos:
      row_frame = ctk.CTkFrame(self.frame_db, fg_color="transparent")
      row_frame.pack(fill="x", pady=1)

      lbl = ctk.CTkLabel(
          row_frame, text=rotulo, width=100, anchor="w", font=ctk.CTkFont(size=10)
      )
      lbl.pack(side="left")

      entry = ctk.CTkEntry(row_frame, width=340, height=22)
      if chave == "password":
        entry.configure(show="*")
      entry.pack(side="left", padx=2)

      self.entries[chave] = entry

    # --- Seção de Saída e Filtros com Labels Individuais ---
    self.frame_filtros = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_filtros.pack(pady=(0, 4), padx=15, fill="x")

    self.label_filtros_titulo = ctk.CTkLabel(
        self.frame_filtros,
        text="Saída, Filtros e Aparência",
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.label_filtros_titulo.pack(anchor="w", pady=(0, 1))

    # 1. Arquivo de Saída
    lbl_arquivo = ctk.CTkLabel(
        self.frame_filtros,
        text="Nome do arquivo de saída:",
        font=ctk.CTkFont(size=10),
    )
    lbl_arquivo.pack(anchor="w", pady=(1, 0))

    self.entry_arquivo = ctk.CTkEntry(
        self.frame_filtros,
        placeholder_text="ex: PROTHEUS.xlsx",
        width=440,
        height=22,
    )
    self.entry_arquivo.pack(pady=(0, 2))

    # 2. Ignorar Tabelas
    lbl_ignorar = ctk.CTkLabel(
        self.frame_filtros,
        text="Ignorar tabelas contendo (separado por ';'):",
        font=ctk.CTkFont(size=10),
    )
    lbl_ignorar.pack(anchor="w", pady=(1, 0))

    self.entry_ignorar = ctk.CTkEntry(
        self.frame_filtros,
        placeholder_text="ex: 990; 010; LOG",
        width=440,
        height=22,
    )
    self.entry_ignorar.pack(pady=(0, 2))

    # 3. Filtrar por Prefixos
    lbl_prefixos = ctk.CTkLabel(
        self.frame_filtros,
        text="Somente tabelas com prefixos (separado por ';'):",
        font=ctk.CTkFont(size=10),
    )
    lbl_prefixos.pack(anchor="w", pady=(1, 0))

    self.entry_prefixos = ctk.CTkEntry(
        self.frame_filtros,
        placeholder_text="ex: SX; SB; SC (vazio para todas)",
        width=440,
        height=22,
    )
    self.entry_prefixos.pack(pady=(0, 2))

    # Linha horizontal compacta para Checkbox e Tema
    frame_opcoes = ctk.CTkFrame(self.frame_filtros, fg_color="transparent")
    frame_opcoes.pack(fill="x", pady=(1, 1))

    self.check_vazias = ctk.CTkCheckBox(
        frame_opcoes,
        text="Exportar vazias",
        font=ctk.CTkFont(size=10),
        checkbox_width=16,
        checkbox_height=16,
    )
    self.check_vazias.pack(side="left", padx=(2, 0))
    self.check_vazias.select()

    lbl_tema = ctk.CTkLabel(
        frame_opcoes, text="Tema:", font=ctk.CTkFont(size=10)
    )
    lbl_tema.pack(side="left", padx=(25, 5))

    self.option_tema = ctk.CTkOptionMenu(
        frame_opcoes,
        values=["Sistema", "Claro", "Escuro"],
        command=self.mudar_tema_interface,
        width=110,
        height=22,
    )
    self.option_tema.pack(side="left")

    # --- Botões Organizados em Duas Colunas (2 por linha) ---
    LARGURA_BTN = 215
    ALTURA_BTN = 28

    frame_botoes_1 = ctk.CTkFrame(self, fg_color="transparent")
    frame_botoes_1.pack(pady=2)

    self.btn_iniciar = ctk.CTkButton(
        frame_botoes_1,
        text="Iniciar Exportação",
        command=self.iniciar_processo_thread,
        height=ALTURA_BTN,
        width=LARGURA_BTN,
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_iniciar.pack(side="left", padx=5)

    self.btn_parar = ctk.CTkButton(
        frame_botoes_1,
        text="Interromper",
        command=self.solicitar_parada,
        height=ALTURA_BTN,
        width=LARGURA_BTN,
        state="disabled",
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_parar.pack(side="left", padx=5)

    frame_botoes_2 = ctk.CTkFrame(self, fg_color="transparent")
    frame_botoes_2.pack(pady=2)

    self.btn_salvar = ctk.CTkButton(
        frame_botoes_2,
        text="Salvar Parâmetros",
        command=self.acao_salvar_parametros,
        height=ALTURA_BTN,
        width=LARGURA_BTN,
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_salvar.pack(side="left", padx=5)

    self.btn_sobre = ctk.CTkButton(
        frame_botoes_2,
        text="Sobre",
        command=self.mostrar_sobre,
        height=ALTURA_BTN,
        width=LARGURA_BTN,
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_sobre.pack(side="left", padx=5)

    frame_botoes_3 = ctk.CTkFrame(self, fg_color="transparent")
    frame_botoes_3.pack(pady=2)

    self.btn_sair = ctk.CTkButton(
        frame_botoes_3,
        text="Sair do Aplicativo",
        command=self.destroy,
        height=ALTURA_BTN,
        width=440,
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_sair.pack(padx=5)

    self.progress_bar = ctk.CTkProgressBar(self, width=440, height=8)
    self.progress_bar.set(0)
    self.progress_bar.pack(pady=(4, 2))

    self.label_status = ctk.CTkLabel(
        self, text="Pronto para iniciar.", font=ctk.CTkFont(size=10)
    )
    self.label_status.pack(pady=(0, 4))

    self.carregar_valores_campos()

  def carregar_configuracoes_iniciais(self):
    config = configparser.ConfigParser()
    ctk.set_default_color_theme("blue")

    if os.path.exists(CONFIG_FILE):
      config.read(CONFIG_FILE)
      if "FILTROS" in config and "tema" in config["FILTROS"]:
        tema = config["FILTROS"]["tema"]
        if tema == "Claro":
          ctk.set_appearance_mode("Light")
        elif tema == "Escuro":
          ctk.set_appearance_mode("Dark")
        else:
          ctk.set_appearance_mode("System")
      else:
        ctk.set_appearance_mode("System")
    else:
      ctk.set_appearance_mode("System")

  def carregar_valores_campos(self):
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
      config.read(CONFIG_FILE)

      if "DATABASE" in config:
        for chave in self.entries:
          if chave in config["DATABASE"]:
            self.entries[chave].delete(0, "end")
            self.entries[chave].insert(0, config["DATABASE"][chave])

      if "FILTROS" in config:
        if "arquivo_saida" in config["FILTROS"]:
          self.entry_arquivo.delete(0, "end")
          self.entry_arquivo.insert(0, config["FILTROS"]["arquivo_saida"])
        if "ignorar" in config["FILTROS"]:
          self.entry_ignorar.delete(0, "end")
          self.entry_ignorar.insert(0, config["FILTROS"]["ignorar"])
        if "prefixos" in config["FILTROS"]:
          self.entry_prefixos.delete(0, "end")
          self.entry_prefixos.insert(0, config["FILTROS"]["prefixos"])
        if "exportar_vazias" in config["FILTROS"]:
          if config["FILTROS"].getboolean("exportar_vazias"):
            self.check_vazias.select()
          else:
            self.check_vazias.deselect()
        if "tema" in config["FILTROS"]:
          self.option_tema.set(config["FILTROS"]["tema"])
        else:
          self.option_tema.set("Sistema")
    else:
      defaults = {
          "server": "localhost",
          "database": "protheus2510_padrao",
          "username": "sa",
          "password": "sa",
          "driver": "SQL Server Native Client 11.0",
      }
      for chave, valor in defaults.items():
        self.entries[chave].delete(0, "end")
        self.entries[chave].insert(0, valor)
      self.entry_arquivo.insert(0, "PROTHEUS_TODAS_TABELAS.xlsx")
      self.option_tema.set("Sistema")

  def salvar_configuracoes(self):
    config = configparser.ConfigParser()

    config["DATABASE"] = {}
    for chave, entry in self.entries.items():
      config["DATABASE"][chave] = entry.get().strip()

    config["FILTROS"] = {}
    config["FILTROS"]["arquivo_saida"] = (
        self.entry_arquivo.get().strip() or "PROTHEUS_TODAS_TABELAS.xlsx"
    )
    config["FILTROS"]["ignorar"] = self.entry_ignorar.get().strip()
    config["FILTROS"]["prefixos"] = self.entry_prefixos.get().strip()
    config["FILTROS"]["exportar_vazias"] = str(self.check_vazias.get())
    config["FILTROS"]["tema"] = self.option_tema.get()

    with open(CONFIG_FILE, "w", encoding="utf-8") as configfile:
      config.write(configfile)

  def mudar_tema_interface(self, escolha):
    if escolha == "Claro":
      ctk.set_appearance_mode("Light")
    elif escolha == "Escuro":
      ctk.set_appearance_mode("Dark")
    else:
      ctk.set_appearance_mode("System")

  def acao_salvar_parametros(self):
    try:
      self.salvar_configuracoes()
      CustomMessageBox.mostrar_mensagem(
          self,
          "Sucesso",
          "Parâmetros salvos com sucesso no arquivo 'config.ini'!",
          tipo="info",
      )
    except Exception as e:
      CustomMessageBox.mostrar_mensagem(
          self, "Erro", f"Não foi possível salvar os parâmetros: {e}", tipo="info"
      )

  def mostrar_sobre(self):
    janela_sobre = ctk.CTkToplevel(self)
    janela_sobre.title("Sobre")
    janela_sobre.geometry("380x220")
    janela_sobre.resizable(False, False)
    janela_sobre.grab_set()

    janela_sobre.update_idletasks()
    x = self.winfo_x() + (self.winfo_width() // 2) - (380 // 2)
    y = self.winfo_y() + (self.winfo_height() // 2) - (220 // 2)
    janela_sobre.geometry(f"+{x}+{y}")

    lbl_titulo_sobre = ctk.CTkLabel(
        janela_sobre,
        text="Exportador de Tabelas",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    lbl_titulo_sobre.pack(pady=(20, 10))

    texto_info = (
        "Exportador de Tabelas do SQL para Planilha.\n"
        "Versão 1.05\n"
        "Criado por jdrs"
    )
    lbl_texto = ctk.CTkLabel(
        janela_sobre,
        text=texto_info,
        font=ctk.CTkFont(size=13),
        justify="center",
    )
    lbl_texto.pack(pady=(0, 20))

    btn_fechar = ctk.CTkButton(
        janela_sobre,
        text="Fechar",
        command=janela_sobre.destroy,
        width=120,
        height=32,
    )
    btn_fechar.pack()

  def iniciar_processo_thread(self):
    file_name = self.entry_arquivo.get().strip() or "PROTHEUS_TODAS_TABELAS.xlsx"
    if os.path.exists(file_name):
      CustomMessageBox.mostrar_mensagem(
          self,
          "Arquivo Existente",
          (
              f"O arquivo '{file_name}' já existe e será substituído.\nDeseja"
              " continuar?"
          ),
          tipo="pergunta",
          callback_sim=self.executar_fluxo_exportacao,
      )
    else:
      self.executar_fluxo_exportacao()

  def executar_fluxo_exportacao(self):
    self.salvar_configuracoes()
    self.cancelar_execucao = False
    self.btn_iniciar.configure(state="disabled")
    self.btn_parar.configure(state="normal")
    self.progress_bar.set(0)
    threading.Thread(target=self.export_all_tables_to_excel, daemon=True).start()

  def solicitar_parada(self):
    self.cancelar_execucao = True
    self.label_status.configure(
        text="Interrompendo... Aguarde a tabela atual finalizar."
    )
    self.btn_parar.configure(state="disabled")

  def get_db_engine(self):
    db_config = {chave: entry.get().strip() for chave, entry in self.entries.items()}
    conn_str = (
        f"mssql+pyodbc://{db_config['username']}:{db_config['password']}"
        f"@{db_config['server']}/{db_config['database']}"
        f"?driver={db_config['driver'].replace(' ', '+')}&TrustServerCertificate=yes"
    )
    return create_engine(conn_str)

  def get_all_tables(self, engine, ignorar_lista, prefixos_lista):
    query = text(
        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE ="
        " 'BASE TABLE' ORDER BY TABLE_NAME ASC"
    )
    with engine.connect() as connection:
      df_tables = pd.read_sql(query, connection)

    tables = df_tables["TABLE_NAME"].tolist()
    tables_filtradas = []

    for tabela in tables:
      if prefixos_lista:
        atende_prefixo = False
        for pref in prefixos_lista:
          if pref and tabela.startswith(pref):
            atende_prefixo = True
            break
        if not atende_prefixo:
          continue

      deve_ignorar = False
      if ignorar_lista:
        for termo in ignorar_lista:
          if termo and termo in tabela:
            deve_ignorar = True
            break
      if deve_ignorar:
        continue

      tables_filtradas.append(tabela)

    return tables_filtradas

  def sanitizar_nome_aba(self, nome_tabela):
    nome_limpo = re.sub(r'[:\\/\?\*\[\]]', "_", nome_tabela)
    if nome_limpo.upper() in ["SRA", "LOG", "HISTORY"]:
      nome_limpo = f"T_{nome_limpo}"
    return nome_limpo[:31]

  def limpar_dados_dataframe(self, df):
    """Converte colunas problemáticas (como dicts/json) para string para evitar crashes no Excel."""
    for col in df.columns:
      df[col] = df[col].apply(
          lambda x: str(x) if isinstance(x, (dict, list)) else x
      )
    return df

  def abrir_planilha(self, file_name):
    try:
      if os.path.exists(file_name):
        if platform.system() == "Windows":
          os.startfile(file_name)
        elif platform.system() == "Darwin":
          subprocess.call(("open", file_name))
        else:
          subprocess.call(("xdg-open", file_name))
    except Exception as e:
      logging.error(f"Erro ao abrir a planilha: {e}")

  def export_all_tables_to_excel(self):
    registros_log = []
    file_name = self.entry_arquivo.get().strip() or "PROTHEUS_TODAS_TABELAS.xlsx"

    try:
      texto_ignorar = self.entry_ignorar.get().strip()
      ignorar_lista = (
          [t.strip() for t in texto_ignorar.split(";")]
          if texto_ignorar
          else []
      )

      texto_prefixos = self.entry_prefixos.get().strip()
      prefixos_lista = (
          [p.strip() for p in texto_prefixos.split(";")]
          if texto_prefixos
          else []
      )

      exportar_vazias = self.check_vazias.get()

      self.label_status.configure(text="Conectando ao banco de dados...")
      engine = self.get_db_engine()
      tables = self.get_all_tables(engine, ignorar_lista, prefixos_lista)

      total_tables = len(tables)
      if total_tables == 0:
        self.label_status.configure(
            text="Nenhuma tabela encontrada após os filtros."
        )
        self.restaurar_botoes()
        return

      dados_tabelas = {}
      for i, table_name in enumerate(tables, 1):
        if self.cancelar_execucao:
          self.label_status.configure(text="Processo interrompido pelo usuário.")
          registros_log.append({
              "tabela": table_name,
              "mensagem": "Processo interrompido pelo usuário.",
          })
          break

        progresso = (i / total_tables) * 0.5
        self.progress_bar.set(progresso)
        self.label_status.configure(
            text=f"Lendo do Banco ({i}/{total_tables}): {table_name}"
        )

        try:
          query = f"SELECT * FROM [{table_name}]"
          with engine.connect() as connection:
            df = pd.read_sql(query, connection)

          if len(df) == 0 and not exportar_vazias:
            registros_log.append({
                "tabela": table_name,
                "mensagem": (
                    "Ignorada: tabela vazia (zero registros) conforme"
                    " configuração."
                ),
            })
            continue

          # Limpa dados complexos para evitar falhas de gravação
          df = self.limpar_dados_dataframe(df)

          dados_tabelas[table_name] = df
          registros_log.append({
              "tabela": table_name,
              "mensagem": f"Sucesso: Lida com {len(df)} registro(s).",
          })
        except Exception as e:
          erro_msg = f"Erro ao processar/ler a tabela: {e}"
          logging.error(f"Tabela '{table_name}': {erro_msg}")
          registros_log.append({"tabela": table_name, "mensagem": erro_msg})

      if dados_tabelas and not self.cancelar_execucao:
        try:
          with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
            total_gravacao = len(dados_tabelas)
            tabelas_sucesso = 0

            for idx, (table_name, df) in enumerate(dados_tabelas.items(), 1):
              progresso = 0.5 + ((idx / total_gravacao) * 0.5)
              self.progress_bar.set(progresso)
              self.label_status.configure(
                  text=f"Gravando no Excel ({idx}/{total_gravacao}): {table_name}"
              )

              sheet_name = self.sanitizar_nome_aba(table_name)
              try:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                tabelas_sucesso += 1
              except Exception as e_aba:
                erro_aba = f"Erro ao gravar aba '{sheet_name}': {e_aba}"
                logging.error(f"Tabela '{table_name}': {erro_aba}")
                # Atualiza a mensagem na lista de log para registrar o erro de gravação
                for reg in registros_log:
                  if reg["tabela"] == table_name:
                    reg["mensagem"] = erro_aba

            if registros_log:
              df_log = pd.DataFrame(registros_log)
              df_log.to_excel(writer, sheet_name="LOG", index=False)

          self.progress_bar.set(1.0)
          self.label_status.configure(text="Exportação concluída com sucesso!")

          mensagem_final = (
              f"Processo concluído!\nTabelas gravadas com sucesso:"
              f" {tabelas_sucesso}\nArquivo salvo como '{file_name}' (com aba"
              " 'LOG').\n\nDeseja abrir a planilha gerada agora?"
          )

          self.after(
              100,
              lambda: CustomMessageBox.mostrar_mensagem(
                  self,
                  "Sucesso",
                  mensagem_final,
                  tipo="pergunta",
                  callback_sim=lambda: self.abrir_planilha(file_name),
              ),
          )

        except Exception as e_excel:
          erro_gravacao = f"Erro crítico ao gravar o arquivo Excel: {e_excel}"
          logging.error(erro_gravacao)
          self.label_status.configure(text="Erro ao gravar arquivo!")
          self.after(
              100,
              lambda: CustomMessageBox.mostrar_mensagem(
                  self, "Erro Crítico", erro_gravacao, tipo="info"
              ),
          )
      else:
        if not self.cancelar_execucao:
          self.label_status.configure(text="Nenhuma tabela válida para gravar.")
          self.after(
              100,
              lambda: CustomMessageBox.mostrar_mensagem(
                  self,
                  "Atenção",
                  (
                      "Nenhuma tabela foi lida com sucesso para gerar a"
                      " planilha."
                  ),
                  tipo="info",
              ),
          )

    except Exception as e:
      erro_fatal = f"Erro crítico geral: {e}"
      logging.error(erro_fatal)
      self.label_status.configure(text="Erro no processo!")
      self.after(
          100,
          lambda: CustomMessageBox.mostrar_mensagem(
              self, "Erro Crítico", erro_fatal, tipo="info"
          ),
      )

    finally:
      self.restaurar_botoes()

  def restaurar_botoes(self):
    self.btn_iniciar.configure(state="normal")
    self.btn_parar.configure(state="disabled")


if __name__ == "__main__":
  app = AppExportacao()
  app.mainloop()