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
FILE_NAME = "PROTHEUS_TODAS_TABELAS.xlsx"


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

    # Centraliza em relação à janela principal
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
          fg_color="#7F8C8D",
          hover_color="#62686F",
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
    self.geometry("480x500")  # Janela compactada na vertical
    self.resizable(False, False)

    self.cancelar_execucao = False

    self.carregar_configuracoes_iniciais()

    # Título da Aplicação
    self.label_titulo = ctk.CTkLabel(
        self,
        text="Exportar Tabelas do Protheus",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    self.label_titulo.pack(pady=(10, 5))

    # --- Seção de Configurações do Banco de Dados ---
    self.frame_db = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_db.pack(pady=(0, 5), padx=15, fill="x")

    self.label_db_titulo = ctk.CTkLabel(
        self.frame_db,
        text="Configurações de Conexão (SQL Server)",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.label_db_titulo.pack(anchor="w", pady=(0, 2))

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

      lbl = ctk.CTkLabel(row_frame, text=rotulo, width=105, anchor="w", font=ctk.CTkFont(size=11))
      lbl.pack(side="left")

      entry = ctk.CTkEntry(row_frame, width=335, height=24)
      if chave == "password":
        entry.configure(show="*")
      entry.pack(side="left", padx=2)

      self.entries[chave] = entry

    # --- Seção de Filtros e Aparência ---
    self.frame_filtros = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_filtros.pack(pady=(0, 5), padx=15, fill="x")

    self.label_filtros_titulo = ctk.CTkLabel(
        self.frame_filtros,
        text="Filtros e Aparência",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.label_filtros_titulo.pack(anchor="w", pady=(0, 2))

    # Campo para ignorar tabelas
    self.entry_ignorar = ctk.CTkEntry(
        self.frame_filtros,
        placeholder_text="Ignorar tabelas contendo (ex: 990; 010; LOG)",
        width=445,
        height=26,
    )
    self.entry_ignorar.pack(pady=(2, 4))

    # Campo para considerar somente prefixos
    self.entry_prefixos = ctk.CTkEntry(
        self.frame_filtros,
        placeholder_text="Somente prefixos (ex: SX; SB; SC - vazio para todas)",
        width=445,
        height=26,
    )
    self.entry_prefixos.pack(pady=(0, 4))

    # Linha horizontal compacta para Checkbox e Tema
    frame_opcoes = ctk.CTkFrame(self.frame_filtros, fg_color="transparent")
    frame_opcoes.pack(fill="x", pady=(0, 2))

    self.check_vazias = ctk.CTkCheckBox(
        frame_opcoes,
        text="Exportar vazias",
        font=ctk.CTkFont(size=11),
        checkbox_width=18,
        checkbox_height=18,
    )
    self.check_vazias.pack(side="left", padx=(2, 0))
    self.check_vazias.select()

    lbl_tema = ctk.CTkLabel(
        frame_opcoes, text="Tema:", font=ctk.CTkFont(size=11)
    )
    lbl_tema.pack(side="left", padx=(20, 5))

    self.option_tema = ctk.CTkOptionMenu(
        frame_opcoes,
        values=["Sistema", "Claro", "Escuro"],
        command=self.mudar_tema_interface,
        width=120,
        height=24,
    )
    self.option_tema.pack(side="left")

    # --- Frame para os botões principais ---
    self.frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_botoes.pack(pady=(4, 2))

    self.btn_iniciar = ctk.CTkButton(
        self.frame_botoes,
        text="Iniciar Exportação",
        command=self.iniciar_processo_thread,
        height=32,
        width=218,
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.btn_iniciar.pack(side="left", padx=4)

    self.btn_parar = ctk.CTkButton(
        self.frame_botoes,
        text="Interromper",
        command=self.solicitar_parada,
        height=32,
        width=218,
        fg_color="#C0392B",
        hover_color="#962D22",
        state="disabled",
        font=ctk.CTkFont(size=12, weight="bold"),
    )
    self.btn_parar.pack(side="left", padx=4)

    # --- Frame para botões secundários ---
    self.frame_secundarios = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_secundarios.pack(pady=(2, 5))

    self.btn_salvar = ctk.CTkButton(
        self.frame_secundarios,
        text="Salvar Parâmetros",
        command=self.acao_salvar_parametros,
        height=28,
        width=218,
        fg_color="#27AE60",
        hover_color="#1E8449",
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_salvar.pack(side="left", padx=4)

    self.btn_sobre = ctk.CTkButton(
        self.frame_secundarios,
        text="Sobre",
        command=self.mostrar_sobre,
        height=28,
        width=218,
        fg_color="#7F8C8D",
        hover_color="#62686F",
        font=ctk.CTkFont(size=11, weight="bold"),
    )
    self.btn_sobre.pack(side="left", padx=4)

    # Barra de Progresso
    self.progress_bar = ctk.CTkProgressBar(self, width=445, height=10)
    self.progress_bar.set(0)
    self.progress_bar.pack(pady=(2, 2))

    # Rótulo de Status
    self.label_status = ctk.CTkLabel(
        self, text="Pronto para iniciar.", font=ctk.CTkFont(size=11)
    )
    self.label_status.pack(pady=(0, 5))

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
      self.option_tema.set("Sistema")

  def salvar_configuracoes(self):
    config = configparser.ConfigParser()

    config["DATABASE"] = {}
    for chave, entry in self.entries.items():
      config["DATABASE"][chave] = entry.get().strip()

    config["FILTROS"] = {}
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
    if os.path.exists(FILE_NAME):
      CustomMessageBox.mostrar_mensagem(
          self,
          "Arquivo Existente",
          (
              f"O arquivo '{FILE_NAME}' já existe e será substituído.\nDeseja"
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
    return nome_limpo[:31]

  def abrir_planilha(self):
    try:
      if os.path.exists(FILE_NAME):
        if platform.system() == "Windows":
          os.startfile(FILE_NAME)
        elif platform.system() == "Darwin":
          subprocess.call(("open", FILE_NAME))
        else:
          subprocess.call(("xdg-open", FILE_NAME))
    except Exception as e:
      logging.error(f"Erro ao abrir a planilha: {e}")

  def export_all_tables_to_excel(self):
    registros_log = []

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

      with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        exportadas_count = 0
        for i, table_name in enumerate(tables, 1):
          if self.cancelar_execucao:
            self.label_status.configure(text="Processo interrompido pelo usuário.")
            registros_log.append({
                "tabela": table_name,
                "mensagem": "Processo interrompido pelo usuário.",
            })
            break

          progresso = i / total_tables
          self.progress_bar.set(progresso)
          self.label_status.configure(
              text=f"Processando ({i}/{total_tables}): {table_name}"
          )

          sheet_name = self.sanitizar_nome_aba(table_name)

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

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            exportadas_count += 1
            registros_log.append({
                "tabela": table_name,
                "mensagem": (
                    f"Sucesso: Exportada para a aba '{sheet_name}' com"
                    f" {len(df)} registro(s)."
                ),
            })
          except Exception as e:
            erro_msg = f"Erro ao processar a tabela: {e}"
            logging.error(f"Tabela '{table_name}': {erro_msg}")
            registros_log.append({"tabela": table_name, "mensagem": erro_msg})

        if registros_log:
          df_log = pd.DataFrame(registros_log)
          df_log.to_excel(writer, sheet_name="LOG", index=False)

      self.progress_bar.set(1.0)
      self.label_status.configure(text="Exportação concluída com sucesso!")

      mensagem_final = (
          f"Processo concluído!\nTabelas exportadas: {exportadas_count}\nArquivo"
          f" salvo como '{FILE_NAME}' (com aba 'LOG').\n\nDeseja abrir a"
          " planilha gerada agora?"
      )

      self.after(
          100,
          lambda: CustomMessageBox.mostrar_mensagem(
              self,
              "Sucesso",
              mensagem_final,
              tipo="pergunta",
              callback_sim=self.abrir_planilha,
          ),
      )

    except Exception as e:
      erro_fatal = f"Erro crítico: {e}"
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