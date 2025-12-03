# Gerador Iniciais — Guia de Instalação e Execução (versão para distribuição interna)

Este documento descreve, de forma clara e objetiva, como distribuir e executar a aplicação "Gerador Iniciais" entre colaboradores da empresa.

---  
Índice rápido
- Windows: executar o .exe fornecido.
- macOS: executar dois scripts (instalador + runner) que criam um virtualenv e iniciam a aplicação.
- Estrutura mínima esperada (sempre fornecer templates junto ao pacote).

---

## 1 — Requisitos e estrutura de distribuição

Requisitos mínimos (para ambos os sistemas)
- Fornecer sempre a pasta `templates/` com:
  - `modelo_base.docx`
  - `modelo_base_final.docx`
  - `modelos/` (subpasta com .docx dos pedidos)
- Ter instalado Word (Microsoft 365, Office 2019 ou superior) para abrir os documentos gerados.
- O programa grava logs em `logs/logs.txt` no diretório base do pacote.

Estrutura recomendada quando for distribuir para um usuário:
```
package/
  GeradorIniciais.exe        (Windows build)   <-- somente para Windows
  src/                       (código-fonte quando for instalação em Mac)
  templates/
    modelo_base.docx
    modelo_base_final.docx
    modelos/
      pedido1.docx
      ...
  mac/
    installer.sh
    requirements.txt
  run_mac.sh
  README.md
```

---

## 2 — Instalação e execução (Windows — usuário final)

Requisitos
- Windows 10 ou superior.
- O executável `GeradorIniciais.exe` fornecido pelo departamento de TI.

Passos (usuário)
1. Coloque `GeradorIniciais.exe` e a pasta `templates/` na mesma pasta.
2. Execute o aplicativo por duplo‑clique ou via Prompt/PowerShell:
   ```powershell
   cd C:\caminho\para\package
   .\GeradorIniciais.exe
   ```

Logs
- Consulte `logs/logs.txt` no mesmo diretório para diagnóstico.

---

## 3 — Instalação e execução (macOS — administrador / TI)

Visão geral
- No macOS a distribuição consiste em fornecer os scripts `mac/installer.sh` e `run_mac.sh` junto com o código e a pasta `templates/`.  
- O script `installer.sh` cria um virtualenv `.venv_mac` e instala as dependências listadas em `mac/requirements.txt`.  
- O script `run_mac.sh` ativa esse virtualenv e executa a aplicação.

Pré‑requisitos no Mac (sugerido)
- Python 3.10+ (usar instalador oficial ou `brew install python`).
- Xcode Command Line Tools (necessário para compilar algumas dependências nativas):
  ```bash
  xcode-select --install
  ```
- Homebrew (opcional, recomendado quando lxml/Pillow exigirem libs do sistema):
  ```bash
  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
  brew install libxml2 libxslt pkg-config zlib libjpeg
  ```

Passo a passo (TI / administrador)
1. Copie o pacote para o Mac, preservando a estrutura (veja seção 1).
2. Torne os scripts executáveis:
   ```bash
   chmod +x mac/installer.sh
   chmod +x run_mac.sh
   ```
3. Execute o instalador (cria `.venv_mac` e instala dependências):
   ```bash
   ./mac/installer.sh
   ```
   - Se a instalação falhar por `lxml`/`Pillow`, seguir as instruções apresentadas pelo script (instalar Xcode CLT, dependências via Homebrew e exportar flags de compilação quando necessário).


4. Execute a aplicação:
   ```bash
   ./run_mac.sh
   ```
   - Alternativamente (para depuração):
     ```bash
     source .venv_mac/bin/activate
     python -m src.main
     ```

Logs
- O aplicativo grava `logs/logs.txt` na raiz do projeto (mesmo local onde está `run_mac.sh`).

Distribuição para usuários finais macOS
- TI pode fornecer um arquivo compactado (.zip) contendo o pacote (src/, templates/, mac/ e scripts).  
- Instruir o usuário técnico local (ou TI) a executar `mac/installer.sh` uma vez por máquina e `run_mac.sh` para iniciar.

Importante
- Não será fornecido um .app assinado; usuários veriam bloqueio do Gatekeeper se tentassem executar apps não assinados. A alternativa aprovada pela equipe é: executar via scripts (conforme acima) ou TI criar um .app e assinar/notarizar em ambiente Apple se desejar um fluxo sem intervenção do usuário.

---

## 4 — Tokens especiais e nomes de arquivos

- Nomes de arquivos não podem conter o caractere `/`. Para exibir uma barra no título do documento, use o token `{{BARRA}}` no nome do arquivo (ex.: `Pedido{{BARRA}}A.docx`); o sistema converte `{{BARRA}}` em `/` ao montar os títulos no documento e na interface.
- Evite caracteres especiais não suportados pelo sistema de arquivos do SO destino.

---

## 5 — Troubleshooting (resumo rápido)

Erro ao instalar dependências no mac
- Verifique Python e pip: `python3 --version` e `python3 -m pip --version`.
- Instale Xcode CLT: `xcode-select --install`.
- Instale libs via Homebrew e exporte flags se pip necessitar compilar:
  ```bash
  brew install libxml2 libxslt pkg-config zlib libjpeg
  export LDFLAGS="-L/opt/homebrew/opt/libxml2/lib -L/opt/homebrew/opt/libxslt/lib"
  export CPPFLAGS="-I/opt/homebrew/opt/libxml2/include -I/opt/homebrew/opt/libxslt/include"
  pip install -r mac/requirements.txt
  ```

Aplicativo Windows não inicia ao duplo clique
- Execute via Prompt para obter traceback:
  ```powershell
  cd C:\caminho\para\package
  .\GeradorIniciais.exe
  ```
- Forneça `logs/logs.txt` para diagnóstico.

Erros de import no exe (docxcompose, lxml)
- Verifique que o exe foi gerado no mesmo Python/venv onde as dependências estavam instaladas; use hooks do PyInstaller para pacotes problemáticos.

---