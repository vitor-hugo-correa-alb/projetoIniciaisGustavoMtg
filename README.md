# Gerador Iniciais — Guia de Instalação e Execução (para distribuição interna)

Este documento descreve, de forma objetiva e profissional, como instalar e executar o aplicativo "Gerador Iniciais" nas máquinas dos colaboradores — tanto em Windows quanto em macOS. Siga as instruções do item adequado ao sistema operacional.

Visão geral
- Propósito: aplicação desktop para gerar documentos .docx a partir de modelos e seleção de "Modelos de Pedido".
- Estrutura esperada (quando distribuída):
  - Windows: forneça o executável gerado e a pasta `templates/` junto ao executável.
  - macOS: forneça o repositório (ou pacote) com a estrutura do projeto; o instalador cria um ambiente Python local e instala dependências.

Importante: não misturar arquivos de um build com outra plataforma. O build do macOS deve ser preparado em um Mac; o executável Windows (.exe) é gerado via PyInstaller em ambiente Windows.

1) Distribuição e execução — Windows (usuário final / funcionários)
Requisitos mínimos
- Windows 10 ou superior.
- O arquivo executável (por exemplo `GeradorIniciais.exe`) gerado pelo departamento de TI ou pela equipe de desenvolvimento.
- A pasta `templates/` contendo:
  - `modelo_base.docx`
  - `modelo_base_final.docx` (opcional)
  - `modelos/` (subpasta com os .docx dos pedidos)

Como preparar o pacote a ser entregue
- Coloque o executável e a pasta `templates/` no mesmo diretório. Estrutura recomendada:
  ```
  GeradorIniciais.exe
  templates/
    modelo_base.docx
    modelo_base_final.docx
    modelos/
      pedido1.docx
      pedido2.docx
      ...
  ```
- Inclua também o diretório `logs/` vazio se desejar (o aplicativo criará `logs/logs.txt` automaticamente se não existir).

Como executar (usuário)
- Abra o explorador de arquivos e dê duplo‑clique em `GeradorIniciais.exe`.
- Alternativamente, execute via Prompt/PowerShell (útil para capturar mensagens de erro):
  ```powershell
  cd C:\caminho\para\app
  .\GeradorIniciais.exe
  ```
Onde encontrar logs
- O aplicativo cria/atualiza `logs/logs.txt` no mesmo diretório do executável. Peça esse arquivo ao usuário caso seja necessário investigar erros.

Observações para o departamento de TI
- Se o app não iniciar por dupla-clique em algumas máquinas, peça ao usuário para executar pelo terminal para visualizar mensagens ou libere o executável via políticas de segurança.
- Para distribuir por GPO/instalador interno, cuide de copiar também a pasta `templates/` para o diretório de instalação.

2) Instalação e execução — macOS (procedimento para TI / administrador)
Observação importante
- O aplicatvo para macOS não é gerado a partir do Windows. Para funcionar corretamente, a instalação (criação do venv e instalação das dependências) deve ser feita em um Mac.

Pré-requisitos no macOS
- Python 3.10+ instalado (recomendável usar o instalador oficial em https://www.python.org ou Homebrew).
- Xcode Command Line Tools (para compilar dependências nativas): execute
  ```bash
  xcode-select --install
  ```
- (Opcional / recomendado) Homebrew para instalação de bibliotecas do sistema:
  ```bash
  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
  brew install libxml2 libxslt pkg-config zlib libjpeg
  ```

Arquivos incluídos para facilitar a instalação
- `mac/requirements.txt` — lista de dependências Python do projeto.
- `mac/installer.sh` — script que cria um virtualenv em `.venv_mac` na raiz do projeto e instala as dependências.
- `run_mac.sh` — script que ativa o virtualenv criado e executa o programa.

Passo a passo (para o técnico)
1. Copie todo o repositório para a máquina Mac, mantendo a estrutura:
   ```
   project-root/
     src/
     templates/
     mac/
       installer.sh
       requirements.txt
     run_mac.sh
     ...
   ```
2. Torne os scripts executáveis:
   ```bash
   chmod +x mac/installer.sh
   chmod +x run_mac.sh
   ```
3. Execute o instalador (este passo cria `.venv_mac` na raiz do projeto):
   ```bash
   ./mac/installer.sh
   ```
   - Se a instalação falhar por causa de `lxml`/`Pillow`, siga as instruções exibidas pelo script:
     - Instalar Xcode CLT: `xcode-select --install`
     - Instalar bibliotecas via Homebrew (conforme mostrado no pré‑requisitos)
     - Exportar flags de compilação (se necessário) e reinstalar requisitos:
       ```bash
       export LDFLAGS="-L/opt/homebrew/opt/libxml2/lib -L/opt/homebrew/opt/libxslt/lib"
       export CPPFLAGS="-I/opt/homebrew/opt/libxml2/include -I/opt/homebrew/opt/libxslt/include"
       pip install -r mac/requirements.txt
       ```
4. Execute o aplicativo:
   - Script runner (recomendado):
     ```bash
     ./run_mac.sh
     ```
   - Ou manualmente (caso precise debugar):
     ```bash
     source .venv_mac/bin/activate
     python -m src.main
     ```

Onde encontrar logs
- O aplicativo grava `logs/logs.txt` na raiz do projeto (mesmo local onde está `run_mac.sh`). Solicite este arquivo para diagnóstico.

3) Templates e arquivos necessários
- O funcionamento correto depende da presença de modelos na pasta `templates/`. Verifique:
  - `templates/modelo_base.docx` (arquivo principal)
  - `templates/modelo_base_final.docx` (opcional; se existir, será anexado no final com numeração sequencial)
  - `templates/modelos/*.docx` (cada modelo de pedido)
- Atenção: nomes de arquivos não podem conter o caractere `'/'`. Caso precise representar uma barra visual no título, utilize o token `{{BARRA}}` no nome do arquivo; o aplicativo converterá `{{BARRA}}` para `/` ao inserir o nome no documento e na interface.

4) Segurança e permissões (Caso seja computador Empresarial com políticas rígidas)
- No macOS, aplicativos não assinados aparecerão como "bloqueados" pelo Gatekeeper — abra o app com botão direito → Abrir, ou providencie assinatura/notarização via Apple Developer pela equipe de TI.
- Em Windows, se houver políticas de segurança rígidas, valide a entrega do exe via mecanismos internos seguros (assinatura de código, distribuição por GPO, etc.).

5) Troubleshooting (erros comuns)
- Erro ao abrir o exe sem mensagens: execute via terminal/PowerShell para obter o traceback.
- Falha na instalação de `lxml`/`Pillow` no mac: instale Xcode CLT e as bibliotecas nativas via Homebrew (ver seção macOS).
- Módulo `docxcompose` ausente no exe: verifique se o executável foi gerado com o mesmo Python/venv usado para desenvolvimento e se os hooks do PyInstaller (pasta `hooks/`) foram incluídos.
- Arquivos de template não localizados: verifique a localização da pasta `templates/` em relação ao executável (Windows) ou à raiz do projeto (mac).

6) Contato / suporte
- Se encontrar problemas, ao reportar envie:
  - Descrição do problema e passos para reproduzir.
  - Conteúdo do arquivo de logs: `logs/logs.txt`.
  - Versão do sistema operacional (Windows ou macOS) e arquitetura (Intel / Apple Silicon).
  - Se aplicável, o traceback completo do erro exibido no console.
- Para suporte interno: encaminhe ao responsável pelo projeto (equipe de TI / desenvolvedor) com as informações acima.

7) Versão e rastreabilidade
- Inclua junto ao pacote a versão do software (por ex. nome do arquivo `GeradorIniciais_v1.0.exe` ou um arquivo `VERSION`) para controle de deployment.



[//]: # (Fim do documento — pontos de ação recomendados)
[//]: # (- Para distribuição ampla: equipe de TI deve preparar um instalador/unidade de distribuição &#40;Windows: instalador MSI / copiar exe + templates; macOS: criar app com PyInstaller e assinar&#41;.)
[//]: # (- Se desejar, a equipe de desenvolvimento pode fornecer builds assinados &#40;Windows code signing / macOS Developer ID&#41; e/ou imagens de instalação &#40;MSI / DMG&#41;.)
