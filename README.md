# Análise de Consistência — Labor Rural

Software desktop para análise automática de inconsistências em planilhas MIMC (Mais Inteligência, Mais Cacau).

---

## Funcionalidades

| Aba | Verificações |
|-----|-------------|
| **INVENTARIO** | Valores < R$100 ou > R$500k; data de fabricação posterior à de aquisição |
| **PRODUCAO** | Rateio com talhões de produção faltando; valores divergentes entre talhões |
| **DESPESAS** | Atividades agrícolas sem M.O.; manutenção de máquinas como Administração; R$/ha > R$5.000; valor unitário > R$5.000 ou < R$1,00; Kg/Litros com unitário > R$200; rateio incompleto; lançamentos idênticos duplicados; lacunas de recorrência administrativa |
| **VENDAS** | Preço de venda acima de R$100/sc |

---

## Download do executável

Acesse a aba **[Releases](../../releases)** deste repositório e baixe o arquivo:

```
Analise_Consistencia_LaborRural.exe
```

Nenhuma instalação é necessária. Execute direto no Windows.

---

## Passo a passo — Configurar o repositório e publicar o .exe

### Pré-requisitos
- Conta no [GitHub](https://github.com) (gratuita)
- Git instalado no computador: https://git-scm.com/download/win

---

### 1. Criar o repositório no GitHub

1. Acesse https://github.com e faça login
2. Clique em **New repository** (botão verde no canto superior direito)
3. Preencha:
   - **Repository name**: `analise-consistencia-labor-rural`
   - **Visibility**: Public (para usar GitHub Actions gratuitamente) ou Private
   - Marque **Add a README file**: NÃO (já temos o nosso)
4. Clique em **Create repository**
5. Copie a URL do repositório (formato: `https://github.com/SEU_USUARIO/analise-consistencia-labor-rural.git`)

---

### 2. Enviar os arquivos para o GitHub

Abra o **Prompt de Comando** ou **PowerShell** na pasta do projeto e execute:

```bash
# Inicializar o repositório local
git init

# Adicionar todos os arquivos
git add .

# Fazer o primeiro commit
git commit -m "Versão inicial — Análise de Consistência Labor Rural"

# Conectar ao repositório remoto (substitua pela sua URL)
git remote add origin https://github.com/SEU_USUARIO/analise-consistencia-labor-rural.git

# Enviar para o GitHub
git branch -M main
git push -u origin main
```

---

### 3. Gerar o primeiro .exe (criar uma Release)

O GitHub Actions vai compilar o `.exe` automaticamente quando você criar uma **tag de versão**:

```bash
# Criar e enviar a tag de versão
git tag v1.0.0
git push origin v1.0.0
```

Depois disso:
1. Acesse seu repositório no GitHub
2. Clique na aba **Actions** — você verá o processo de build rodando
3. Aguarde entre **3 a 8 minutos**
4. Quando concluir, acesse **Releases** (menu lateral direito)
5. O arquivo `Analise_Consistencia_LaborRural.exe` estará disponível para download

---

### 4. Publicar uma nova versão (após atualizações)

Sempre que alterar o código e quiser publicar uma nova versão:

```bash
# Adicionar as mudanças
git add .
git commit -m "Descrição das alterações"
git push origin main

# Criar nova tag
git tag v1.1.0
git push origin v1.1.0
```

O GitHub Actions vai gerar um novo `.exe` automaticamente.

---

### 5. Compilar manualmente (opcional)

Se quiser gerar o `.exe` localmente:

```bash
# Instalar Python 3.11+ (https://python.org)
# Depois, no terminal:

pip install openpyxl pandas pyinstaller
pyinstaller build.spec
```

O executável será gerado em `dist/Analise_Consistencia_LaborRural.exe`.

---

## Estrutura do projeto

```
analise-consistencia-labor-rural/
├── main.py                          # Código principal
├── build.spec                       # Configuração do PyInstaller
├── requirements.txt                 # Dependências Python
├── README.md                        # Este arquivo
└── .github/
    └── workflows/
        └── build.yml                # GitHub Actions — build automático
```

---

## Tecnologias utilizadas

- **Python 3.11+**
- **openpyxl** — leitura de planilhas `.xlsx`
- **pandas** — manipulação de dados
- **tkinter** — interface gráfica nativa
- **PyInstaller** — empacotamento em `.exe`
- **GitHub Actions** — build e release automáticos

---

## Labor Rural

Consultoria técnica especializada em cacau e café.  
Desenvolvido para uso interno na plataforma MIMC.
