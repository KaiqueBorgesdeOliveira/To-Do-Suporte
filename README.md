# 📋 Controle de Chamados - Suporte

Sistema de controle de chamados para equipes de suporte técnico, desenvolvido em Python com interface gráfica moderna.

## 🚀 Funcionalidades

- ✅ **Cadastro de Chamados**: Registre chamados com número, tipo de item e quantidade
- 📊 **Visualização Organizada**: Tabela dinâmica que agrupa itens por tipo
- 📤 **Exportação CSV**: Exporte os dados para planilhas com formatação otimizada
- 🔄 **Reset de Dados**: Limpe todos os registros quando necessário
- 🎨 **Interface Moderna**: Design escuro e responsivo
- 💾 **Banco de Dados**: Armazenamento local com SQLite

## 📦 Instalação

### Pré-requisitos
- Python 3.7 ou superior
- Bibliotecas: `tkinter`, `sqlite3`, `csv`

### Instalação das Dependências
```bash
pip install -r requirements.txt
```

### Executando o Programa

#### Opção 1: Executar o Script Python
```bash
python ToDo.py
```

#### Opção 2: Usar o Executável (Recomendado)
1. Execute o comando para gerar o executável:
```bash
python -m PyInstaller --noconsole --onefile ToDo.py
```

2. Acesse a pasta `dist` e execute o arquivo `ToDo.exe`

## 🎯 Como Usar

### 1. Adicionando um Novo Item
1. **Número do Chamado**: Digite o número (já inicia com "SC-")
2. **Tipo do Item**: Selecione na lista predefinida
3. **Quantidade**: Digite a quantidade (já inicia com "1")
4. Clique em **"Adicionar Item"**

### 2. Visualizando os Dados
- A tabela mostra todos os itens agrupados por tipo
- A última linha exibe os totais por categoria
- Os dados são atualizados automaticamente

### 3. Exportando Dados
- Clique em **"Exportar CSV"**
- Escolha o local para salvar o arquivo
- O arquivo será salvo com codificação UTF-8 e separador ponto e vírgula

### 4. Resetando os Dados
- Clique em **"Resetar Planilha"**
- Confirme a ação na caixa de diálogo
- ⚠️ **Atenção**: Esta ação apaga todos os dados permanentemente

## 📋 Tipos de Itens Disponíveis

### Periféricos
- Mouse Dell, Logitech, Multilazer
- Teclado Dell, Multilazer
- Headset Jabra, Logitech

### Energia e Carregamento
- Pilhas AA, AAA
- Carregadores de notebook (Dell Type C, Lenovo Type C, Dell 3420, Dell 7490)
- Carregador de MacBook

### Acessórios
- Dockstation (Prata, Preta)
- Cabos (Rede, Energia, HDMI)

### Monitores
- Monitor Dell 24 polegadas
- Monitor Lenovo 23 polegadas

## 🗄️ Estrutura do Banco de Dados

O sistema utiliza SQLite com a seguinte estrutura:

```sql
CREATE TABLE chamados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_chamado TEXT NOT NULL,
    tipo_item TEXT NOT NULL,
    quantidade INTEGER NOT NULL
)
```

## 🎨 Personalização

### Cores do Tema
- **Fundo Principal**: `#1e1e2d`
- **Fundo da Tabela**: `#3f3f4d`
- **Texto**: `#ffffff`
- **Campos de Entrada**: `#2e2e3d`

### Dimensões da Janela
- **Largura**: 450px
- **Altura**: 500px
- **Centralizada**: Automaticamente na tela

## 🔧 Desenvolvimento

### Estrutura do Projeto
```
To-Do/
├── ToDo.py          # Arquivo principal
├── chamados.db      # Banco de dados SQLite
├── dist/            # Pasta com executável
│   └── ToDo.exe
└── README.md        # Este arquivo
```

### Tecnologias Utilizadas
- **Python 3.x**: Linguagem principal
- **Tkinter**: Interface gráfica
- **SQLite3**: Banco de dados
- **CSV**: Exportação de dados
- **PyInstaller**: Geração de executável

## 🐛 Solução de Problemas

### Erro ao Executar o Executável
1. Mova o arquivo para uma pasta fora do OneDrive
2. Clique com botão direito → Propriedades → Desbloquear
3. Execute como administrador
4. Adicione exceção no antivírus/Windows Defender

### Problemas com Exportação CSV
- O arquivo é salvo com codificação UTF-8-SIG
- Separador: ponto e vírgula (;)
- Compatível com Excel e LibreOffice

## 📝 Changelog

### v1.0.0
- ✅ Interface gráfica moderna com tema escuro
- ✅ Cadastro de chamados com validação
- ✅ Tabela dinâmica com agrupamento
- ✅ Exportação CSV otimizada
- ✅ Reset de dados com confirmação
- ✅ Executável sem console
- ✅ Valores padrão nos campos (SC- e quantidade 1)

## 👥 Contribuição

Para contribuir com o projeto:
1. Faça um fork do repositório
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 📞 Suporte

Para dúvidas ou sugestões, entre em contato através dos issues do repositório.

---

**Desenvolvido com ❤️ para equipes de suporte técnico**
