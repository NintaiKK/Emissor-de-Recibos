# Gerador de Recibos Automático
Python 3.8+
License GNU GPLv3

Aplicativo desktop para geração automática de recibos profissionais em formato Word (.docx), desenvolvido para simplificar o processo de emissão de documentos financeiros.

## Visão Geral
Este projeto oferece uma solução completa para:

Emissão rápida de recibos padronizados

Cálculo automático de valores e descontos

Personalização de dados dos prestadores de serviço

Geração de documentos prontos para impressão ou envio digital

## Funcionalidades
✔ Modelos pré-configurados com dados de prestadores
✔ Cadastro de novos prestadores (opção "Outro")
✔ Cálculos automáticos:

Valor do serviço

Taxas adicionais

Descontos (DAS MEI)

Total líquido

## Formatação profissional:

Layout padronizado

Valores no formato monetário brasileiro

Data por extenso

Rodapé personalizável

## Gerenciamento de arquivos:

Nomes únicos para evitar sobrescrita

Opção de diretório padrão ou seleção manual

## Tecnologias Utilizadas
Python 3.8+ - Linguagem principal

python-docx - Geração de documentos Word

Tkinter - Interface gráfica

PyInstaller - Criação de executável (opcional)

## Instalação
Clone o repositório:
```
git clone https://github.com/seu-usuario/gerador-recibos.git
```
Instale as dependências:

```
pip install -r requirements.txt
```

# Como Usar
Execute o aplicativo:

```
python gerador_recibos.py
```

Preencha os dados:

Selecione o prestador de serviço

Insira valores quando necessário

Escolha o local para salvar

Clique em "Gerar Recibo"

# Personalização
Edite as variáveis no código para adaptar ao seu negócio:

```
# Dados da empresa
EMPRESA = "SUA EMPRESA AQUI"
CNPJ_EMPRESA = "00.000.000/0000-00"
CIDADE = "Sua Cidade"

# Valores padrão
VALOR_SERVICO = 150.00
VALOR_PASSAGEM = 20.00
DESCONTO_DAS = 80.90

# Prestadores cadastrados
funcionarios = {
    'PRESTADOR 1': {
        'cnpj': '00.000.000/0000-00',
        'servico': 'Descrição do serviço'
    }
}

# Diretório de saída
OUTPUT_FOLDER = r"C:\seu\diretorio\aqui"
```

# Dicas de Personalização
Imagem do rodapé:

Substitua rodape_img.png por sua logo/marca d'água

Dimensões recomendadas: largura 6.1 polegadas (≈15.5cm)

Estilo visual:

Altere a fonte em doc.styles['Normal'].font.name

Ajuste tamanhos com Pt()

Para converter em executável:
```
pyinstaller --onefile --windowed --icon=app.ico main.py
```
# Licença
Este projeto está licenciado sob a licença MIT - consulte o arquivo LICENSE para obter mais detalhes.

> Nota para Implementação: Substitua todos os campos marcados com ###SUBSTITUIR DADOS PELOS REAIS no código por suas informações reais antes de utilizar o sistema.
